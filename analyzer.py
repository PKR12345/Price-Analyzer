import streamlit as st
import tempfile
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.colors import ColorChoice
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Font
import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

def filter(data, filters, num_filters, key_column):
    filters_dict = {}
    data_dict = {}
    meta_dict = {}
    for i in range(num_filters):
        st.subheader(f"Select Filter {i + 1}")
        filters_dict[i] = {}
        options_dict = {}
        for j in ["Target_Brand", "Competitor_Brand"]:
            options_dict[j] = {}
            for k in filters:
                option = st.selectbox(f"{j}- Select the Filter {i+1} for column {k}", data[k].unique())
                options_dict[j][k] = option
            query = " and ".join([f"`{k}` == '{options_dict[j][k]}'" for k in options_dict[j].keys()])
            data_subset = data.query(query)
            data_subset = data_subset.groupby("Week End Date").agg({"Vol Share": "sum", "Average Unit Price": "mean"}).reset_index()
            data_subset.columns = ["Week End Date", f"Vol_Share_{j}", f"Average_Unit_Price_{j}"]
            if not data_subset.empty:
                filters_dict[i][j] = data_subset
        if not filters_dict[i]["Target_Brand"].empty and not filters_dict[i]["Competitor_Brand"].empty:
            data_merged = pd.merge(filters_dict[i]["Target_Brand"], filters_dict[i]["Competitor_Brand"], on="Week End Date", how="left")
            data_merged["Target_Brand Price ix"] = (data_merged["Average_Unit_Price_Target_Brand"]/data_merged["Average_Unit_Price_Competitor_Brand"])*100
            data_merged["Competitor_Brand Price ix"] = (data_merged["Average_Unit_Price_Competitor_Brand"]/data_merged["Average_Unit_Price_Target_Brand"])*100
            data_merged["Vol_Share_Total"] = data_merged["Vol_Share_Target_Brand"] + data_merged["Vol_Share_Competitor_Brand"]
            key = str(options_dict["Target_Brand"]["Brand"][:3])+"_"+str(options_dict["Target_Brand"][key_column])+"_"+str(options_dict["Competitor_Brand"]["Brand"][:3])+"_"+str(options_dict["Competitor_Brand"][key_column])
            data_dict[key] = data_merged
            meta_dict[key] = options_dict
    return data_dict, meta_dict

@st.cache_resource
def view_data(data_dict, meta_dict):
    for key in data_dict.keys():
        st.markdown(f"<h3>{key}</h3>", unsafe_allow_html=True)
        st.write("Target Brand Filters:")
        for k in meta_dict[key]["Target_Brand"].keys():
            st.write(f"{k} : {meta_dict[key]['Target_Brand'][k]}")
        st.write("Competitor Brand Filters:")
        for k in meta_dict[key]["Competitor_Brand"].keys():
            st.write(f"{k} : {meta_dict[key]['Competitor_Brand'][k]}")
        st.write("Data:")
        st.dataframe(data_dict[key])
    
@st.cache_resource        
def _add_data_worksheet(data_dict, meta_dict):
    wb = Workbook()
    for sheetname, df in data_dict.items():
        ws = wb.create_sheet(title=sheetname)
        ws.append(df.columns.tolist())
        for row in df.itertuples(index=False, name=None):
            ws.append(row)
        ws["N1"].value = "Metadata:"
        ws["N1"].font = openpyxl.styles.Font(bold=True)
        ws["N1"].alignment = openpyxl.styles.Alignment(horizontal="center")
        ws["N1"].fill = openpyxl.styles.PatternFill("solid", fgColor="00b0f0")
        ws["N3"].value = "Target Brand:"
        ws["N3"].font = openpyxl.styles.Font(bold=True)
        ws["N3"].alignment = openpyxl.styles.Alignment(horizontal="center")
        ws["P3"].value = "Competitor Brand:"
        ws["P3"].font = openpyxl.styles.Font(bold=True)
        ws["P3"].alignment = openpyxl.styles.Alignment(horizontal="center")
        p = 4
        for k in meta_dict[sheetname]["Target_Brand"].keys():
            ws[f"N{p}"].value = k
            ws[f"O{p}"].value = meta_dict[sheetname]["Target_Brand"][k]
            ws[f"P{p}"].value = k
            ws[f"Q{p}"].value = meta_dict[sheetname]["Competitor_Brand"][k]
            p+=1
            
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    return wb

def generate_charts(wb, x_col_index, y_col_index_1, y_col_index_2, config_dict, header, header_index, start_index, limit_dict):
    for name in wb.sheetnames:
        ws = wb[name]
        xvalues = Reference(ws, min_col=y_col_index_2, min_row=x_col_index, max_row=config_dict[name])
        yvalues = Reference(ws, min_col=y_col_index_1, min_row=x_col_index, max_row=config_dict[name])
        
        scatter_chart = ScatterChart()
        scatter_chart.title = "Price Index vs Share of Volume"
        scatter_chart.x_axis.title = 'Price Index'
        scatter_chart.y_axis.title = 'Share of Volume'
        scatter_chart.scatterStyle = None
        
        series = Series(yvalues, xvalues, title="Price Index vs Share of Volume")
        scatter_chart.series.append(series)
        scatter_chart.x_axis.major_gridlines = None
        scatter_chart.y_axis.major_gridlines = None
        scatter_chart.x_axis.scaling.min = limit_dict[name]['x'][0]
        scatter_chart.x_axis.scaling.max = limit_dict[name]['x'][1]
        scatter_chart.y_axis.scaling.min = limit_dict[name]['y'][0]
        scatter_chart.y_axis.scaling.max = limit_dict[name]['y'][1]
        scatter_chart.x_axis.number_format = '0'
        scatter_chart.legend = None

        trendline = Trendline()
        trendline.type = "linear" 
        # trendline.spPr = GraphicalProperties(solidFill=ColorChoice(prstClr="orange"))
        series.trendline = trendline
        
        ws[header_index].value = header
        ws[header_index].font = openpyxl.styles.Font(bold=True)
        ws.add_chart(scatter_chart, start_index)
    return wb
        
def create_scatter_plot(df, x_column, y_column, x_lable, y_label, title):
    # Scatter plot using Plotly Express
    fig = px.scatter(df, x=x_column, y=y_column, 
                     labels={x_column: x_lable, y_column: y_label},
                     title=title)
    
    x_numeric = df[x_column].values
    z = np.polyfit(x_numeric, df[y_column], 1)
    p = np.poly1d(z)

    fig.add_scatter(x=df[x_column], y=p(x_numeric), mode='lines', name='Trendline', line=dict(color='red'))

    return fig
def display_chart(data):
    for key in data.keys():
        img = create_scatter_plot(data[key], "Target_Brand Price ix", "Vol_Share_Target_Brand", "Price Index", "Target Brand Vol Share", "Target Brand Vol Share vs Price Index")
        st.markdown(f"<h3>{key}</h3>", unsafe_allow_html=True)
        st.plotly_chart(img, caption='Price Index vs Target Brand Volume Share', use_column_width=True)
        img1 = create_scatter_plot(data[key], "Target_Brand Price ix", "Vol_Share_Total", "Price Index", "Total Vol Share", "Total Vol Share vs Price Index")
        st.plotly_chart(img1, caption='Price Index vs Total Volume Share', use_column_width=True)
@st.cache_resource
def generate_stats(data_dict, x_column, y_column):
    stats_dict = {}
    for key in data_dict.keys():
        stats_dict[key] = {}
        data_filtered = data_dict[key]
        x_values = data_filtered[x_column].values
        y_values = data_filtered[y_column].values
        model = LinearRegression()
        model.fit(x_values.reshape(-1, 1), y_values)
        y_pred = model.predict(x_values.reshape(-1, 1))
        r2 = r2_score(y_values, y_pred)
        correl = data_filtered[x_column].corr(data_filtered[y_column])
        stats_dict[key]["Coefficient"] = model.coef_[0]
        stats_dict[key]["Intercept"] = model.intercept_
        stats_dict[key]["r2"] = r2
        stats_dict[key]["correl"] = correl
    return stats_dict


def write_stats(wb, stats_dict, header, header_start, start_column_label, start_column_value, start_index):
    for name in wb.sheetnames:
        ws = wb[name]
        ws[header_start].value = header
        ws[header_start].font = openpyxl.styles.Font(bold=True)
        ws[f"{start_column_label}{start_index}"].value = "Coefficient"
        ws[f"{start_column_value}{start_index}"].value = round(stats_dict[name]["Coefficient"],3)
        ws[f"{start_column_label}{start_index+1}"].value = "Intercept"
        ws[f"{start_column_value}{start_index+1}"].value = round(stats_dict[name]["Intercept"],2)
        ws[f"{start_column_label}{start_index+2}"].value = "r2"
        ws[f"{start_column_value}{start_index+2}"].value = round(stats_dict[name]["r2"],2)
        ws[f"{start_column_label}{start_index+3}"].value = "correl"
        ws[f"{start_column_value}{start_index+3}"].value = round(stats_dict[name]["correl"],2)
        ws[f"{start_column_label}{start_index+4}"].value = "Equation"
        ws[f"{start_column_value}{start_index+4}"].value = f"y = {stats_dict[name]['Coefficient']}x + {stats_dict[name]['Intercept']}"
    return wb

def display_stats(stats, stats_1):
    for key in stats.keys():
        st.markdown(f"<h3>{key}</h3>", unsafe_allow_html=True)
        st.write(f"**Target Volume Share vs Price Index for {key}:**")
        st.markdown(f"Coefficient: {round(stats[key]['Coefficient'],3)}")
        st.markdown(f"Intercept: {round(stats[key]['Intercept'],2)}")
        st.markdown(f"r2: {round(stats[key]['r2'],2)}")
        st.markdown(f"correl: {round(stats[key]['correl'],2)}")
        st.markdown(f"Equation: y = {stats[key]['Coefficient']}x + {stats[key]['Intercept']}")
        st.write(f"**Total Volume Share vs Price Index for {key}:**")
        st.markdown(f"Coefficient: {round(stats_1[key]['Coefficient'],3)}")
        st.markdown(f"Intercept: {round(stats_1[key]['Intercept'],2)}")
        st.markdown(f"r2: {round(stats_1[key]['r2'],2)}")
        st.markdown(f"correl: {round(stats_1[key]['correl'],2)}")
        st.markdown(f"Equation: y = {stats_1[key]['Coefficient']}x + {stats_1[key]['Intercept']}")
    
def main():
    st.set_page_config(page_title="Price Analytics",
                       page_icon=":books:")
    st.header("Price Analyzer for Home Care")
    uploaded_data = st.sidebar.file_uploader("Upload the price data", accept_multiple_files=False, type=["xlsx", "csv"])
    data = pd.read_excel(uploaded_data) if uploaded_data.name.endswith(".xlsx") else pd.read_csv(uploaded_data)
    data['Week End Date'] = pd.to_datetime(data['Week End Date'])
    filters = st.sidebar.multiselect("Select the colums for filters", data.columns)
    num_filters = st.sidebar.number_input("Number of filters", min_value=1, max_value=10)
    key_column = st.sidebar.selectbox("Select the column to use as key", data.columns)
    data_dict, meta_dict = filter(data, filters, num_filters, key_column)
    if 'data_dict' not in st.session_state:
        st.session_state['data_dict'] = None
    if 'meta_dict' not in st.session_state:
        st.session_state['meta_dict'] = None
    if 'analyzed' not in st.session_state:
        st.session_state['analyzed'] = False
    if 'wb' not in st.session_state:
        st.session_state['wb'] = None
    config_dict = {}
    if st.button("Analyze"):
        with st.spinner("Analyzing..."):
            st.session_state['data_dict'] = data_dict
            st.session_state['meta_dict'] = meta_dict
            st.session_state['analyzed'] = True
            for key in st.session_state["data_dict"].keys():
                config_dict[key] = len(st.session_state["data_dict"][key]) + 1
            if st.session_state['analyzed']:
                tab1, tab2, tab3 = st.tabs(["View Data", "View Scatter Plots", "View Statistics"])
                with tab1:
                    view_data(data_dict, meta_dict) 
                    wb1 = _add_data_worksheet(data_dict, meta_dict)
                    st.session_state.wb = wb1
                with tab2:
                    limit_dict = {}
                    for key in data_dict.keys():
                        limit_dict[key] = {}
                        limit_dict[key]['x'] = [(data_dict[key]['Target_Brand Price ix'].min())*0.9, (data_dict[key]['Target_Brand Price ix'].max())*1.1]
                        limit_dict[key]['y'] = [(data_dict[key]['Vol_Share_Target_Brand'].min())*0.9, (data_dict[key]['Vol_Share_Target_Brand'].max())*1.1]
                    wb2 = generate_charts(st.session_state.wb, 2, 2, 6, config_dict, "Target Brand Vol Share vs Price Index:", "N9", "N10", limit_dict)
                    st.session_state.wb = wb2
                    
                    limit_dict = {}
                    for key in data_dict.keys():
                        limit_dict[key] = {}
                        limit_dict[key]['x'] = [(data_dict[key]['Target_Brand Price ix'].min())*0.9, (data_dict[key]['Target_Brand Price ix'].max())*1.1]
                        limit_dict[key]['y'] = [(data_dict[key]['Vol_Share_Total'].min())*0.9, (data_dict[key]['Vol_Share_Total'].max())*1.1]
                    wb3 = generate_charts(st.session_state.wb, 2, 8, 6, config_dict, "Total Vol Share vs Price Index:", "X9", "X10", limit_dict)
                    st.session_state.wb = wb3
                    display_chart(data_dict)
                with tab3:
                    stats = generate_stats(data_dict, "Target_Brand Price ix", "Vol_Share_Target_Brand")
                    wb4 = write_stats(st.session_state.wb, stats, "Target Brand Stats", "N27", "N", "O", 28)
                    st.session_state.wb = wb4
                    stats_1 = generate_stats(data_dict, "Target_Brand Price ix", "Vol_Share_Total")
                    wb5 = write_stats(st.session_state.wb, stats_1, "Target Brand Stats", "X27", "X", "Y", 28)
                    st.session_state.wb = wb5
                    display_stats(stats, stats_1)
                output = BytesIO()
                wb_final = st.session_state.wb
                wb_final.save(output)
                output.seek(0)
                st.download_button("Download Analysis", output, "Price_Analytics.xlsx", "xlsx")
                
if __name__ == "__main__":
    main()
