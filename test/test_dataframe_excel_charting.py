
# coding: utf-8

# In[1]:

# test data frame
from dataframe_excel_charting import DataFrameExcelCharting
import pandas as pd
import numpy as np
import xlsxwriter

# In[5]:
# write to single workbook
workbook = xlsxwriter.Workbook("test_charting.xlsx")

# test NaN and Inf
df = pd.DataFrame(np.random.randn(100, 4), columns=["col1", "col2", "col3", "col4"])
df["category"] = ["category{}".format(i) for i in range(100)]
df["col4"][0] = None
df["col4"][1] = np.inf
print df.head()

test_class = DataFrameExcelCharting(df, workbook)
# test_class.aggregateBucketsChart("col2", 8, ["col4", "col3"], "count")
# test_class.getTopN(["col1", "col3"])
test_class.writeToExcel("test basic charts")
test_class.topNChart(columns=["col1", "col3"], n=10, category_col="category")
test_class.aggregateBucketsChart("col4", 5, ["col1", "col2", "col3"], "mean", "test_basic_chart1", "x", "y", "what title it is")
test_class.bucketsNChart(column="col2",n_buckets=10)
# test_class.getBucketsCounts(column="col4", n_buckets=10)
test_class.scatterPlot(["col1", "col4"], "category", "scatter")
test_class.aggregateBucketsChart("col2", 10, ["col4"], "count")
# test_class.insertImage("A", test_class.num_rows + 3 + test_class._num_charts * 15, "test_geoplot.png")
# for i in range(117, 137):
#     test_class.worksheet.set_row(i, None, None, {'hidden': True})

# edge cases test 
df = pd.DataFrame({
    'A': [['a', 'b', 'c'], ["A", "B", 321], np.nan, (123, 43), "asdf", None, np.ndarray(10)]
    })
print df.head()
test_class = DataFrameExcelCharting(df, workbook)
# test_class.getTopN(["Marginal_Population_Served", "Marginal_Total_Demand", "Total_Demand"])
test_class.writeToExcel("test write list")

# list of DFs
df1 = pd.DataFrame(np.random.randn(100, 4), columns=["col1", "col2", "col3", "col4"])
df1["category"] = ["category{}".format(i) for i in range(100)]
df2 = pd.DataFrame(np.random.randn(100, 4), columns=["col5", "col6", "col7", "col8"])
df2["cate"] = ["cate{}".format(i) for i in range(100)]
print df1.head()
print df2.head()


# In[7]:

# list of DFs
dfs = [df1, df2]
i = 1
for df in dfs:
    test_class = DataFrameExcelCharting(df, workbook)
    test_class.writeToExcel("test multiple sheet{0}".format(i))
    i = i + 1

# test geo plot
df = pd.read_csv('https://raw.githubusercontent.com/plotly/datasets/master/2014_us_cities.csv')
print df.head()

test_class = DataFrameExcelCharting(df, workbook)
test_class.writeToExcel(sheet_name="geo_plot")
test_class.geoPlot(text_col="name", value_col="pop", lat="lat", lon="lon", 
                   n_buckets=5, image_name="geo_plot", 
                   scale=5000, plot_type="scattergeo", 
                   scope='world', map_type="natural earth")


# test with data
expected_metrics_by_node_df = pd.DataFrame(
                                           columns=[
                                                    "Settlement_ID",
                                                    "Network_ID",
                                                    "Node_ID",
                                                    "lat",
                                                    "lon",
                                                    "Marginal_Population_Served",
                                                    "Population_Served",
                                                    "Marginal_Total_Demand",
                                                    "Total_Demand",
                                                    "Total_Transit",
                                                    "Tower_Cost",
                                                    "Radio_Cost",
                                                    "Power_Cost",
                                                    "Hop_Count",
                                                    "Single_Site_Failure_Effect",
                                                    "Node_Utility",
                                                    ]
                                           )
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_216:1",
                                                                    "lat_-15.092711061872_lon_-73.745867783712",
                                                                     "F1",
                                                                     -15.092711061872,
                                                                     -73.745867783712,
                                                                     0.0,
                                                                     0.0,
                                                                     0.0,
                                                                     0.0,
                                                                     26608.5,
                                                                     15000.0,
                                                                     2500.0,
                                                                     10000.0,
                                                                     np.nan,
                                                                     26608.5,
                                                                     np.nan,
                                                                     ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_216:1",
                                                                    "lat_-15.092711061872_lon_-73.745867783712",
                                                                     "N1",
                                                                     -15.090779,
                                                                     -73.720231,
                                                                     10000.0,
                                                                     10000.0,
                                                                     17739.0,
                                                                     17739.0,
                                                                     26608.5,
                                                                     11000.0,
                                                                     57300.0,
                                                                     10000.0,
                                                                     1.0,
                                                                     26608.5,
                                                                     1.0,
                                                                     ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_216:1",
                                                                    "lat_-15.092711061872_lon_-73.745867783712",
                                                                    "N2",
                                                                    -15.113981,
                                                                    -73.723150,
                                                                    4000.0,
                                                                    4000.0,
                                                                    5913.0,
                                                                    5913.0,
                                                                    5913.0,
                                                                    15000.0,
                                                                    24800.0,
                                                                    10000.0,
                                                                    2.0,
                                                                    5913.0,
                                                                    1.0,
                                                                    ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_216:1",
                                                                    "lat_-15.092711061872_lon_-73.745867783712",
                                                                    "N3",
                                                                    -15.121024,
                                                                    -73.682895,
                                                                    2000.0,
                                                                    2000.0,
                                                                    2956.5,
                                                                    2956.5,
                                                                    2956.5,
                                                                    28000.0,
                                                                    24800.0,
                                                                    10000.0,
                                                                    2.0,
                                                                    2956.5,
                                                                    1.0,
                                                                    ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_121:1",
                                                                    "lat_-15.05366466651_lon_-73.770922082143",
                                                                    "F2",
                                                                    -15.05366466651,
                                                                    -73.770922082143,
                                                                    0.0,
                                                                    0.0,
                                                                    0.0,
                                                                    0.0,
                                                                    2217.375,
                                                                    28000.0,
                                                                    5000.0,
                                                                    10000.0,
                                                                    np.nan,
                                                                    2217.375,
                                                                    np.nan,
                                                                    ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_121:1",
                                                                    "lat_-15.05366466651_lon_-73.770922082143",
                                                                    "N4",
                                                                    -15.027052,
                                                                    -73.775133,
                                                                    1000.0,
                                                                    1000.0,
                                                                    1478.25,
                                                                    1478.25,
                                                                    1478.25,
                                                                    15000.0,
                                                                    24800.0,
                                                                    10000.0,
                                                                    1.0,
                                                                    1478.25,
                                                                    1.0,
                                                                    ]
expected_metrics_by_node_df.loc[len(expected_metrics_by_node_df)] = [
                                                                    "PER_5_121:1",
                                                                    "lat_-15.05366466651_lon_-73.770922082143",
                                                                    "N5",
                                                                    -15.006327,
                                                                    -73.759684,
                                                                    500.0,
                                                                    500.0,
                                                                    739.125,
                                                                    739.125,
                                                                    739.125,
                                                                    28000.0,
                                                                    24800.0,
                                                                    10000.0,
                                                                    1.0,
                                                                    739.125,
                                                                    1.0,
                                                                    ]

test_class = DataFrameExcelCharting(expected_metrics_by_node_df, workbook)
test_class.getTopN(["Marginal_Population_Served", "Marginal_Total_Demand", "Total_Demand"])
test_class.writeToExcel("test data")
test_class.topNChart(columns=["Marginal_Population_Served", "Marginal_Total_Demand", "Total_Demand"], 
                     n=10, category_col="Node_ID")
test_class.bucketsNChart(column="Marginal_Total_Demand",n_buckets=3)
test_class.geoPlot(text_col="Node_ID", value_col="Total_Transit", lat="lat", lon="lon", 
                   n_buckets=5, image_name="node_total_transit_plot", 
                   scale=100, plot_type="scattergeo", 
                scope='world', map_type="natural earth")
# test aggregate buckets chart
test_class.aggregateBucketsChart("Marginal_Total_Demand", 5, ["Total_Demand",
                                                    "Total_Transit",
                                                    "Tower_Cost",
                                                    "Radio_Cost",
                                                    "Power_Cost",])
# close the workbook
workbook.close()
