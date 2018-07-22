
# coding: utf-8

# In[1]:

# test data frame
from dataframe_excel_charting import DataFrameExcelCharting
import pandas as pd
import numpy as np
import xlsxwriter


# In[2]:

df = pd.DataFrame(np.random.randn(100, 4), columns=["col1", "col2", "col3", "col4"])
df["category"] = ["category{}".format(i) for i in range(100)]
df["col4"][0] = None
df["col4"][1] = np.inf
print df.head()


# In[5]:

workbook = xlsxwriter.Workbook("test_charting.xlsx")
test_class = DataFrameExcelCharting(df, workbook)
test_class.getTopN(["col1", "col3"], 10)
test_class.writeToExcel("test_sheet")
test_class.topNChart(columns=["col1", "col3"], n=10, category_col="category")
test_class.bucketsNChart(column="col4",n_buckets=10)
test_class.getBucketsCounts(column="col4", n_buckets=10)
test_class.scatterPlot(["col1", "col4"], "category", "scatter")
test_class.insertImage("A", test_class.num_rows + 3 + test_class._num_charts * 15, "test_geoplot.png")
test_class.closeWorkBook()


# In[6]:

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
workbook = xlsxwriter.Workbook("huge_workbook.xlsx")
i = 1
for df in dfs:
    test_class = DataFrameExcelCharting(df, workbook)
    test_class.writeToExcel("test_sheet{0}".format(i))
    i = i + 1
workbook.close()


df = pd.read_csv('https://raw.githubusercontent.com/plotly/datasets/master/2014_us_cities.csv')
print df.head()

workbook = xlsxwriter.Workbook("geo_plot.xlsx")
test_class = DataFrameExcelCharting(df, workbook)
test_class.writeToExcel(sheet_name="geo_plot")
test_class.geoPlot(text_col="name", value_col="pop", lat="lat", lon="lon", 
                   n_buckets=5, image_name="geo_plot", 
                   scale=5000, plot_type="scattergeo", 
                   scope='usa', map_type="albers usa")
workbook.close()


