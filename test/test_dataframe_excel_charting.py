
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
test_class.writeToExcel("test_sheet")
test_class.topNChart(columns=["col1", "col3"], n=10, category_col="category")
test_class.bucketsNChart(column="col4",n_buckets=10)
test_class.getBucketsCounts(column="col4", n_buckets=10)
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


# In[ ]:



