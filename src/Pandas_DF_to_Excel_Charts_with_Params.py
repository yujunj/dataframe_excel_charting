
# coding: utf-8

# In[1]:

# import modules
import xlsxwriter
import random
import pandas as pd
import numpy as np


# In[2]:

# top N bar plot
def top_n_chart(df, columns=None, n=5, category=None, file_name="test_file.xlsx", x_axis="name", y_axis="value", title="title"):
    """
    params:
    
    df: pandas dataframe
    columns: list of columns to plot
    n: number of top to plot
    category: string of categorical column name
    """
    # Create workbook and add worksheet
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    
    # Create a new Chart object.
    chart_type="column"
    chart = workbook.add_chart({'type': chart_type})
    chart.set_y_axis({'name': y_axis})
    chart.set_x_axis({'name': x_axis})
    chart.set_title({'name': title})
    
    # convert single column to list
    if not isinstance(columns, list):
        columns = list(columns)
    
    # add category to columns
    if category:
        columns.insert(0, category)
    
    # Write some data to add to plot on the chart.
    if columns:
        data = df[columns]
    else:
        data = df
    
    for i in range(0, len(data.columns)):
        col = xlsxwriter.utility.xl_col_to_name(i)
        worksheet.write_column("{0}1".format(col), data[data.columns[i]].values)
        if data.columns[i] == category:
            continue
        # Configure the chart. In simplest case we add one or more data series.
        if category:
            chart.add_series({'values': "=Sheet1!${0}$1:${0}${1}".format(col, n), 
                              'name': data.columns[i], 
                              'categories': "=Sheet1!$A$1:$A${0}".format(n)})
        else:
            chart.add_series({'values': "=Sheet1!${0}$1:${0}${1}".format(col, n), 
                              'name': data.columns[i]})
    
    # Insert the chart into the worksheet.
    worksheet.insert_chart('{0}{1}'.format(xlsxwriter.utility.xl_col_to_name(i + 1), n + 2), chart)

    workbook.close()


# In[3]:

# bucket plot
def n_buckets_chart(df, column=None, n_buckets=5, file_name="test_file.xlsx", x_axis="value", y_axis="count", title="title"):
    """
    params:
    
    df: pandas dataframe
    column: string of column name to bucketize
    n_buckets: number of buckets
    """
    # Create workbook and add worksheet
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    
    # Create a new Chart object.
    chart_type="column"
    chart = workbook.add_chart({'type': chart_type})
    chart.set_y_axis({'name': y_axis})
    chart.set_x_axis({'name': x_axis})
    chart.set_title({'name': title})
    
    # generate buckets and counts
    lower = np.floor(df[column].min())
    upper = np.ceil(df[column].max())
    diff = upper - lower
    bins = [np.arange(lower, upper, diff / n_buckets)]
    bins = np.append(bins, upper)
    count, interval = np.histogram(df[column], bins=bins)
    interval = ["{0} to {1}".format(interval[i], interval[i + 1]) for i in range(0,len(interval) - 1)] # + diff / n_buckets / 2.0 
    
    # insert columns to xlsx
    worksheet.write_column('A1', count)
    worksheet.write_column('B1', interval)


    # Configure the chart. In simplest case we add one or more data series.
    chart.add_series({'values': '=Sheet1!$A$1:$A${0}'.format(n_buckets),
                      'categories': '=Sheet1!$B$1:$B${0}'.format(n_buckets),
                      'gap': 5
                     })

    # Insert the chart into the worksheet.
    worksheet.insert_chart('{0}{1}'.format("A", n_buckets + 2), chart)

    workbook.close()


# In[4]:

# test data frame
df = pd.DataFrame(np.random.randn(50, 4), columns=list('ABCD'))
df["category"] = ["category{}".format(i) for i in range(50)]
print df.head()


# In[5]:

# sample use case
top_n_chart(df, 
            columns=["A", "B", "C"], 
            n=10, 
            category="category", 
            file_name="test_top_N.xlsx", 
            x_axis="Category", 
            y_axis="Value", 
            title="Top N of Category")


# In[6]:

# sample use case
n_buckets_chart(df, 
                column="A", 
                n_buckets=10, 
                file_name="test_n_buckets.xlsx", 
                x_axis="Interval", 
                y_axis="Count", 
                title="N Buckets of A")


# In[ ]:



