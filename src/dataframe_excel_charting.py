
# coding: utf-8

# In[2]:

# import modules
import xlsxwriter
import numpy as np


# In[1]:

class DataFrameExcelCharting(object):
    def __init__(self, df, workbook):
        """takes in a dataframe and a workbook object"""
        self.data = df
        self.workbook = workbook
        self.num_rows = len(df)
        self.column_map = dict()
        self._to_excel = 0
        
#     def createWorkBook(self, workbook_name):
#         """create work book"""
#         self.workbook_name = workbook_name
#         self.workbook = xlsxwriter.Workbook(self.workbook_name)
        
    def createWorkSheet(self, worksheet_name):
        """create work sheet in self.workbook"""
        self.worksheet_name = worksheet_name
        self.worksheet = self.workbook.add_worksheet(self.worksheet_name)
        
    def closeWorkBook(self):
        """close work book"""
        self.workbook.close()
        
    def createChart(self, chart_type, x_axis, y_axis, title):
        """create chart"""
        self.chart = self.workbook.add_chart({'type': chart_type})
        self.chart.set_y_axis({'name': y_axis})
        self.chart.set_x_axis({'name': x_axis})
        self.chart.set_title({'name': title})
        
    def insertChart(self, insert_col, insert_row):
        """insert chart to work sheet"""
        self.worksheet.insert_chart('{0}{1}'.format(insert_col, insert_row), self.chart)
        
    def writeToExcel(self, sheet_name="Sheet1"):
        """write data frame to excel with header"""
        # create workbook and worksheet
        self.createWorkSheet(sheet_name)
        # write data frame to excel
        header_row = 1
        for i in range(0, len(self.data.columns)):
            col = xlsxwriter.utility.xl_col_to_name(i)
            self.worksheet.write("{0}{1}".format(col, header_row), self.data.columns[i])
            # fill NaN with -1
            self.worksheet.write_column("{0}{1}".format(col, header_row + 1), 
                                        self.data[self.data.columns[i]].replace([np.inf, -np.inf], np.nan).fillna(-1).values)
            self.column_map[self.data.columns[i]] = col
        # change the flag
        self._to_excel = 1
            
    def getTopN(self, columns=None, n=5, ascending=True, inplace=True):
        """Given one or more columns """
        assert columns is not None, "Please specify a list of columns to get top"
        # sort the dataframe 
        self.data.sort(columns=columns, axis=0, ascending=ascending, inplace=inplace)
        self._to_excel = 0
        print "Dataframe has been changed, please write to Excel again"
        
        
        
    def topNChart(self, columns=None, n=5, category_col=None, 
                  chart_type="column", x_axis="name", y_axis="value", title="title"):
        """top N chart"""
        assert self._to_excel == 1, "Please write data to excel first"
        assert columns is not None, "Please specify a list of columns to plot"
        assert category_col is not None, "Please specify a categorical column"
        
        if not isinstance(columns, list):
            columns = list(columns)
        
        cat_col = self.column_map[category_col]
            
        # create a chart
        self.createChart(chart_type, x_axis, y_axis, title)
        
        # add series to chart
        data_row = 2
        for i in range(0, len(columns)):
            col = self.column_map[columns[i]]
            self.chart.add_series({'values': "={0}!${1}${2}:${1}${3}".format(self.worksheet_name, col, data_row, data_row + n - 1), 
                                   'name': columns[i], 
                                   'categories': "={0}!${1}${2}:${1}${3}".format(self.worksheet_name, cat_col, data_row, data_row + n - 1)
                                  })
        # insert the chart    
        self.insertChart(col, self.num_rows + 2)
    
    def getBucketsCounts(self, column=None, n_buckets=5):
        """get buckets and counts"""
        # None column check
        assert column is not None, "Please specify a column name"
        assert column in self.data.columns, "Please choose a valid column"
        
        # check n_buckets
        assert n_buckets >= 1, "Please specify a positive number of buckets"
        assert isinstance(n_buckets, int), "Please use a positive integer as n_buckets"
        
        # filter out NaN and Inf
        data_array = self.data[column]
        data_array = data_array[~np.isnan(data_array)]
        data_array = data_array[~np.isinf(data_array)]
        # calculate count and interval
        lower = np.floor(data_array.min())
        upper = np.ceil(data_array.max())
        diff = upper - lower
        bins = [np.arange(lower, upper, diff / n_buckets)]
        bins = np.append(bins, upper)
        count, interval = np.histogram(data_array, bins=bins)
        interval = ["[{0}, {1})".format(interval[i], interval[i + 1]) for i in range(0,len(interval) - 1)] # + diff / n_buckets / 2.0 
        return count, interval
            
    def bucketsNChart(self, column=None, n_buckets=5,
                      chart_type="column", x_axis="name", y_axis="value", title="title"):
        """top N chart"""
        assert self._to_excel == 1, "Please write data to excel first"
        
        # create a new chart
        self.createChart(chart_type, x_axis, y_axis, title)
        
        """plot bucket chart"""
        count, interval = self.getBucketsCounts(column=column, n_buckets=n_buckets)
        col = self.column_map[column]
        row = self.num_rows + 3
        self.worksheet.write_column('{0}{1}'.format(col, row), count)
        self.worksheet.write_column('{0}{1}'.format(col, row + n_buckets), interval)
        
        self.chart.add_series({'values': '={0}!${1}${2}:${1}${3}'.format(self.worksheet_name, col, row, row + n_buckets - 1),
                               'categories': '={0}!${1}${2}:${1}${3}'.format(self.worksheet_name, col, row + n_buckets, row + n_buckets + n_buckets -1),
                               'gap': 5
                              })
        # insert chart
        self.insertChart(col, row)
        
    # TODO: scatter plot    
    def scatterPlot(self, columns=None):
        return


# In[ ]:



