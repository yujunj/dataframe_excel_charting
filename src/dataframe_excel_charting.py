
# coding: utf-8

# In[2]:

# import modules
import xlsxwriter
import numpy as np
import plotly.offline as offline
import plotly.plotly as py
import colorsys




# In[1]:

class DataFrameExcelCharting(object):
    def __init__(self, df, workbook):
        """takes in a dataframe and a workbook object"""
        self.data = df
        self.workbook = workbook
        self.num_rows = len(df)
        self.column_map = dict()
        self._to_excel = 0
        self._num_charts = 0
        self._num_urls = 0
        
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
        self._num_charts = self._num_charts + 1
        
    def insertURL(self, insert_col, insert_row, file_url, string):
        """insert URL to work sheet"""
        self.worksheet.write_url('{0}{1}'.format(insert_col, insert_row), file_url, string=string)
        self._num_urls = self._num_urls + 1
        
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
            
    def getTopN(self, columns=None, n=5, ascending=False, inplace=True):
        """Given one or more columns """
        assert columns is not None, "Please specify a list of columns to get top"
        # sort the dataframe 
        self.data.sort_values(by=columns, axis=0, ascending=ascending, inplace=inplace)
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
        # insert the chart and prevent charts overlapping 
        self.insertChart(col, self.num_rows + 2 + 15 * self._num_charts)
    
    def getBucketsCounts(self, column=None, n_buckets=5, str_interval=True):
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
        if str_interval:
            interval = ["[{0}, {1})".format(interval[i], interval[i + 1]) for i in range(0,len(interval) - 1)]
        else:
            interval = [(interval[i], interval[i + 1]) for i in range(0,len(interval) - 1)]
        return count, interval
            
    def bucketsNChart(self, column=None, n_buckets=5,
                      chart_type="column", x_axis="name", y_axis="value", title="title"):
        """top N chart"""
        assert self._to_excel == 1, "Please write data to excel first"
        
        # create a new chart
        self.createChart(chart_type, x_axis, y_axis, title)
        
        # plot bucket chart
        count, interval = self.getBucketsCounts(column=column, n_buckets=n_buckets)
        col = self.column_map[column]
        row = self.num_rows + 3
        self.worksheet.write_column('{0}{1}'.format(col, row), count)
        self.worksheet.write_column('{0}{1}'.format(col, row + n_buckets), interval)
        
        self.chart.add_series({'values': '={0}!${1}${2}:${1}${3}'.format(self.worksheet_name, col, row, row + n_buckets - 1),
                               'categories': '={0}!${1}${2}:${1}${3}'.format(self.worksheet_name, col, row + n_buckets, row + n_buckets + n_buckets -1),
                               'gap': 5
                              })
        # insert chart and prevent charts overlapping
        self.insertChart(col, row + self._num_charts * 15)
        
    def scatterPlot(self, columns=None, category_col=None, 
                    chart_type="scatter", x_axis="name", y_axis="value", title="title"):
        """plot scatter chart of x_column vs y_column"""
        assert self._to_excel == 1, "Please write data to excel first"
        assert columns is not None, "Please specify a list of columns to plot"
        assert category_col is not None, "Please specify a categorical column"
        
        if not isinstance(columns, list):
            columns = list(columns)
        
        cat_col = self.column_map[category_col]
            
        # create a chart
        self.createChart(chart_type, x_axis, y_axis, title)
        for column in columns:
            col = self.column_map[column]
            self.chart.add_series(
                {
                'values': '={0}!${1}{2}:${1}{3}'.format(self.worksheet_name, col, 2, self.num_rows), 
                'categories': "={0}!${1}{2}:${1}{3}".format(self.worksheet_name, cat_col, 2, self.num_rows),
                'name': "={0}!${1}{2}".format(self.worksheet_name, col, 1)
                }
            )
        # insert chart and prevent charts overlapping
        self.insertChart(col, self.num_rows + 3 + self._num_charts * 15)
    
    def insertImage(self, insert_col, insert_row, image_path=None):
        """insert image to worksheet at insert col and insert row"""
        self.worksheet.insert_image("{0}{1}".format(insert_col, insert_row), image_path)
        self._num_charts = self._num_charts + 1
        
    def _getRGBColors(self, n=5):
        """generate n colors and convert to RGB strings"""
        assert n >= 1, "Please specify a positive number of colors"
        assert isinstance(n, int), "Please use a positive integer as n"
        
        HSV_tuples = [(x * 1.0 / n, 0.5, 0.5) for x in range(n)]
        RGB_tuples = map(lambda x: colorsys.hsv_to_rgb(*x), HSV_tuples)
        return map(lambda x: "rgb({0}, {1}, {2})".format(round(x[0] * 255), round(x[1] * 255), round(x[2] * 255)), RGB_tuples)
        
    def geoPlot(self, text_col=None, value_col=None, lat="lat_col", lon="lon_col", 
                n_buckets=5, image_name=None, 
                scale=100, plot_type="scattergeo", 
                scope='south america', map_type="equirectangular"):
        """Plot Geo map based on latitude and longitude"""
        assert set([text_col, value_col, lat, lon]).issubset(set(self.data.columns)), "Please specify valid column names for text, value, lat and lon"
                
        counts, limits = self.getBucketsCounts(value_col, n_buckets, str_interval=False)
        colors = self._getRGBColors(n_buckets)
        if np.max(counts) - np.min(counts) > 0.9 * self.num_rows:
            print "Buckets are highly skewed"
            
        places = []
        for i in range(len(limits)):
            lim = limits[i]
            # handle the edge case
            if i == len(limits) - 1:
                df_sub = self.data[(self.data[value_col] >= lim[0]) & (self.data[value_col] <= lim[1])]
            else:
                df_sub = self.data[(self.data[value_col] >= lim[0]) & (self.data[value_col] < lim[1])]
            place = dict(
                type = plot_type,
                # locations = ["peru"],
                # locationmode = "USA-states",
                lon = df_sub[lon],
                lat = df_sub[lat],
                text = df_sub[text_col],
                marker = dict(
                    size = df_sub[value_col] / scale,
                    color = colors[i],
                    line = dict(width=0.5, color='rgb(40,40,40)'),
                    sizemode = 'area'
                ),
                name = '[{0}, {1})'.format(lim[0],lim[1]) )
            places.append(place)
    
        layout = dict(
                title = value_col,
                showlegend = True,
                geo = dict(
                    scope = scope,
                    projection=dict( type=map_type, scale = 1),
                    showland = True,
                    landcolor = 'rgb(217, 217, 217)',
                    subunitwidth=1,
                    countrywidth=1,
                    subunitcolor="rgb(255, 255, 255)",
                    countrycolor="rgb(255, 255, 255)"
                ),
            )
    
        fig = dict(data=places, layout=layout)
        # py.iplot(fig, validate=False, filename='d3-bubble-map-populations' )
        file_url = offline.plot(fig, validate=False, 
                            filename='{}.html'.format(image_name), auto_open=False)
        # py.image.save_as(fig, '{}.png'.format(image_name), scale=3)
        if self._to_excel:
            self.insertURL(insert_col="A", insert_row=self.num_rows + 2 + self._num_urls, file_url=file_url, string=image_name)


# In[ ]:



