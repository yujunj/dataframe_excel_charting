
# coding: utf-8

# In[2]:

# import modules
import xlsxwriter
import numpy as np
import plotly.offline as offline
# import plotly.plotly as py
import colorsys


class DataFrameExcelCharting(object):
    """Pandas Data Frame saving and plotting in Excel.
    
    DataFrameExcelChartingclass uses xlsxwriter as driver to interact 
    with Pandas Data Frame and Microsoft Excel. The class also supports multiple 
    charting options described below.

    Attributes:
        data (pandas.DataFrame): Input Pandas Dataframe.
        workbook(xlsxwriter.Workbook): Input MS Workbook.
        num_rows (int): length of the input dataframe.
        column_map (dict(str)): dictionary of input data frame columns to Excel column indicator.
        
    TODO:
    save the geoplot image locally and insert into work sheet

    """
    def __init__(self, df, workbook):
        """Constructor
        
        _to_excel (boolean): indicates whether the data has been written to Excel.
        _num_charts (int): counter of number of charts inserted.
        _num_charts (int): counter of number of URL inserted.
        
        """
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
        """Create Work Sheet Method

        Note:
            Work Book object should be passed in when instantiation.
            In order for multiple sheets to write to the same work book. 

        Args:
            worksheet_name: the name of the work sheet.
            
        """
        self.worksheet_name = worksheet_name
        self.worksheet = self.workbook.add_worksheet(self.worksheet_name)
        
    def closeWorkBook(self):
        """Close Work Book Method.

        Note:
            A work book should be closed after all data finished saving to it.
            
        """
        self.workbook.close()
        
    def createChart(self, chart_type, x_axis, y_axis, title):
        """Create Chart Method.

        Note:
            Simply create a chart without any data.
            The chart object is save to self.chart as an attribute.

        Args:
            chrat_type: The type of the chart, i.e. "column" or "scatter" etc.
            x_axis: The name of x axis.
            y_axis: The name of y axis.
            title: The name of the chart
            
        """
        self.chart = self.workbook.add_chart({'type': chart_type})
        self.chart.set_y_axis({'name': y_axis})
        self.chart.set_x_axis({'name': x_axis})
        self.chart.set_title({'name': title})
        
    def insertChart(self, insert_col, insert_row):
        """Insert Chart Method.

        Note:
            Insert self.chart into specified location in self.worksheet.
            Then increase the chart counter by one

        Args:
            insert_col: Excel column indicator to insert the chart.
            insert_row: Excel row number to insert the chart.
            
        """
        self.worksheet.insert_chart('{0}{1}'.format(insert_col, insert_row), self.chart)
        self._num_charts = self._num_charts + 1
        
    def insertURL(self, insert_col, insert_row, file_url, string):
        """Insert URL Method.

        Note:
            Insert file_url into specified location with string as name.

        Args:
            insert_col: Excel column indicator to insert the URL.
            insert_row: Excel row number to insert the URL.
            file_url: File URL to insert.
            string: Name or mask of the file_url.
            
        """
        self.worksheet.write_url('{0}{1}'.format(insert_col, insert_row), file_url, string=string)
        self._num_urls = self._num_urls + 1
        
    def writeToExcel(self, sheet_name="Sheet1"):
        """Write to Excel Method.

        Note:
            Create a work sheet as the name specified. 
            Write the self.data, a.k.a input Pandas Data Frame to specified work sheet.
            The process is needed before any Excel charting methods are called.

        Args:
            sheet_name: The sheet name specified.
            
        """
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
            
    def getTopN(self, columns=None, n=10, ascending=False, inplace=True):
        """Get Top N Method.

        Note:
            Sort self.data according to the desired columns on specified order.
            Since the order is changed, the data needs to be re-written to work sheet. 

        Args:
            columns: A list of columns to sort on.
            ascending: False by default to sort the desired columns on a descending order.
            inplace: Operation is done on self.data instead of creating a new Data Frame. 
            
        """
        assert columns is not None, "Please specify a list of columns to get top"
        # sort the dataframe 
        self.data.sort_values(by=columns, axis=0, ascending=ascending, inplace=inplace)
        self._to_excel = 0
        print "Dataframe has been changed, please write to Excel again"
        
        
    def topNChart(self, columns=None, n=5, category_col=None, 
                  chart_type="column", x_axis="name", y_axis="value", title="title"):
        """Top N Chart Method.

        Note:
            Generate a bar chart of columns based on input parameters.
            Assuming self.data has already been sorted and written to Excel
            Increase the chart at the end of the data.
            
        Args:
            columns: Desired column(s) to plot.
            n: Number of rows to plot.
            category_col: Categorical column to be shown as the value on x axis.
            chart_type: "column" by default for bar chart.
            x_axis: name of x axis.
            y_axis: name of y axis.
            title: name of the chart title.

        """
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
        """Get Buckets Counts Method.

        Note:
            Get the counts and buckets based on input column and n_buckets.

        Args:
            column: The interested column.
            n_buckets: Number of buckets needed.
            str_interval: Whether to return the interval as list of strings or list of tuples.

        Returns:
            count (list(int)): Number of records fall in the corresponding bucket.
            interval (list(str)): List of string of buckets boundary, if str_interval is True. Only for charting.
            or
            interval (list(tuple)): List of tuples of buckets boundary, if str_interval is False.  

        """
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
        """Buckets N Chart Method.

        Note:
            Generate buckets chart based on desired columns and other inputs.
            Assuming self.data has already been sorted and written to Excel
            Increase the chart at the end of the data.

        Args:
            column: Desired column to do buckets plot.
            n_buckets: Number of buckets needed.
            chart_type: Column by default for bar chart.
            x_axis: Name of x axis.
            y_axis: Name of y axis.
            title: Name of chart title. 

        """
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
        """Scatter Plot Method.

        Note:
            Generate scatter plot based on desired columns and other inputs.
            Assuming self.data has already been sorted and written to Excel
            Increase the chart at the end of the data.

        Args:
            columns: Desired column(s) to plot.
            category_col: Categorical column to show on the x axis.
            chart_type: "scatter" by default for scatter plot.
            x_axis: Name of x axis.
            y_axis: Name of y axis.
            title: Name of chart title. 

        """
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
        """Insert Image Method.

        Note:
            Insert the image from image_path to specified location.

        Args:
            insert_col: Excel column indicator to insert the image.
            insert_row: Excel row number to insert the image.
            image_path: Path to the image to insert.

        """
        self.worksheet.insert_image("{0}{1}".format(insert_col, insert_row), image_path)
        self._num_charts = self._num_charts + 1
        
    def _getRGBColors(self, n=5):
        """Get RGB Colors Method.

        Note:
            Interval method to generate RGB colors.

        Args:
            n: Number of colors to generate.

        """
        assert n >= 1, "Please specify a positive number of colors"
        assert isinstance(n, int), "Please use a positive integer as n"
        
        HSV_tuples = [(x * 1.0 / n, 0.5, 0.5) for x in range(n)]
        RGB_tuples = map(lambda x: colorsys.hsv_to_rgb(*x), HSV_tuples)
        return map(lambda x: "rgb({0}, {1}, {2})".format(round(x[0] * 255), round(x[1] * 255), round(x[2] * 255)), RGB_tuples)
        
    def geoPlot(self, text_col=None, value_col=None, lat="lat_col", lon="lon_col", 
                n_buckets=5, image_name=None, 
                scale=100, plot_type="scattergeo", 
                scope='south america', map_type="equirectangular"):
        """Geo Plot Method.

        Note:
            Generate GEO based plot.
            Output a interactive map as HTML in the current folder.
            Insert the URI to the file into current work sheet.
            
        Reference:
            https://plot.ly/python/reference/#layout-geo-scope

        Args:
            text_col: Column to show on the map as text, which determines the name of bubbles.
            value_col: Column to show on the map as value, which determines the size of bubbles. 
            lat: latitude column.
            lon: longitude column.
            n_buckets: Number of categories desired. 
            image_name: Name of the HTML file generated. 
            scale: A number to adjust magnitude of value_col to proper bubble size. bubble size is equal to magnitude of value_col divided by scale.
            plot_type: "scattergeo' by default
            scope: enumeration of ("world" | "usa" | "europe" | "asia" | "africa" | "north america" | "south america").
            map_type: enumberation of ("equirectangular" | "mercator" | "orthographic" | "natural earth" | "kavrayskiy7" | "miller" | "robinson" | "eckert4" | "azimuthal equal area" | "azimuthal equidistant" | "conic equal area" | "conic conformal" | "conic equidistant" | "gnomonic" | "stereographic" | "mollweide" | "hammer" | "transverse mercator" | "albers usa" | "winkel tripel" | "aitoff" | "sinusoidal")
            
        """
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



