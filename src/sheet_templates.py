'''
Created on Aug 5, 2018

@author: arthur
'''
# import xlsxwriter
from xlsxwriter import worksheet

class SheetTemplates(object):
    '''
    classdocs
    '''


    def __init__(self, workbook):
        '''
        Constructor
        '''
        self.workbook = workbook
        
    def createWorkSheet(self, worksheet_name):
        """Create Work Sheet
        
        Args:
            worksheet_name: name of the sheet
        """
        return self.workbook.add_worksheet(worksheet_name)
        
    def openWorkSheet(self, worksheet_name):
        """Open Work Sheet
        
        Args:
            worksheet_name: name of the sheet
        
        """
        assert worksheet_name in map(lambda x: x.get_name(), self.workbook.worksheets()), "Sheet not found in workbook, try creating a new one"
        return self.workbook.get_worksheet_by_name(worksheet_name)
    
    def mergeCellsAndWrite(self, worksheet, cell_range, title):
        """Write Sheet Title
        
        Args:
            worksheet: xlsxwriter worksheet object
            cell_range: Excel sheet cell range, e.g. "C1:L1"
            title: string
        
        """
        # define merge_format first
        merge_format = self.workbook.add_format({'align': 'center'})
        # merge the first row from C1 to L1 and write title to it
        worksheet.merge_range(cell_range, title, merge_format)
        
    def writeTableFrame(self, worksheet, table_type):
        """Write Table Frame
        
        Args:
            table_type: 1: Competition Analysis
                        2: Country TAM Analysis-PER
                        3: Network Analysis
        
        """
        if table_type == 1:
            # first line of table
            worksheet.write_string("B5", "Operator Summary")
            self.mergeCellsAndWrite(worksheet, "C5:G5", "Pops")
            self.mergeCellsAndWrite(worksheet, "H5:L5", "Settlements")
            # second line of table
            worksheet.write_string("C6", "Total CPOPs")
            worksheet.write_string("D6", "4G")
            worksheet.write_string("E6", "3G+4G")
            worksheet.write_string("F6", "2G Only")
            worksheet.write_string("G6", "Unconnected")
            worksheet.write_string("H6", "Total")
            worksheet.write_string("I6", "4G")
            worksheet.write_string("J6", "3G+4G")
            worksheet.write_string("K6", "2G Only")
            worksheet.write_string("L6", "Unconnected")
        elif table_type == 2:
            # first line of table
            self.mergeCellsAndWrite(worksheet, "C4:D4", "PER")
            self.mergeCellsAndWrite(worksheet, "E4:K4", "Settlements")
            # second line of table
            worksheet.write_string("C5", "Population")
            worksheet.write_string("D5", "%")
            worksheet.write_string("E5", "Total")
            worksheet.write_string("F5", "5000+")
            worksheet.write_string("G5", "3000->5000")
            worksheet.write_string("H5", "1000->3000")
            worksheet.write_string("I5", "500->1000")
            worksheet.write_string("J5", "300->500")
            worksheet.write_string("K5", "300->5000")
            # third line of table
            worksheet.write_string("B6", "Total Count")
            # fourth line of table
            worksheet.write_string("B7", "Total Pops")
            # fifth line of table
            worksheet.write_string("B8", "Existing CPOPS")
            # sixth line of table
            worksheet.write_string("B9", "Total 4G CPOPS")
            # seventh line of table
            worksheet.write_string("B10", "Total 3G CPOPS")
            # eighth line of table
            worksheet.write_string("B11", "Total Fixed/WIFI CPOPS")
            # ninth line of table
            worksheet.set_row(12, None, None, {'collapsed': 1, 'hidden': True})
            # tenth line of table
            worksheet.write_string("B13", "3G-Only CPOPS")
            # eleventh line of table
            worksheet.write_string("B14", "2G-Only CPOPS")
            # twelfth line of table
            worksheet.write_string("B15", "Uncovered POPs")
            # thirteenth line of table
            worksheet.write_string("B16", "Total Opportunity POPs")
        elif table_type == 3:
            # first line of table
            worksheet.write_string("B4", "Capex/cpop Summary")
            worksheet.write_string("C4", "Total")
            worksheet.write_string("D4", "<$10/cpop")
            worksheet.write_string("E4", "$10<$20/cpop")
            worksheet.write_string("F4", "$20<$40/cpop")
            worksheet.write_string("G4", "$40<$60/cpop")
            worksheet.write_string("H4", "$60<$80/cpop")
            worksheet.write_string("I4", ">$80/cpop")
            # second to thirteen line of table
            suffixes = ["Opportunity POPs", "RAN CPOPs", "Sites"]
            categories = ["Total", "Greenfield", "2G Overlay", "3G Overlay"]
            row = 5
            for suffix in suffixes:
                for category in categories:
                    worksheet.write_string("B{0}".format(row), "{0} {1}".format(category, suffix))
                    row += 1
            # fourteenth to eighteenth line of table
            worksheet.write_string("B{}".format(row), "Capex/cpop")
            row += 1
            worksheet.write_string("B{}".format(row), "Capex/site")
            row += 1
            suffixes = ["Capex/site"]
            categories = ["Greenfield", "2G Overlay", "3G Overlay"]
            for suffix in suffixes:
                for category in categories:
                    worksheet.write_string("B{0}".format(row), "{0} {1}".format(category, suffix))
                    row += 1
            # ninteenth to 22th line of table
            worksheet.write_string("B{}".format(row), "Total CapEx")
            row += 1
            prefixes = ["Total CapEx"]
            categories = ["Greenfield Sites", "2G Overlay", "3G Overlay"]
            for prefix in prefixes:
                for category in categories:
                    worksheet.write_string("B{0}".format(row), "{0}- {1}".format(prefix, category))
                    row += 1
            
    def setTableBorder(self, worksheet, table_type):
        """Set Table Border
        
        Args:
            table_type: 1: Competition Analysis
                        2: Country TAM Analysis-PER
                        3: Network Analysis
        """
        border_format = self.workbook.add_format({
                            'border':1,
                            'align': "right", 
                            'font_size':10
                           })
        cell_format = { 'type' : 'no_blanks' , 'format' : border_format}
        if table_type == 1:
            worksheet.conditional_format("B5:L11", cell_format)
        elif table_type == 2:
            worksheet.conditional_format("B4:K16", cell_format)
        elif table_type == 3:
            worksheet.conditional_format("B4:I25", cell_format)
            
    def writeFormulaToCell(self, worksheet, cell, formula):
        """Write Formula to Cell
        
        """
        worksheet.write_formula(cell, formula)
        
    def competitionAnalysis(self, worksheet_name, base_sheet_name):
        """Competition Analysis
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "C1:L1", "COMPETITION ANALYSIS")
        self.mergeCellsAndWrite(worksheet, "A2:M2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=1)
        # set border
        self.setTableBorder(worksheet, table_type=1)
        # insert formula
        self.writeFormulaToCell(worksheet, "C7", "=SUM('{}'!F2:F8)".format(base_sheet_name))
        
    def countryTAMAnalysisPER(self, worksheet_name, base_sheet_name):
        """Country TAM Analysis-PER
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "C1:L1", "Country TAM Analysis-PER")
        self.mergeCellsAndWrite(worksheet, "A2:M2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=2)
        # set border
        self.setTableBorder(worksheet, table_type=2)
        # insert formula
        self.writeFormulaToCell(worksheet, "C6", "=SUM('{}'!F2:F8)".format(base_sheet_name))
        
    def networkAnalysis(self, worksheet_name, base_sheet_name, partner_name):
        """Network Analysis
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "E1:I1", "Network Analysis: Partner = {}".format(partner_name))
        self.mergeCellsAndWrite(worksheet, "A2:M2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=3)
        # set border
        self.setTableBorder(worksheet, table_type=3)
        # insert formula
        self.writeFormulaToCell(worksheet, "C5", "=SUM('{}'!F2:F8)".format(base_sheet_name))