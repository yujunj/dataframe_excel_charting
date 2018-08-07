'''
Created on Aug 5, 2018

@author: arthur
'''
# import xlsxwriter
from xlsxwriter import worksheet
from xlsxwriter import utility 

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
        
    def writeStringsToMultipleCells(self, worksheet, cell_range, input_strings_list, axis=0):
        """write Strings to Multiple Cells
        
        Args:
            cell_range: "A1:A10"
            input_strings_list: ["asd", "awasd", ..., "nju"]
            axis: 
                0 means write to next row first (default)
                1 means write to next column first
            
        Usage:
            self.writeStringsToMultipleCells(
                worksheet, 
                "C6:G6", 
                ["Total CPOPs", "4G", "3G+4G", "2G Only", "Unconnected"], 
                axis=1
            )
        
        """
        # get number from cell range
        (start_row, start_col), (end_row, end_col) = self.getNumberFromCellsRange(cell_range)
        # strings counter 
        i = 0
        # next row first
        if axis == 0:
            for col in range(start_col, end_col + 1):
                for row in range(start_row, end_row + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    worksheet.write_string(cell, input_strings_list[i])
                    i += 1
        # next column first
        elif axis == 1:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    worksheet.write_string(cell, input_strings_list[i])
                    i += 1
    
    def writeNumbersToMultipleCells(self, worksheet, cell_range, input_numbers_list, axis=0):
        """write Numbers to Multiple Cells
        
        Args:
            cell_range: "A1:A10"
            input_numbers_list: [0, 1, ..., 10]
            axis: 
                0 means write to next row first (default)
                1 means write to next column first
            
        Usage:
            self.writeStringsToMultipleCells(
                worksheet, 
                "C6:G6", 
                ["Total CPOPs", "4G", "3G+4G", "2G Only", "Unconnected"], 
                axis=1
            )
        
        """
        # get number from cell range
        (start_row, start_col), (end_row, end_col) = self.getNumberFromCellsRange(cell_range)
        # strings counter 
        i = 0
        # next row first
        if axis == 0:
            for col in range(start_col, end_col + 1):
                for row in range(start_row, end_row + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    worksheet.write_string(cell, input_numbers_list[i])
                    i += 1
        # next column first
        elif axis == 1:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    worksheet.write_string(cell, input_numbers_list[i])
                    i += 1
    
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
            self.writeStringsToMultipleCells(
                worksheet, 
                "C6:G6", 
                ["Total CPOPs", "4G", "3G+4G", "2G Only", "Unconnected"], 
                axis=1
            )
            self.writeStringsToMultipleCells(
                worksheet, 
                "H6:L6", 
                ["Total", "4G", "3G+4G", "2G Only", "Unconnected"], 
                axis=1
            )
        elif table_type == 2:
            # first line of table
            self.mergeCellsAndWrite(worksheet, "C4:D4", "PER")
            self.mergeCellsAndWrite(worksheet, "E4:K4", "Settlements")
            # second line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "C5:K5", 
                ["Population", "%", "Total", "5000+", "3000->5000", 
                 "1000->3000", "500->1000", "300->500", "300->5000"], 
                axis=1
            )
            # third to eighth line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B6:B11", 
                ["Total Count", "Total Pops", "Existing CPOPS", 
                 "Total 4G CPOPS", "Total 3G CPOPS", "Total Fixed/WIFI CPOPS"], 
                axis=0
            )
            # ninth line of table
            # worksheet.set_row(11, None, None, {'collapsed': 1, 'hidden': True})
            # tenth to thirteenth line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B12:B15", 
                ["3G-Only CPOPS", "2G-Only CPOPS", "Uncovered POPs", "Total Opportunity POPs"], 
                axis=0
            )
        elif table_type == 3:
            # first line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B4:I4", 
                ["Capex/cpop Summary", "Total", "<$10/cpop", "$10<$20/cpop", "$20<$40/cpop", 
                 "$40<$60/cpop", "$60<$80/cpop", ">$80/cpop"], 
                axis=1
            )
            # second to thirteen line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B5:B16", 
                ['Total Opportunity POPs', 'Greenfield Opportunity POPs', '2G Overlay Opportunity POPs', '3G Overlay Opportunity POPs', 
                 'Total RAN CPOPs', 'Greenfield RAN CPOPs', '2G Overlay RAN CPOPs', '3G Overlay RAN CPOPs', 
                 'Total Sites', 'Greenfield Sites', '2G Overlay Sites', '3G Overlay Sites'], 
                axis=0
            )
            # fourteenth to eighteenth line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B17:B21", 
                ["Capex/cpop", "Capex/site", 
                 "Greenfield Capex/site", "2G Overlay Capex/site", "3G Overlay Capex/site"], 
                axis=0
            )
            # ninteenth to 22th line of table
            self.writeStringsToMultipleCells(
                worksheet, 
                "B22:B25", 
                ["Total CapEx", "Total CapEx- Greenfield Sites", "Total CapEx- 2G Overlay", "Total CapEx- 3G Overlay"], 
                axis=0
            )

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
            worksheet.conditional_format("B1:L11", cell_format)
        elif table_type == 2:
            worksheet.conditional_format("B1:K16", cell_format)
        elif table_type == 3:
            worksheet.conditional_format("B1:I25", cell_format)
            
    def writeFormulaToCell(self, worksheet, cell, formula):
        """Write Formula to Cell
        
        """
        worksheet.write_formula(cell, formula)
        
    def getNumberFromCellsRange(self,cell_range):
        """Get Number from Cells Range
        
        Args:
            cell_range: "B2:B10"
        
        Returns:
            tuple(int, int):
            (start_row, start_col), (end_row, end_col)
            
        """
        # parce the cell range input
        start_cell, end_cell = cell_range.split(":")
        (start_row, start_col) = utility.xl_cell_to_rowcol(start_cell)
        (end_row, end_col) = utility.xl_cell_to_rowcol(end_cell)
        return (start_row, start_col), (end_row, end_col)
        
    def writeColumnSum(self, worksheet, cell_range, columns_list, base_sheet_name, axis=0):
        """Write Column Sum
        
        Args:
            cell_range: "B2:B10"
            columns_list: ["B", "AS"]
            axis: 
                0 means write to next row first (default)
                1 means write to next column first
            
        Usage:
            self.writeColumnSum(
                worksheet, "C7:C8", ["B", "AS"], "data_sheet"
            )
        
        """
        # get number from cell range
        (start_row, start_col), (end_row, end_col) = self.getNumberFromCellsRange(cell_range)
        # column counter
        i = 0
        # write to next row first
        if axis == 0:
            for col in range(start_col, end_col + 1):
                for row in range(start_row, end_row + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    formula = "=SUM('{0}'!{1}:{1})".format(base_sheet_name, columns_list[i])
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
        # write to next column first
        elif axis == 1:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    formula = "=SUM('{0}'!{1}:{1})".format(base_sheet_name, columns_list[i])
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
                    
    def formatCountIfsCondition(self, column_criterion, base_sheet_name):
        """format count ifs condition
        
        Args:
            column_criteria = [
                {"AZ": ">0", "B": [">=3000", "<5000"]}, 
                {"A": "<10", "C": "<10"}
            ]
        """
        condition_list = []
        for key in column_criterion.keys():
            if isinstance(column_criterion[key], list):
                temp_str = map(lambda x: "'{0}'!${1}:${1},\"{2}\"".format(base_sheet_name, key, x), column_criterion[key])
                condition_list.extend(temp_str)
            else:
                temp_str = "'{0}'!${1}:${1},\"{2}\"".format(base_sheet_name, key, column_criterion[key])
                condition_list.append(temp_str)
        return ",".join(condition_list)

    def writeColumnCountIfs(self, worksheet, cell_range, columns_criteria, base_sheet_name, axis=0):
        """Write Column Count IFs
        
        Args:
            cell_range: "A1:D2"
            column_criteria: [
                {"AZ": ">0", "B": [">=3000", "<5000"]}, 
                {"A": "<10", "C": "<10"}
            ]
            axis: 
                0 means write to next row first (default)
                1 means write to next column first
            
        Usage:
            self.writeColumnCountIfs(
                worksheet, 
                "G8:G10", 
                [
                    {"D": ">0", "B": [">=3000", "<5000"]},
                    {"E": ">.25", "B": [">=3000", "<5000"]}, 
                    {"G": ">.25", "B": [">=3000", "<5000"]}
                ], 
                base_sheet_name, 
                axis=0
            )
        
        """
        # get number from cell range
        (start_row, start_col), (end_row, end_col) = self.getNumberFromCellsRange(cell_range)
        # loop through the column criteria
        i = 0
        # write to next row first
        if axis == 0:
            for col in range(start_col, end_col + 1):
                for row in range(start_row, end_row + 1):
                    column_criterion = columns_criteria[i]
                    cell = utility.xl_rowcol_to_cell(row, col)
                    # get condition string
                    condition = self.formatCountIfsCondition(column_criterion, base_sheet_name)
                    formula = "=COUNTIFS({})".format(condition)
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
        # write to next column first
        elif axis == 1:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    column_criterion = columns_criteria[i]
                    cell = utility.xl_rowcol_to_cell(row, col)
                    # get condition string
                    condition = self.formatCountIfsCondition(column_criterion, base_sheet_name)
                    formula = "=COUNTIFS({})".format(condition)
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
                    
    def writeCellDivisionDenomFixed(self, worksheet, cell_range, numer_list, denom_cell, base_sheet_name, axis=0):
        """Write Column Division, Fixed Denominator
        
        Args:
            cell_range: "A1:A10"
            numer_list: ["B1", "C2", ..., "H10"]
            denom_cell: "C7"
            axis: 
                0 means write to next row first (default)
                1 means write to next column first
        
        Usage:
            self.writeCellDivisionDenomFixed(
                worksheet, 
                "A1:A3", 
                ["B1", "B2", "B3"],
                "C8", 
                "data sheet", 
                axis=0
            )
        """
        # get number from cell range
        (start_row, start_col), (end_row, end_col) = self.getNumberFromCellsRange(cell_range)
        # loop through the numer list
        i = 0
        # write to next row first
        if axis == 0:
            for col in range(start_col, end_col + 1):
                for row in range(start_row, end_row + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    formula = "='{0}'!{1}/'{0}'!{2}".format(base_sheet_name, numer_list[i], denom_cell)
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
        # write to next column first
        elif axis == 1:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = utility.xl_rowcol_to_cell(row, col)
                    formula = "='{0}'!{1}/'{0}'!{2}".format(base_sheet_name, numer_list[i], denom_cell)
                    self.writeFormulaToCell(worksheet, cell, formula)
                    i += 1
            
    def competitionAnalysis(self, worksheet_name, base_sheet_name):
        """Competition Analysis
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "B1:L1", "COMPETITION ANALYSIS")
        self.mergeCellsAndWrite(worksheet, "B2:L2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=1)
        # set border
        self.setTableBorder(worksheet, table_type=1)
        # insert formula
        self.writeFormulaToCell(worksheet, "C7", "=SUM('{}'!F2:F8)".format(base_sheet_name))
        
    def countryTAMAnalysis(self, worksheet_name, base_sheet_name, country):
        """Country TAM Analysis
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "B1:K1", "Country TAM Analysis-{}".format(country))
        self.mergeCellsAndWrite(worksheet, "B2:K2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=2)
        # set border
        self.setTableBorder(worksheet, table_type=2)
        
        # insert formula sum
        self.writeColumnSum(
            worksheet, 
            "C7:C14", 
            ["B", "AS", "AW", "AX", "J", "AX", "AY", "AZ"], 
            base_sheet_name, 
            axis=0
        )
        # single cell insert 
        self.writeFormulaToCell(worksheet, "C15", "=C12+C13+C14")
        
        # insert division formula
        self.writeCellDivisionDenomFixed(
            worksheet, 
            "D8:D15", 
            ["C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15"],
            "C7", 
            base_sheet_name, 
            axis=0
        )
        # single column
        self.writeFormulaToCell(
            worksheet, "E6", "=COUNTA('{}'!A:A)-1".format(base_sheet_name)
        )
        # insert formula countifs
        self.writeColumnCountIfs(
            worksheet, 
            "E8:E10", 
            [{"AS": ">0"}, 
             {"E": ">0.25"}, 
             {"G": "0.25"}], 
            base_sheet_name, 
            axis=0
        )
        self.writeColumnCountIfs(
            worksheet, 
            "E12:E14", 
            [{"AV": "=3G"}, 
             {"AU": ">0.25"}, 
             {"D": "<0.25"}], 
            base_sheet_name, 
            axis=0
        )
        # single cell insert 
        self.writeFormulaToCell(worksheet, "E15", "=E12+E13+E14")
        # insert countifs
        self.writeColumnCountIfs(
            worksheet, 
            "G8:G10", 
            [{"D": ">0", "B": [">=3000", "<5000"]},
             {"E": ">.25", "B": [">=3000", "<5000"]}, 
             {"G": ">.25", "B": [">=3000", "<5000"]}], 
            base_sheet_name, 
            axis=0
        )
        self.writeColumnCountIfs(
            worksheet, 
            "G12:G14", 
            [{"AX": ">0", "B": [">=3000", "<5000"]}, 
             {"AY": ">0", "B": [">=3000", "<5000"]}, 
             {"AZ": ">0", "B": [">=3000", "<5000"]}], 
            base_sheet_name, 
            axis=0
        )
        
    def networkAnalysis(self, worksheet_name, base_sheet_name, partner_name):
        """Network Analysis
        
        """
        try:
            worksheet = self.openWorkSheet(worksheet_name)
        except AssertionError:
            worksheet = self.createWorkSheet(worksheet_name)
        # write title and note on 1st and 2nd rows
        self.mergeCellsAndWrite(worksheet, "B1:I1", "Network Analysis: Partner = {}".format(partner_name))
        self.mergeCellsAndWrite(worksheet, "B2:I2", "[SPACO:XXXX]")
        # write table frame
        self.writeTableFrame(worksheet, table_type=3)
        # set border
        self.setTableBorder(worksheet, table_type=3)
        # insert formula
        self.writeFormulaToCell(worksheet, "C5", "=SUM('{}'!F2:F8)".format(base_sheet_name))