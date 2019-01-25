from win32com import client
from enum import IntEnum
from datetime import datetime

class TableType(IntEnum):
    ttAggregated = 0
    ttProfile = 1

class XlDirection(IntEnum):
    xlDown = -4121
    xlToLeft = -4159
    xlToRight = -4161
    xlUp = -4162

class Constants(IntEnum):
    xlLeft = -4131

class Formatter:

    PROJECT_ANNOTATION = 0
    TABLE_NAME_ANNOTATION = 1
    BASE_ANNOTATION = 2
    RUN_ANNOTATION = 3

    def __init__(self, mtd_path, xls_path):
        self.mtd_path = mtd_path
        self.xls_path = xls_path
        self.in_context_manager = False

    def format(self):

        if not self.in_context_manager:
            print('please use inside context manager')
            return

        for t in self.tables:
            t.format()

    def create_toc(self):

        if not self.in_context_manager:
            print('please use inside context manager')
            return

        project_name = self.mtd.Tables[0].Annotations[Formatter.PROJECT_ANNOTATION]
        run_name = self.mtd.Tables[0].Annotations[Formatter.RUN_ANNOTATION]
        
        if self.xls.Worksheets[0].Name == 'Content':
            self.xls.Worksheets[0].Delete
        
        self.content_ws = self.xls.Worksheets.Add(Before=self.xls.Worksheets[0])
        self.content_ws.Name = 'Content'
    
        # write header
        header_range = self.content_ws.Range('B2:C4')
        header_range.Cells(1, 1).Value = 'Project:'
        header_range.Cells(2, 1).Value = 'Wave:'
        header_range.Cells(3, 1).Value = 'Date:'
        header_range.Cells(1, 2).Value = project_name
        header_range.Cells(2, 2).Value = run_name
        header_range.Cells(3, 2).Value = datetime.now()

        # format header
        header_range.Columns[0].Font.Name = 'Arial'
        header_range.Columns[0].Font.Size = 12
        header_range.Columns[0].Font.Italic = True
        header_range.Columns[0].Font.Bold = True
        header_range.Columns[1].HorizontalAlignment = Constants.xlLeft

        # write tables header
        self.content_ws.Cells(6, 2).Value = 'Sheet'
        self.content_ws.Cells(6, 3).Value = 'Question'
        self.content_ws.Cells(6, 4).Value = 'Base'
            
        current_row = 7

        for t in self.tables:
            
            self.content_ws.Cells(current_row, 2).Value = t.final_name
            self.content_ws.Hyperlinks.Add(
                Anchor=self.content_ws.Cells(current_row, 2),
                Address='',
                SubAddress=f"'{t.final_name}'!A1",
                ScreenTip='Go to Table',
                TextToDisplay=t.final_name
            )
            t.xl_worksheet.Hyperlinks.Add(
                Anchor=t.back_to_content_cell,
                Address='',
                SubAddress='Content!A1'
            )
            t.back_to_content_cell.Value = '<< Back to Content'
          
            self.content_ws.Cells(current_row, 3).Value = t.table_description
            self.content_ws.Cells(current_row, 4).Value = t.base_description

            current_row += 1
            
        self.content_ws.Range('A:A').ColumnWidth = 3
        self.content_ws.Range('B:B').EntireColumn.AutoFit
        self.content_ws.Range('C:C').ColumnWidth = 70
        self.content_ws.Range('C:C').EntireColumn.WrapText = True
        self.content_ws.Range('D:D').EntireColumn.AutoFit()
        
        self.content_ws.Activate()
        self.xl_app.ActiveWindow.FreezePanes = False
        self.content_ws.Cells(7, 1).Select()
        self.xl_app.ActiveWindow.FreezePanes = True

    def __enter__(self):

        self.in_context_manager = True

        self.xl_app = client.Dispatch('Excel.Application')
        self.xl_app.DisplayAlerts = False
        # self.xl_app.Visible = False

        self.xls = self.xl_app.Workbooks.Open(self.xls_path)
        self.mtd = client.Dispatch('TOM.Document')
        self.mtd.Open(self.mtd_path)
        self.tables = []
        idx = 0
        for t in self.mtd.Tables:
            xl = self.xls.Sheets[idx]
            self.tables.append(Table(t, xl))
            idx += 1
        
        return self

    def __exit__(self, exception_type, exception_value, traceback):

        # saving, closing and cleaning up
        self.xls.Save()
        self.xls.Close()
        del self.xls

        self.xl_app.DisplayAlerts = True
        # self.xl_app.Visible = True
        del self.xl_app

        self.mtd.Clear()
        del self.mtd

        self.in_context_manager = False

class Table:

    def __init__(self, mtd_table, ws):
        self.mtd_table = mtd_table
        self.xl_worksheet = ws

        self.initial_name = mtd_table.Name

        # calculates header size
        self.header_size = 0
        for i in range(3): # checks only top annotations
            annotation_text = self.mtd_table.Annotations[i].Text
            if annotation_text:
                self.header_size += len(annotation_text.split('<BR/>'))

        # get axis depth
        self.side_axis_depth = self._get_axis_depth(self.mtd_table.Axes['Side'])
        self.top_axis_depth = 0
        if len(self.mtd_table.Axes) > 1:
            self.top_axis_depth = self._get_axis_depth(self.mtd_table.Axes['Top'])

        # autofit ranges: last column in side axis and last row in top axis
        self.autofit_columns = self.xl_worksheet.Columns(self.side_axis_depth * 2)
        if self.mtd_table.Type == TableType.ttAggregated:
            self.autofit_rows = self.xl_worksheet.Rows(self.header_size + 1 + self.top_axis_depth * 2)
        elif self.mtd_table.Type == TableType.ttProfile:
            self.autofit_rows = self.xl_worksheet.Rows(self.header_size + 3)

        # freeze cell: includes first data column (total)
        # and first data row (base) for percent only tables and 2 first data rows (base) for percent/absolute tables
        if self.mtd_table.Type == TableType.ttAggregated:
            self.freeze_cell = self.xl_worksheet.Cells(self.autofit_rows[0].Row + 
                self.mtd_table.CellItems.Count + 1, self.autofit_columns[0].Column + 2)
        elif self.mtd_table.Type == TableType.ttProfile:
            self.freeze_cell = self.xl_worksheet.Cells(self.autofit_rows[0].Row + 1, 
                self.autofit_columns[0].Column + 1)

        # hyperlink cell: cell 2 rows below the header
        self.xl_worksheet.Rows(self.header_size + 1).Insert(Shift=XlDirection.xlDown)
        self.xl_worksheet.Rows(self.header_size + 1).Insert(Shift=XlDirection.xlDown)
        self.back_to_content_cell = self.xl_worksheet.Cells(self.header_size + 2, 1)
        
        # table & base descriptions
        self.table_description = self.mtd_table.Annotations[Formatter.TABLE_NAME_ANNOTATION].Text
        self.base_description = self.mtd_table.Annotations(Formatter.BASE_ANNOTATION).Text

    def _get_axis_depth(self, axis):

            # traverses all the nodes of the tree (breadth-first)
            # increments a counter for each new level
            axis_depth = 0

            # initializes axes collection for the top-level axis
            axes = [a for a in axis.SubAxes]

            # Loops through the tree level by level
            while True:
                axis_depth += 1

                # reassigns axes to subaxes of all axes
                axes = [sub for a in axes for sub in a.SubAxes]

                # exits if axes collection is empty
                if not axes:
                    return axis_depth

    def format(self):

        # renames worksheet
        self.final_name = self.rename_worksheet()
        self.xl_worksheet.Name = self.final_name
        
        # auto fits columns/rows
        for c in self.autofit_columns.Columns:
            c.AutoFit()
        for r in self.autofit_rows.Rows:
            r.AutoFit()
        
        # freezing cells
        self.xl_worksheet.Activate()
        active_window = self.xl_worksheet.Application.ActiveWindow
        if active_window.FreezePanes:
            active_window.FreezePanes = False
        self.freeze_cell.Select()
        active_window.FreezePanes = True
        self.xl_worksheet.Cells(1, 1).Select()
    
    def rename_worksheet(self):
        temp_name = self.initial_name

        #removes illegal characters from sheetname
        illegal_characters = '/\\[]*?:'
        for c in illegal_characters:
            temp_name = temp_name.replace(c, '')

        # truncates string to 31 characters (max allowed in excel)
        MAX_SHEET_NAME_LENGTH = 31
        temp_name = temp_name[:MAX_SHEET_NAME_LENGTH + 1]        

        # checks if modified worksheet name already exist in the workbook
        # if yes, increments name by 1 and tries again until no duplicates are found
        workbook = self.xl_worksheet.Parent
        counter = 1
        while True:
            current_name = temp_name if counter == 1 else f'{temp_name} ({counter})'
            for ws in workbook.Worksheets:
                if ws.Name == current_name:
                    counter += 1
                    break
            else:
                return current_name