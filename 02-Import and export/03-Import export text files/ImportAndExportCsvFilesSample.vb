' ***********************************************************************************************
' Required Notice: Copyright (C) EPPlus Software AB. 
' This software is licensed under PolyForm Noncommercial License 1.0.0 
' and may only be used for noncommercial purposes 
' https://polyformproject.org/licenses/noncommercial/1.0.0/
' 
' A commercial license to use this software can be purchased at https://epplussoftware.com
' ************************************************************************************************
' Date               Author                       Change
' ************************************************************************************************
' 01/27/2020         EPPlus Software AB           Initial release EPPlus 5
' ***********************************************************************************************
Imports System
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports OfficeOpenXml.Drawing.Chart
Imports System.Globalization
Imports System.Threading.Tasks
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.LoadDataFromCsvFilesIntoTables
    ''' <summary>
    ''' This sample shows how to load/save CSV files using the LoadFromText and SaveToText methods, how to use tables and
    ''' how to use charts with more than one chart type and secondary axis
    ''' </summary>
    Public Module ImportAndExportCsvFilesSample
        ''' <summary>
        ''' Loads two CSV files into tables and adds a chart to each sheet.
        ''' </summary>
        ''' <paramname="outputDir"></param>
        ''' <returns></returns>
        Public Async Function RunAsync() As Task
            Console.WriteLine("Running sample 2.3")
            Dim newFile = FileUtil.GetCleanFileInfo("2.3-LoadDataFromCsvFilesIntoTables.xlsx")

            Using package As ExcelPackage = New ExcelPackage()
                LoadFile1(package)                 'Load the text file without async
                Await LoadFile2Async(package)      'Load the second text file with async
                Await ExportTableAsync(package)
                Await package.SaveAsAsync(newFile)
            End Using
            Console.WriteLine("Sample 2.3 created: {0}", newFile.FullName)
            Console.WriteLine()
        End Function

        Private Async Function ExportTableAsync(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets(1)
            Dim tbl = ws.Tables(0)
            Dim format = New ExcelOutputTextFormat With {
    .Delimiter = ";"c,
    .Culture = New CultureInfo("en-GB"),
    .Encoding = New UTF8Encoding(),
    .SkipLinesEnd = 1  'Skip the totals row                
}
            Await ws.Cells(tbl.Address.Address).SaveToTextAsync(FileUtil.GetCleanFileInfo("3.1-ExportedFromEPPlus.csv"), format)

            Console.WriteLine($"Writing the text file 'ExportedTable.csv'...")
        End Function

        Private Sub LoadFile1(ByVal package As ExcelPackage)
            'Create the Worksheet
            Dim sheet = package.Workbook.Worksheets.Add("Csv1")

            'Create the format object to describe the text file
            Dim format = New ExcelTextFormat With {
    .EOL = vbLf,
    .TextQualifier = """"c,
    .SkipLinesBeginning = 2,
    .SkipLinesEnd = 1
}

            Dim file1 = FileUtil.GetFileInfo("02-Import and Export\03-Import export text files", "Sample2.3-1.txt")

            'Now read the file into the sheet. Start from cell A1. Create a table with style 27. First row contains the header.
            Console.WriteLine("Load the text file...")
            Dim range = sheet.Cells("A1").LoadFromText(file1, format, TableStyles.Medium27, True)

            Console.WriteLine("Format the table...")
            'Tables don't support custom styling at this stage(you can of course format the cells), but we can create a Namedstyle for a column...
            Dim dateStyle = package.Workbook.Styles.CreateNamedStyle("TableDate")
            dateStyle.Style.Numberformat.Format = "YYYY-MM"

            Dim numStyle = package.Workbook.Styles.CreateNamedStyle("TableNumber")
            numStyle.Style.Numberformat.Format = "#,##0.0"

            'Now format the table...
            Dim tbl = sheet.Tables(0)
            tbl.ShowTotal = True
            tbl.Columns(0).TotalsRowLabel = "Total"
            tbl.Columns(0).DataCellStyleName = "TableDate"
            tbl.Columns(1).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(1).DataCellStyleName = "TableNumber"
            tbl.Columns(2).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(2).DataCellStyleName = "TableNumber"
            tbl.Columns(3).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(3).DataCellStyleName = "TableNumber"
            tbl.Columns(4).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(4).DataCellStyleName = "TableNumber"
            tbl.Columns(5).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(5).DataCellStyleName = "TableNumber"
            tbl.Columns(6).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(6).DataCellStyleName = "TableNumber"

            Console.WriteLine("Create the chart...")
            'Now add a stacked areachart...
            Dim chart = sheet.Drawings.AddChart("chart1", eChartType.AreaStacked)
            chart.SetPosition(0, 630)
            chart.SetSize(800, 600)

            'Create one series for each column...
            For col = 1 To 6
                Dim ser = chart.Series.Add(range.Offset(1, col, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1))
                ser.HeaderAddress = range.Offset(0, col, 1, 1)
            Next

            'Set the style to predefied style 27. You can also use the chart.StyleManager.SetChartStyle method to set more modern styles. See for example the csv2 sheet in this sample. 
            chart.Style = eChartStyle.Style27

            sheet.View.ShowGridLines = False
            sheet.Calculate()
            sheet.Cells(sheet.Dimension.Address).AutoFitColumns()
        End Sub

        Private Async Function LoadFile2Async(ByVal package As ExcelPackage) As Task
            'Create the Worksheet
            Dim sheet = package.Workbook.Worksheets.Add("Csv2")

            'Create the format object to describe the text file
            Dim format = New ExcelTextFormat With {
    .EOL = vbLf,
    .Delimiter = ChrW(9),       'Tab
    .SkipLinesBeginning = 1
}
            Dim ci As CultureInfo = New CultureInfo("sv-SE")          'Use your choice of Culture
            ci.NumberFormat.NumberDecimalSeparator = ","       'Decimal is comma
            format.Culture = ci

            'Now read the file into the sheet.
            Console.WriteLine("Load the text file...")
            Dim file2 = FileUtil.GetFileInfo("02-Import and Export\03-Import export text files", "Sample2.3-2.txt")

            Dim range = Await sheet.Cells("A1").LoadFromTextAsync(file2, format)

            'Add a formula
            range.Offset(1, range.End.Column, range.End.Row - range.Start.Row, 1).FormulaR1C1 = "RC[-1]-RC[-2]"

            'Add a table...
            Dim tbl = sheet.Tables.Add(range.Offset(0, 0, range.End.Row - range.Start.Row + 1, range.End.Column - range.Start.Column + 2), "Table")
            tbl.ShowTotal = True
            tbl.Columns(0).TotalsRowLabel = "Total"
            tbl.Columns(1).TotalsRowFormula = "COUNT(3,Table[Product])"    'Add a custom formula
            tbl.Columns(2).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(3).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(4).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(5).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(5).Name = "Profit"
            tbl.TableStyle = TableStyles.Medium10

            'To the header row and totals font to italic, use the HeaderRowStyle and the TotalsRowStyle property. You can also use the tbl.DataStyle to style the data part of the table.
            tbl.HeaderRowStyle.Font.Italic = True
            tbl.TotalsRowStyle.Font.Italic = True

            sheet.Cells(sheet.Dimension.Address).AutoFitColumns()

            'Add a chart with two charttypes (Column and Line) and a secondary axis...
            Dim chart = sheet.Drawings.AddChart("chart2", eChartType.ColumnStacked)
            chart.SetPosition(0, 540)
            chart.SetSize(800, 600)

            Dim serie1 = chart.Series.Add(range.Offset(1, 3, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1))
            serie1.Header = "Purchase Price"
            Dim serie2 = chart.Series.Add(range.Offset(1, 5, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1))
            serie2.Header = "Profit"

            'Add a Line series
            Dim chartType2 = chart.PlotArea.ChartTypes.Add(eChartType.LineStacked)
            chartType2.UseSecondaryAxis = True
            Dim serie3 = chartType2.Series.Add(range.Offset(1, 2, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1))
            serie3.Header = "Items in stock"

            'By default the secondary XAxis is not visible, but we want to show it...
            chartType2.XAxis.Deleted = False
            chartType2.XAxis.TickLabelPosition = eTickLabelPosition.High

            'Set the max value for the Y axis...
            chartType2.YAxis.MaxValue = 50

            chart.StyleManager.SetChartStyle(ePresetChartStyle.ComboChartStyle2)

            sheet.View.ShowGridLines = False
            sheet.Calculate()
        End Function
    End Module
End Namespace
