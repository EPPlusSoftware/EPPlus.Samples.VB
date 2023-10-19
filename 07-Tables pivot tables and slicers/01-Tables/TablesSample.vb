Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Table
Imports System
Imports System.Data.SQLite
Imports System.Drawing
Imports System.Threading.Tasks

Namespace EPPlusSamples.TablesPivotTablesAndSlicers
    ''' <summary>
    ''' This sample demonstrates how work with Excel tables in EPPlus.
    ''' Tables can easily be added by many of the ExcelRange - Load methods as demonstrated in earlier samples.
    ''' This sample will focus on how to add and setup tables from the ExcelWorksheet.Tables collection.
    ''' </summary>
    Public Module TablesSample
        Public Async Function RunAsync() As Task
            Using p = New ExcelPackage()
                Await CreateTableWithACalculatedColumnAsync(p).ConfigureAwait(False)
                Await StyleTablesAsync(p).ConfigureAwait(False)
                Await CreateTableFilterAndSlicerAsync(p).ConfigureAwait(False)

                p.SaveAs(FileUtil.GetCleanFileInfo("7.1-Tables.xlsx"))
            End Using
        End Function
        ''' <summary>
        ''' This sample creates a table with a calculated column. A totals row is added and styling is applied to some of the columns.
        ''' </summary>
        ''' <paramname="connectionString">The connection string to the database</param>
        ''' <paramname="p">The package</param>
        ''' <returns></returns>
        Private Async Function CreateTableWithACalculatedColumnAsync(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("SimpleTable")

            Dim range = Await LoadDataAsync(ws).ConfigureAwait(False)
            Dim tbl1 = ws.Tables.Add(range, "Table1")

            tbl1.ShowTotal = True
            'Format the OrderDate column and add a Count Numbers subtotal.
            tbl1.Columns("OrderDate").TotalsRowFunction = RowFunctions.CountNums
            tbl1.Columns("OrderDate").DataStyle.NumberFormat.Format = "yyyy-MM-dd"
            tbl1.Columns("OrderDate").TotalsRowStyle.NumberFormat.Format = "#,##0"

            'Format the OrderValue column and add a Sum subtotal.
            tbl1.Columns("OrderValue").TotalsRowFunction = RowFunctions.Sum
            tbl1.Columns("OrderValue").DataStyle.NumberFormat.Format = "#,##0"
            tbl1.Columns("OrderValue").TotalsRowStyle.NumberFormat.Format = "#,##0"

            'Adds a calculated formula referencing the OrderValue column within the same row.
            tbl1.Columns.Add(1)
            Dim addedcolumn = tbl1.Columns(tbl1.Columns.Count - 1)
            addedcolumn.Name = "OrderValue with Tax"
            addedcolumn.CalculatedColumnFormula = "Table1[[#This Row],[OrderValue]] * 110%" 'Sets the calculated formula referencing the OrderValue column within this row.
            addedcolumn.TotalsRowFunction = RowFunctions.Sum
            addedcolumn.DataStyle.NumberFormat.Format = "#,##0"
            addedcolumn.TotalsRowStyle.NumberFormat.Format = "#,##0"

            tbl1.ShowLastColumn = True

            tbl1.Range.AutoFitColumns()

            'Calculate the formulas so we get the calculated column values as well.
            ws.Calculate()

            'Create a data table from the table
            Dim dataTable = tbl1.ToDataTable(Sub(x)
                                                 x.DataTableName = "DataTable1"
                                                 x.SkipNumberOfRowsEnd = 2
                                             End Sub)
            'Then create a new table from the data table
            Dim range2 = ws.Cells("K1").LoadFromDataTable(dataTable, True, TableStyles.Dark4)
            Dim tbl2 = ws.Tables.GetFromRange(range2)

            'Format the OrderDate column and add a Count Numbers subtotal.
            tbl2.Columns("OrderDate").TotalsRowFunction = RowFunctions.CountNums
            tbl2.Columns("OrderDate").DataStyle.NumberFormat.Format = "yyyy-MM-dd"
            tbl2.Columns("OrderDate").TotalsRowStyle.NumberFormat.Format = "#,##0"

            'Format the OrderValue column and add a Sum subtotal.
            tbl2.Columns("OrderValue").TotalsRowFunction = RowFunctions.Sum
            tbl2.Columns("OrderValue").DataStyle.NumberFormat.Format = "#,##0"
            tbl2.Columns("OrderValue").TotalsRowStyle.NumberFormat.Format = "#,##0"

            range2.AutoFitColumns()
        End Function
        ''' <summary>
        ''' This sample creates a two table and a custom table style. The first table is styled using different style objects of the table. 
        ''' The second table is styled using the custom table style
        ''' </summary>
        ''' <paramname="connectionString"></param>
        ''' <paramname="p"></param>
        ''' <returns></returns>
        Private Async Function StyleTablesAsync(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("StyleTables")

            Dim range1 = Await LoadDataAsync(ws).ConfigureAwait(False)
            SetEmailAsHyperlink(range1)
            'Add the table and set some styles and properties.
            Dim tbl1 = ws.Tables.Add(range1, "StyleTable1")
            tbl1.TableStyle = TableStyles.Medium24
            tbl1.DataStyle.Font.Size = 10
            tbl1.Columns("E-Mail").DataStyle.Font.Underline = Style.ExcelUnderLineType.Single
            tbl1.HeaderRowStyle.Font.Italic = True
            tbl1.ShowTotal = True
            tbl1.TotalsRowStyle.Font.Italic = True
            tbl1.Range.Style.Font.Name = "Arial"
            tbl1.Range.AutoFitColumns()

            'Add two rows at the end.
            Dim addedRange = tbl1.AddRow(2)
            addedRange.Offset(0, 0, 1, 1).Value = "Added Row 1"
            addedRange.Offset(1, 0, 1, 1).Value = "Added Row 2"

            'Add a custom formula to display number of items in the CompanyName column
            tbl1.Columns(0).TotalsRowFormula = """Total Count is "" & SUBTOTAL(103,StyleTable1[CompanyName])"
            tbl1.Columns(0).TotalsRowStyle.Font.Color.SetColor(Color.Red)

            'We create a custom named style via the Workbook.Styles object. For more samples on custom styles see sample 27
            Dim customStyleName = "EPPlus Created Style"
            Dim customStyle = p.Workbook.Styles.CreateTableStyle(customStyleName, TableStyles.Medium13)
            customStyle.HeaderRow.Style.Font.Color.SetColor(eThemeSchemeColor.Text1)
            customStyle.FirstColumn.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent5)
            customStyle.FirstColumn.Style.Fill.BackgroundColor.Tint = 0.3
            customStyle.FirstColumn.Style.Font.Color.SetColor(eThemeSchemeColor.Text1)

            Dim range2 = Await LoadDataAsync(ws, "K1").ConfigureAwait(False)
            Dim tbl2 = ws.Tables.Add(range2, "StyleTable2")
            'To apply the custom style we set the StyleName property to the name we choose for our style.
            tbl2.StyleName = customStyleName
            tbl2.ShowFirstColumn = True

            tbl2.Range.AutoFitColumns()
        End Function

        Private Sub SetEmailAsHyperlink(ByVal range As ExcelRangeBase)
            For row = 1 To range.Rows
                Dim cell = range.Offset(row, 2, 1, 1)
                If cell.Value IsNot Nothing Then
                    cell.Hyperlink = New Uri($"mailto:{cell.Value}")
                End If
            Next
        End Sub

        ''' <summary>
        ''' This sample creates a table and a slicer. 
        ''' </summary>
        ''' <paramname="connectionString">The connection string to the database</param>
        ''' <paramname="p">The package</param>
        ''' <returns></returns>
        Private Async Function CreateTableFilterAndSlicerAsync(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("Slicer")

            Dim range = Await LoadDataAsync(ws).ConfigureAwait(False)
            Dim tbl = ws.Tables.Add(range, "FilterTable1")
            tbl.TableStyle = TableStyles.Medium1

            'Add a slicer and filter on company name. A table slicer is connected to a table columns value filter.
            Dim slicer1 = tbl.Columns(0).AddSlicer()
            slicer1.FilterValues.Add("Cremin-Kihn")
            slicer1.FilterValues.Add("Senger LLC")
            range.AutoFitColumns()

            'Apply the column filter, otherwise the slicer may be hidden when the filter is applied.
            tbl.AutoFilter.ApplyFilter()
            slicer1.SetPosition(2, 0, 10, 0)

            'For more samples on filters and slicers see Sample 13 and 24.
        End Function
        Private Async Function LoadDataAsync(ByVal ws As ExcelWorksheet, ByVal Optional startCell As String = "A1") As Task(Of ExcelRangeBase)
            Dim range As ExcelRangeBase
            'Lets connect to the sample database for some data
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        range = Await ws.Cells(startCell).LoadFromDataReaderAsync(sqlReader, True)
                    End Using
                End Using
                sqlConn.Close()
            End Using
            Return range
        End Function
    End Module
End Namespace
