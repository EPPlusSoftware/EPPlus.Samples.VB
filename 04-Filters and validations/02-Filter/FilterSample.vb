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
Imports OfficeOpenXml
Imports OfficeOpenXml.Filter
Imports OfficeOpenXml.Table.PivotTable
Imports System
Imports System.Data.SQLite
Imports System.Threading.Tasks

Namespace EPPlusSamples.FiltersAndValidations
    Public Class Filter
        Public Shared Async Function RunAsync() As Task
            Dim p = New ExcelPackage()

            'Autofilter on the worksheet
            Await ValueFilter(p)
            Await DateTimeFilter(p)
            Await CustomFilter(p)
            Await Top10Filter(p)
            Await DynamicAboveAverageFilter(p)
            Await DynamicDateAugustFilter(p)

            'Filter on a table, also see sample 24-Slicers. 
            Await TableFilter(p)

            'Filter on a pivot table, also see sample 24-Slicers. 
            Await PivotTableFilter(p)

            p.SaveAs(FileUtil.GetCleanFileInfo("4.2-Filters.xlsx"))
        End Function

        Private Shared Async Function ValueFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("ValueFilter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim colCompany = ws.AutoFilter.Columns.AddValueFilterColumn(0)
            colCompany.Filters.Add("Walsh LLC")
            colCompany.Filters.Add("Harber-Goldner")
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function DateTimeFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("DateTimeFilter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim col = ws.AutoFilter.Columns.AddValueFilterColumn(5)
            col.Filters.Add(New ExcelFilterDateGroupItem(2017, 8))
            col.Filters.Add(New ExcelFilterDateGroupItem(2017, 7, 5))
            col.Filters.Add(New ExcelFilterDateGroupItem(2017, 7, 7))
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function CustomFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("CustomFilter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim colCompany = ws.AutoFilter.Columns.AddCustomFilterColumn(6)
            colCompany.And = True
            colCompany.Filters.Add(New ExcelFilterCustomItem("999.99", eFilterOperator.GreaterThan))
            colCompany.Filters.Add(New ExcelFilterCustomItem("1500", eFilterOperator.LessThanOrEqual))
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function Top10Filter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("Top10Filter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim colTop10 = ws.AutoFilter.Columns.AddTop10FilterColumn(6)
            colTop10.Percent = False    'If set to true, the value takes top the percentage. Otherwise it relates to the number of items.
            colTop10.Value = 10         'The value to relate to.
            colTop10.Top = False        'Top if true, bottom if false
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function DynamicAboveAverageFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("DynamicAboveAverageFilter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim colDynamic = ws.AutoFilter.Columns.AddDynamicFilterColumn(6)
            colDynamic.Type = eDynamicFilterType.AboveAverage
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function DynamicDateAugustFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("DynamicAugustFilter")
            Dim range = Await LoadFromDatabase(ws)

            range.AutoFilter = True
            Dim colDynamic = ws.AutoFilter.Columns.AddDynamicFilterColumn(5)
            colDynamic.Type = eDynamicFilterType.M8
            ws.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function TableFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("TableFilter")
            Dim range = Await LoadFromDatabase(ws)

            Dim tbl = ws.Tables.Add(range, "tblFilter")
            tbl.TableStyle = Table.TableStyles.Medium23
            tbl.ShowFilter = True
            'Add a value filter
            Dim colCompany = tbl.AutoFilter.Columns.AddValueFilterColumn(0)
            colCompany.Filters.Add("Walsh LLC")
            colCompany.Filters.Add("Harber-Goldner")
            colCompany.Filters.Add("Sporer, Mertz and Jaskolski")

            'Add a second filter on order value
            Dim colOrderValue = tbl.AutoFilter.Columns.AddCustomFilterColumn(6)
            colOrderValue.Filters.Add(New ExcelFilterCustomItem("500", eFilterOperator.GreaterThanOrEqual))
            tbl.AutoFilter.ApplyFilter()
            range.AutoFitColumns(0)
        End Function
        Private Shared Async Function PivotTableFilter(ByVal p As ExcelPackage) As Task
            Dim ws = p.Workbook.Worksheets.Add("PivotTableFilter")
            Dim range = Await LoadFromDatabase(ws)

            Dim tbl = ws.Tables.Add(range, "ptFilter")
            tbl.TableStyle = Table.TableStyles.Medium23

            Dim pt1 = ws.PivotTables.Add(ws.Cells("J1"), tbl, "PivotTable1")
            Dim rowField = pt1.RowFields.Add(pt1.Fields("CompanyName"))
            Dim dataField = pt1.DataFields.Add(pt1.Fields("OrderValue"))

            'First deselect a company in the items list. To do so we first need to refresh the items from the range.
            rowField.Items.Refresh()  'Refresh the items from the range.
            rowField.Items.GetByValue("Roberts-Cruickshank").Hidden = True
            'Add a caption filter on Company Name between A and D
            rowField.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBetween, "A", "D")
            'Add a value filter where OrderValue >= 100
            rowField.Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThanOrEqual, dataField, 100)

            'Add a second pivot table with some different filters.
            Dim pt2 = ws.PivotTables.Add(ws.Cells("M1"), tbl, "PivotTable2")
            Dim rowField1 = pt2.RowFields.Add(pt2.Fields("Currency"))
            Dim rowField2 = pt2.RowFields.Add(pt2.Fields("OrderDate"))
            rowField2.Format = "yyyy-MM-dd"
            Dim dataField1 = pt2.DataFields.Add(pt2.Fields("OrderValue"))
            Dim dataField2 = pt2.DataFields.Add(pt2.Fields("OrderId"))
            dataField2.Function = DataFieldFunctions.CountNums

            Dim slicer = rowField1.AddSlicer()
            slicer.SetPosition(11, 0, 9, 0)
            'Add a date filter between first of Mars 2017 to 30th of June
            rowField2.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, New DateTime(2017, 3, 1), New DateTime(2017, 6, 30))
            'Add a filter on the bottom 25 percent of the OrderValue
            rowField2.Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, dataField1, 25, False)
            pt2.DataOnRows = False

            range.AutoFitColumns(0)
        End Function

        Private Shared Async Function LoadFromDatabase(ByVal ws As ExcelWorksheet) As Task(Of ExcelRangeBase)
            Dim range As ExcelRangeBase
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        range = Await ws.Cells("A1").LoadFromDataReaderAsync(sqlReader, True)
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = True
                        range.Offset(0, 5, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd"
                    End Using
                    'Set the numberformat
                End Using
            End Using

            Return range
        End Function
    End Class
End Namespace
