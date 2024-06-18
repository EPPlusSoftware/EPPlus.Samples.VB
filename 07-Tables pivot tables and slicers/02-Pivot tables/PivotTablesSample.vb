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
Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Table.PivotTable
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Data.SQLite
Imports EPPlusSamples.FiltersAndValidations

Namespace EPPlusSamples.PivotTables
    ''' <summary>
    ''' This class shows how to use pivot tables 
    ''' </summary>
    Public Module PivotTablesSample
        Public Class SalesDTO
            Public Property CompanyName As String
            Public Property Name As String
            Public Property Email As String
            Public Property Country As String
            Public Property OrderId As Integer
            Public Property OrderDate As Date
            Public Property OrderValue As Decimal
            Public Property Tax As Decimal
            Public Property Freight As Decimal
            Public Property Currency As String
        End Class
        Public Function Run() As String
            Console.WriteLine("Running sample 7.3-Pivot Table")

            Dim list = GetDataFromSQL()

            Dim newFile As FileInfo = EPPlusSamples.FileUtil.GetCleanFileInfo("7.2-PivotTables.xlsx")
            Using pck As ExcelPackage = New ExcelPackage(newFile)
                ' get the handle to the existing worksheet
                Dim wsData = pck.Workbook.Worksheets.Add("SalesData")

                Dim dataRange = wsData.Cells("A1").LoadFromCollection((From s In list Order By s.Name Select s), True, Table.TableStyles.Medium2)

                wsData.Cells(2, 6, dataRange.End.Row, 6).Style.Numberformat.Format = "mm-dd-yy"
                wsData.Cells(2, 7, dataRange.End.Row, 11).Style.Numberformat.Format = "#,##0"

                dataRange.AutoFitColumns()

                Dim pt1 = CreatePivotTableWithPivotChart(pck, dataRange)
                Dim pt2 = CreatePivotTableWithDataGrouping(pck, dataRange)
                Dim pt3 = CreatePivotTableWithPageFilter(pck, pt2.CacheDefinition)
                Dim pt4 = CreatePivotTableWithASlicer(pck, pt2.CacheDefinition)
                Dim pt5 = CreatePivotTableWithACalculatedField(pck, pt2.CacheDefinition)
                Dim pt6 = CreatePivotTableCaptionFilter(pck, dataRange)
                Dim pt7 = CreatePivotTableWithDataFieldsUsingShowAs(pck, dataRange)


                pt1.Calculate()
                CreatePivotTableSorting(pck, dataRange)

                pck.Save()
            End Using
            Return newFile.FullName
        End Function


        Private Function CreatePivotTableWithPivotChart(pck As ExcelPackage, dataRange As ExcelRangeBase) As ExcelPivotTable
            Dim wsPivot = pck.Workbook.Worksheets.Add("PivotSimple")
            Dim pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells("A1"), dataRange, "PerCountry")

            pivotTable.RowFields.Add(pivotTable.Fields("Country"))
            Dim dataField = pivotTable.DataFields.Add(pivotTable.Fields("OrderValue"))
            dataField.Format = "#,##0"
            pivotTable.DataOnRows = True

            Dim chart = wsPivot.Drawings.AddPieChart("PivotChart", ePieChartType.PieExploded3D, pivotTable)
            chart.SetPosition(1, 0, 4, 0)
            chart.SetSize(800, 600)
            chart.Legend.Remove()
            chart.Series(0).DataLabel.ShowCategory = True
            chart.Series(0).DataLabel.Position = eLabelPosition.OutEnd
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle6)
            Return pivotTable
        End Function

        Private Function CreatePivotTableWithDataGrouping(pck As ExcelPackage, dataRange As ExcelRangeBase) As ExcelPivotTable
            Dim wsPivot2 = pck.Workbook.Worksheets.Add("PivotDateGrp")
            Dim pivotTable2 = wsPivot2.PivotTables.Add(wsPivot2.Cells("A3"), dataRange, "PerEmploeeAndQuarter")

            pivotTable2.RowFields.Add(pivotTable2.Fields("Name"))

            'Add a rowfield
            Dim rowField = pivotTable2.RowFields.Add(pivotTable2.Fields("OrderDate"))
            'This is a date field so we want to group by Years and quaters. This will create one additional field for years.
            rowField.AddDateGrouping(eDateGroupBy.Years Or eDateGroupBy.Quarters)
            rowField.Name = "Quarters" 'We rename the field OrderDate to Quarters.

            'Get the Quaters field and change the texts
            Dim quaterField = pivotTable2.Fields.GetDateGroupField(eDateGroupBy.Quarters)
            quaterField.Items(0).Text = "<" 'Values below min date, but we use auto so its not used
            quaterField.Items(1).Text = "Q1"
            quaterField.Items(2).Text = "Q2"
            quaterField.Items(3).Text = "Q3"
            quaterField.Items(4).Text = "Q4"
            quaterField.Items(5).Text = ">" 'Values above max date, but we use auto so its not used

            'Add a pagefield
            Dim pageField = pivotTable2.PageFields.Add(pivotTable2.Fields("CompanyName"))

            'Add the data fields and format them
            Dim dataField As ExcelPivotTableDataField
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields("OrderValue"))
            dataField.Format = "#,##0"
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields("Tax"))
            dataField.Format = "#,##0"
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields("Freight"))
            dataField.Format = "#,##0"

            'We want the datafields to appear in columns
            pivotTable2.DataOnRows = False
            Return pivotTable2
        End Function
        Private Function CreatePivotTableWithPageFilter(pck As ExcelPackage, pivotCache As ExcelPivotCacheDefinition) As ExcelPivotTable
            Dim wsPivot3 = pck.Workbook.Worksheets.Add("PivotWithPageField")

            'Create a new pivot table using the same cache as pivot table 2.
            Dim pivotTable3 = wsPivot3.PivotTables.Add(wsPivot3.Cells("A3"), pivotCache, "PerEmploeeSelectedCompanies")

            pivotTable3.RowFields.Add(pivotTable3.Fields("Name"))

            'Add a rowfield
            Dim rowField = pivotTable3.RowFields.Add(pivotTable3.Fields("OrderDate"))

            'Add a pagefield
            Dim pageField = pivotTable3.PageFields.Add(pivotTable3.Fields("CompanyName"))
            pageField.Items.Refresh()  'Refresh the items from the source range.

            pageField.Items(1).Hidden = True   'Hide item with index 1 in the items collection
            pageField.Items.GetByValue("Walsh LLC").Hidden = True  'Hide the item with supplied the value . 
            'pageField.Items.SelectSingleItem(3); //You can also select a single item with this method

            'Add the data fields and format them
            Dim dataField As ExcelPivotTableDataField
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields("OrderValue"))
            dataField.Format = "#,##0"
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields("Tax"))
            dataField.Format = "#,##0"
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields("Freight"))
            dataField.Format = "#,##0"


            'We want the datafields to appear in columns
            pivotTable3.DataOnRows = False
            Return pivotTable3
        End Function
        Private Function CreatePivotTableWithASlicer(pck As ExcelPackage, pivotCache As ExcelPivotCacheDefinition) As ExcelPivotTable
            'This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            Dim wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithSlicer")

            'Create a new pivot table using the same cache as pivot table 2.
            Dim pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells("A3"), pivotCache, "PerEmploeeSelectedCompSlicer")

            pivotTable4.RowFields.Add(pivotTable4.Fields("Name"))

            'Add a rowfield
            pivotTable4.RowFields.Add(pivotTable4.Fields("OrderDate"))

            'Add slicer
            Dim companyNameField = pivotTable4.Fields("CompanyName")
            Dim slicer = companyNameField.AddSlicer()
            slicer.SetPosition(3, 0, 5, 0) 'Set top left to row 4, column F

            companyNameField.Items.Refresh()  'Refresh the items from the source range.

            companyNameField.Items(1).Hidden = True   'Hide item with index 1 in the items collection
            companyNameField.Items.GetByValue("Walsh LLC").Hidden = True  'Hide the item with supplied the value . 

            'Add the data fields and format them
            Dim dataField As ExcelPivotTableDataField
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("OrderValue"))
            dataField.Format = "#,##0"
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("Tax"))
            dataField.Format = "#,##0"
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("Freight"))
            dataField.Format = "#,##0"

            'We want the data fields to appear in columns
            pivotTable4.DataOnRows = False
            Return pivotTable4
        End Function
        Private Function CreatePivotTableWithACalculatedField(pck As ExcelPackage, pivotCache As ExcelPivotCacheDefinition) As ExcelPivotTable
            'This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            Dim wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithCalculatedField")

            'Create a new pivot table using the same cache as pivot table 2.
            Dim pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells("A3"), pivotCache, "PerWithCalculatedField")

            pivotTable4.RowFields.Add(pivotTable4.Fields("CompanyName"))
            'Be careful with formulas as they are not validated and can cause the pivot table to become corrupt. 

            'Be careful with formulas as they can cause the pivot table to become corrupt if they are entered invalidly.
            Dim calcField = pivotTable4.Fields.AddCalculatedField("Total", "'OrderValue'+'Tax'+'Freight'")
            calcField.Format = "#,##0"

            'Add the data fields and format them
            Dim dataField As ExcelPivotTableDataField
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("OrderValue"))
            dataField.Format = "#,##0"
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("Tax"))
            dataField.Format = "#,##0"
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("Freight"))
            dataField.Format = "#,##0"
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields("Total"))
            dataField.Format = "#,##0"


            'We want the data fields to appear in columns
            pivotTable4.DataOnRows = False
            Return pivotTable4
        End Function
        Private Function CreatePivotTableCaptionFilter(pck As ExcelPackage, dataRange As ExcelRangeBase) As ExcelPivotTable
            Dim wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithCaptionFilter")

            'Create a new pivot table with a new cache.
            Dim pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells("A3"), dataRange, "WithCaptionFilter")

            Dim rowField1 = pivotTable4.RowFields.Add(pivotTable4.Fields("Name"))

            'Add the Caption filter (Label filter in Excel) to the pivot table.
            rowField1.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBeginsWith, "C")

            'Add a rowfield
            Dim rowField2 = pivotTable4.RowFields.Add(pivotTable4.Fields("OrderDate"))

            'Add a date value filter to the pivot table.
            rowField2.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, New DateTime(2017, 8, 1), New DateTime(2017, 8, 31))

            'Filters will apply on top of any selection made directly on the items.
            rowField2.Items.Refresh()
            rowField2.Items(8).Hidden = True

            'Number formats can be set directly on fields as well as on datafields...
            pivotTable4.Fields("OrderDate").Format = "yyyy-MM-dd hh:mm:ss"
            pivotTable4.Fields("OrderValue").Format = "#,##0"
            pivotTable4.Fields("Tax").Format = "#,##0"
            pivotTable4.Fields("Freight").Format = "#,##0"

            'Add the data fields and format them
            pivotTable4.DataFields.Add(pivotTable4.Fields("OrderValue"))
            pivotTable4.DataFields.Add(pivotTable4.Fields("Tax"))
            pivotTable4.DataFields.Add(pivotTable4.Fields("Freight"))

            'We want the datafields to appear in columns
            pivotTable4.DataOnRows = False
            Return pivotTable4
        End Function
        Private Function CreatePivotTableWithDataFieldsUsingShowAs(pck As ExcelPackage, dataRange As ExcelRangeBase) As ExcelPivotTable
            Dim wsPivot5 = pck.Workbook.Worksheets.Add("PivotWithShowAsFields")

            'Create a new pivot table with a new cache.
            Dim pivotTable5 = wsPivot5.PivotTables.Add(wsPivot5.Cells("A3"), dataRange, "WithCaptionFilter")

            Dim rowField1 = pivotTable5.RowFields.Add(pivotTable5.Fields("CompanyName"))
            Dim rowField2 = pivotTable5.RowFields.Add(pivotTable5.Fields("Name"))
            Dim colField1 = pivotTable5.ColumnFields.Add(pivotTable5.Fields("Currency"))

            'Collapses all row and column fields
            rowField1.Items.Refresh()
            rowField1.Items.ShowDetails(False)

            rowField2.Items.Refresh()
            rowField2.Items.ShowDetails(False)

            colField1.Items.Refresh()
            colField1.Items.ShowDetails(False)

            'Sets the ∑ Values position within column or row fields collection.
            'The value of the pivotTable5.DataOnRows will determin if the rowFields or columnsFields collection is used.
            'A negative or out of range value will add the values to the end of the collection.
            pivotTable5.DataOnRows = False
            pivotTable5.ValuesFieldPosition = 0    'Set values first in the row fields collection

            Dim df1 = pivotTable5.DataFields.Add(pivotTable5.Fields("OrderValue"))
            df1.Name = "Order value"
            df1.Format = "#,##0"

            Dim df2 = pivotTable5.DataFields.Add(pivotTable5.Fields("OrderValue"))
            df2.Name = "Order value % of total"
            df2.ShowDataAs.SetPercentOfColumn()
            df2.Format = "0.0%;"

            Dim df3 = pivotTable5.DataFields.Add(pivotTable5.Fields("OrderValue"))
            df3.Name = "Count Difference From Previous"
            df3.ShowDataAs.SetDifference(rowField1, ePrevNextPivotItem.Previous)
            df3.Function = DataFieldFunctions.Count
            df3.Format = "#,##0"

            pivotTable5.SetCompact(False)
            pivotTable5.ColumnHeaderCaption = "Data"
            pivotTable5.ShowColumnStripes = True
            wsPivot5.Columns(1).Width = 30

            Return pivotTable5
        End Function
        Private Sub CreatePivotTableSorting(pck As ExcelPackage, dataRange As ExcelRangeBase)
            Dim wsPivot = pck.Workbook.Worksheets.Add("PivotSorting")

            'Sort by the row field
            Dim pt1 = wsPivot.PivotTables.Add(wsPivot.Cells("A1"), dataRange, "PerCountrySorted")
            pt1.DataOnRows = True

            Dim rowField1 = pt1.RowFields.Add(pt1.Fields("Country"))
            rowField1.Sort = eSortType.Ascending
            Dim dataField = pt1.DataFields.Add(pt1.Fields("OrderValue"))
            dataField.Format = "#,##0"


            'Sort by the datafield field
            Dim pt2 = wsPivot.PivotTables.Add(wsPivot.Cells("D1"), dataRange, "PerCountrySortedByData")
            pt2.DataOnRows = True

            rowField1 = pt2.RowFields.Add(pt2.Fields("Country"))
            dataField = pt2.DataFields.Add(pt2.Fields("OrderValue"))
            dataField.Format = "#,##0"
            rowField1.SetAutoSort(dataField, eSortType.Descending)


            'Sort by the data field for a specific column using pivot areas.
            'In this case we sort on the order value column for "Poland". 
            Dim pt3 = wsPivot.PivotTables.Add(wsPivot.Cells("G1"), dataRange, "PerCountrySortedByDataColumn")
            pt3.DataOnRows = True

            rowField1 = pt3.RowFields.Add(pt3.Fields("Name"))
            Dim columnField1 = pt3.ColumnFields.Add(pt3.Fields("Country"))
            dataField = pt3.DataFields.Add(pt3.Fields("OrderValue"))
            dataField.Format = "#,##0"
            rowField1.SetAutoSort(dataField, eSortType.Ascending)

            Dim conditionField = rowField1.AutoSort.Conditions.Fields.Add(columnField1)
            'Before setting a reference to a value column we need to refresh the items cache.
            columnField1.Items.Refresh()
            conditionField.Items.AddByValue("Poland")
        End Sub

        Private Function GetDataFromSQL() As List(Of SalesDTO)
            Dim ret = New List(Of SalesDTO)()
            Using sqlConn = New SQLiteConnection(EPPlusSamples.SampleSettings.ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(EPPlusSamples.FiltersAndValidations.SqlStatements.OrdersWithTaxAndFreightSql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        'Get the data and fill rows 5 onwards
                        While sqlReader.Read()
                            ret.Add(New SalesDTO With {
    .CompanyName = sqlReader("companyName").ToString(),
    .Name = sqlReader("name").ToString(),
    .Email = sqlReader("e-mail").ToString(),
    .Country = sqlReader("country").ToString(),
    .OrderId = Convert.ToInt32(sqlReader("orderId")),
    .OrderDate = sqlReader("OrderDate"),
    .OrderValue = Convert.ToDecimal(sqlReader("OrderValue")),
    .Tax = Convert.ToDecimal(sqlReader("tax")),
    .Freight = Convert.ToDecimal(sqlReader("freight")),
    .Currency = sqlReader("currency").ToString()
})
                        End While
                    End Using
                End Using
            End Using
            Return ret
        End Function
    End Module
End Namespace
