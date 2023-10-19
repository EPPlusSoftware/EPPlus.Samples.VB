Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Filter
Imports System
Imports System.Data.SQLite

Namespace EPPlusSamples.TablesPivotTablesAndSlicers
    Public Class SlicerSample
        Public Shared Sub Run()
            Console.WriteLine("Running sample 7.2-Table and Pivot Table Slicers")
            Using p = New ExcelPackage()
                'Creates a worksheet with one table and several slicers.
                TableSlicerSample(p)

                'Creates the source data for the pivot tables in a separate sheet.
                CreatePivotTableSourceWorksheet(p)

                'Create a pivot table with a slicer connected to one field.
                PivotTableSlicerSample(p)
                'Create three slicers and two pivot tables, one that connects to both tables and two that connect to each of the tables.
                PivotTableOneSlicerToMultiplePivotTables(p)

                p.SaveAs(FileUtil.GetCleanFileInfo("7.3-Slicers.xlsx"))
            End Using
            Console.WriteLine("Sample 7.2 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub

        Private Shared Sub PivotTableOneSlicerToMultiplePivotTables(ByVal p As ExcelPackage)
            Dim wsSource = p.Workbook.Worksheets("PivotTableSourceData")
            Dim wsPivot = p.Workbook.Worksheets.Add("OneSlicerToTwoPivotTables")

            Dim pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells("A15"), wsSource.Cells(wsSource.Dimension.Address), "PivotTable1")
            pivotTable1.RowFields.Add(pivotTable1.Fields("CompanyName"))
            pivotTable1.DataFields.Add(pivotTable1.Fields("OrderValue"))
            pivotTable1.DataFields.Add(pivotTable1.Fields("Tax"))
            pivotTable1.DataFields.Add(pivotTable1.Fields("Freight"))
            pivotTable1.DataOnRows = False

            'To connect a slicer to multiple pivot tables the tables need to use the same pivot table cache, so we use pivotTable1's cache as source to pivotTable2...
            Dim pivotTable2 = wsPivot.PivotTables.Add(wsPivot.Cells("F15"), pivotTable1.CacheDefinition, "PivotTable2")
            pivotTable2.RowFields.Add(pivotTable2.Fields("Country"))
            pivotTable2.DataFields.Add(pivotTable2.Fields("OrderValue"))
            pivotTable2.DataFields.Add(pivotTable2.Fields("Tax"))
            pivotTable2.DataFields.Add(pivotTable2.Fields("Freight"))
            pivotTable2.DataOnRows = False

            Dim slicer1 = pivotTable1.Fields("Country").AddSlicer()
            slicer1.Caption = "Country - Both"

            'Now add the second pivot table to the slicer cache. This require that the pivot tables share the same cache. 
            slicer1.Cache.PivotTables.Add(pivotTable2)
            slicer1.SetPosition(0, 0, 0, 0)
            slicer1.Style = eSlicerStyle.Light4

            Dim slicer2 = pivotTable1.Fields("CompanyName").AddSlicer()
            slicer2.Caption = "Company Name - PivotTable1"
            slicer2.ChangeCellAnchor(eEditAs.Absolute)
            slicer2.SetPosition(0, 192)
            slicer2.SetSize(256, 260)

            Dim slicer3 = pivotTable2.Fields("Orderdate").AddSlicer()
            slicer3.Caption = "Order date - PivotTable2"
            slicer3.ChangeCellAnchor(eEditAs.Absolute)
            slicer3.SetPosition(0, 448)
            slicer3.SetSize(256, 260)
        End Sub
        Private Shared Sub TableSlicerSample(ByVal p As ExcelPackage)
            Dim worksheet1 = p.Workbook.Worksheets.Add("TableSlicer")
            Dim worksheet2 = p.Workbook.Worksheets.Add("TableSlicerToOtherWorksheet")
            ' Lets connect to the sample database for some data
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        Dim range = worksheet1.Cells("A14").LoadFromDataReader(sqlReader, True, "tblSalesData", Table.TableStyles.Medium6)
                        Dim tbl = worksheet1.Tables.GetFromRange(range)
                        range.Offset(1, 4, range.Rows - 1, 1).Style.Numberformat.Format = "yyyy-MM-dd"
                        range.Offset(1, 5, range.Rows - 1, 3).Style.Numberformat.Format = "#,##0"
                        range.AutoFitColumns()

                        'You can either add a slicer via the table column...
                        Dim slicer1 = tbl.Columns(0).AddSlicer()
                        'Filter values are compared to the Text property of the cell. 
                        slicer1.FilterValues.Add("Barton-Veum")
                        slicer1.FilterValues.Add("Christiansen LLC")
                        slicer1.SetPosition(0, 0, 0, 0)

                        '... or you can add it via the drawings collection.
                        Dim slicer2 = worksheet1.Drawings.AddTableSlicer(tbl.Columns("Country"))
                        slicer2.SetPosition(0, 192)

                        'A slicer also supports date groups...
                        Dim slicer3 = tbl.Columns("OrderDate").AddSlicer()
                        slicer3.FilterValues.Add(New ExcelFilterDateGroupItem(2017, 6))
                        slicer3.FilterValues.Add(New ExcelFilterDateGroupItem(2017, 7))
                        slicer3.FilterValues.Add(New ExcelFilterDateGroupItem(2017, 8))
                        slicer3.SetPosition(0, 384)

                        '... You can also add a slicer to another worksheet, if you use the drawings collection...
                        Dim slicer4 = worksheet2.Drawings.AddTableSlicer(tbl.Columns("E-Mail"))
                        slicer4.Caption = "E-Mail - TableSlicer Worksheet"
                        slicer4.Description = "This slicer reference a table in another worksheet."
                        slicer4.SetPosition(1, 0, 1, 0)
                        slicer4.To.Column = 7  'Set the end position anchor to column H, to make the slicer wider.

                        Dim shape = worksheet2.Drawings.AddShape("InfoText", eShapeStyle.Rect)
                        shape.SetPosition(1, 0, 8, 0)
                        shape.Text = "This slicer filters the table located in the TableSlicer worksheet"
                    End Using
                End Using
                sqlConn.Close()
            End Using
            worksheet1.View.FreezePanes(14, 1)
        End Sub
        Private Shared Sub PivotTableSlicerSample(ByVal p As ExcelPackage)
            Dim wsSource = p.Workbook.Worksheets("PivotTableSourceData")
            Dim wsPivot = p.Workbook.Worksheets.Add("PivotTableSlicer")

            Dim pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells("A1"), wsSource.Cells(wsSource.Dimension.Address), "PivotTable1")
            pivotTable.RowFields.Add(pivotTable.Fields("CompanyName"))
            pivotTable.DataFields.Add(pivotTable.Fields("OrderValue"))
            pivotTable.DataFields.Add(pivotTable.Fields("Tax"))
            pivotTable.DataFields.Add(pivotTable.Fields("Freight"))
            pivotTable.DataOnRows = False

            Dim slicer1 = pivotTable.Fields("Name").AddSlicer()
            slicer1.SetPosition(0, 0, 10, 0)
            slicer1.SetSize(400, 208)
            slicer1.Style = eSlicerStyle.Light4
            slicer1.Cache.Data.Items.GetByValue("Brown Kutch").Hidden = True
            slicer1.Cache.Data.Items.GetByValue("Tierra Ratke").Hidden = True
            slicer1.Cache.Data.Items.GetByValue("Jamarcus Schimmel").Hidden = True

            'Add a column with two columns and start showing the item 3.
            slicer1.ColumnCount = 2 'Use two columns on this slicer
            slicer1.StartItem = 3   'First visible item is 3
            slicer1.Cache.Data.CrossFilter = eCrossFilter.ShowItemsWithNoData
            slicer1.Cache.Data.SortOrder = eSortOrder.Descending
        End Sub
        Private Shared Sub CreatePivotTableSourceWorksheet(ByVal p As ExcelPackage)
            Dim wsSource = p.Workbook.Worksheets.Add("PivotTableSourceData")
            'Lets connect to the sample database for some data
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(OrdersWithTaxAndFreightSql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        Dim range = wsSource.Cells("A1").LoadFromDataReader(sqlReader, True)
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = True
                        range.Offset(1, 4, range.Rows - 1, 1).Style.Numberformat.Format = "yyyy-MM-dd hh:mm"
                        range.Offset(1, 5, range.Rows - 1, 3).Style.Numberformat.Format = "#,##0"
                    End Using
                End Using
                sqlConn.Close()
            End Using
            wsSource.Cells.AutoFitColumns()
        End Sub
    End Class
End Namespace
