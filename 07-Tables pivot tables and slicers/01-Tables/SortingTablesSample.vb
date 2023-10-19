Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.Data.SQLite
Imports System.Threading.Tasks

Namespace EPPlusSamples.TablesPivotTablesAndSlicers
    ''' <summary>
    ''' This sample demonstrates how to sort Excel tables in EPPlus.
    ''' </summary>
    Public Module SortingTablesSample
        Public Async Function RunAsync() As Task
            Dim file = FileUtil.GetCleanFileInfo("7.1-SortingTables.xlsx")
            Using package As ExcelPackage = New ExcelPackage(file)
                ' Sheet 1
                Dim sheet1 = package.Workbook.Worksheets.Add("Sheet1")
                sheet1.Cells("B1").Value = "This table is sorted by country DESC, then name ASC, then orderValue ASC"
                Using sqlConn = New SQLiteConnection(ConnectionString)
                    sqlConn.Open()
                    Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                        Dim range = Await sheet1.Cells("B3").LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), True, "Table1", TableStyles.Medium10)
                        range.AutoFitColumns()
                    End Using
                End Using
                ' sort this table by country DESC, then by sales persons name ASC, then by Order value ASC
                Dim table1 = sheet1.Tables(0)
                table1.Sort(Sub(x) x.SortBy.ColumnNamed("Country", eSortOrder.Descending).ThenSortBy.ColumnNamed("Name").ThenSortBy.ColumnNamed("OrderValue"))


                ' Sheet 2
                Dim sheet2 = package.Workbook.Worksheets.Add("Using custom list")
                sheet2.Cells("B1").Value = "This table is sorted by country with a custom list, then name ASC, then orderValue ASC. The custom lists ensures that Greenland and Costa Rica comes first in the sort"
                Using sqlConn = New SQLiteConnection(ConnectionString)
                    sqlConn.Open()
                    Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                        Dim range = Await sheet2.Cells("B3").LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), True, "Table2", TableStyles.Medium10)
                        range.AutoFitColumns()
                    End Using
                End Using
                ' sort this table by country ASC, then by sales persons name ASC, then by Order value ASC
                Dim table2 = sheet2.Tables("Table2")
                table2.Sort(Sub(x) x.SortBy.ColumnNamed("country", eSortOrder.Descending).UsingCustomList("Greenland", "Costa Rica").ThenSortBy.ColumnNamed("name").ThenSortBy.ColumnNamed("orderValue"))

                Await package.SaveAsync()
            End Using
        End Function
    End Module
End Namespace
