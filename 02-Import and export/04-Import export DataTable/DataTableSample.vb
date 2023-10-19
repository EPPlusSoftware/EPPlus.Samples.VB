Imports OfficeOpenXml
Imports OfficeOpenXml.Export.ToDataTable
Imports OfficeOpenXml.Table
Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.Text

Namespace EPPlusSamples
    Public Class DataTableSample
        Public Shared Sub Run()
            Console.WriteLine("Running sample 2.4 - Import and Export DataTable")
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Dim sql = GetSql()
                Using sqlCmd = New SQLiteCommand(sql, sqlConn)
                    Dim reader = sqlCmd.ExecuteReader()
                    Dim dataTable = New DataTable()
                    dataTable.Load(reader)

                    ' Create a workbook 
                    Using package = New ExcelPackage()
                        Dim sheet = package.Workbook.Worksheets.Add("DataTable Samples")

                        ' *** Load from DataTable **

                        ' Import the DataTable using LoadFromDataTable
                        sheet.Cells("A1").LoadFromDataTable(dataTable, True, TableStyles.Dark11)

                        ' Now let's export this data back to a DataTable. We know that the data is in a 
                        ' table, so we are using the ExcelTables interface to get the range
                        Dim dt1 = sheet.Tables(0).ToDataTable()
                        PrintDataTable(dt1)


                        ' *** Export to DataTable **

                        ' Export a specific range instead of the entire table
                        ' and use the config action to set the table name
                        Dim dt2 = sheet.Cells("A1:F11").ToDataTable(Sub(o) o.DataTableName = "dt2")
                        PrintDataTable(dt2)

                        ' Configure some properties on how the table is generated
                        Dim dt3 = sheet.Cells("A1:F11").ToDataTable(Sub(c)
                                                                        ' set name and namespace
                                                                        c.DataTableName = "MyDataTable"
                                                                        c.DataTableNamespace = "MyNamespace"
                                                                        ' Removes spaces in column names when read from the first row
                                                                        c.ColumnNameParsingStrategy = NameParsingStrategy.RemoveSpace
                                                                        ' Rename the third column from E-Mail to EmailAddress
                                                                        c.Mappings.Add(2, "EmailAddress")
                                                                        ' Ensure that the OrderDate column is casted to DateTime (in Excel it can sometimes be stored as a double/OADate)
                                                                        c.Mappings.Add(4, "OrderDate", GetType(Date))
                                                                        ' Change the OrderValue to a string
                                                                        c.Mappings.Add(5, "OrderValue", GetType(String), False, Function(cellVal) "Val: " & cellVal.ToString())
                                                                        ' Skip the first 2 rows
                                                                        c.SkipNumberOfRowsStart = 2
                                                                        ' Skip the last 100 rows
                                                                        c.SkipNumberOfRowsEnd = 4

                                                                    End Sub)
                        PrintDataTable(dt3)

                        ' Export to existing DataTable

                        ' Create the DataTable
                        Dim dataTable2 = New DataTable("myDataTable", "myNamespace")
                        dataTable2.Columns.Add("Company Name", GetType(String))
                        dataTable2.Columns.Add("E-Mail")
                        sheet.Cells("A1:F11").ToDataTable(Sub(o) o.FirstRowIsColumnNames = True, dataTable2)
                        PrintDataTable(dataTable2)

                        ' Create the DataTable, use mappings if names of columns/range headers differ
                        Dim dataTable3 = New DataTable("myDataTableWithMappings", "myNamespace")
                        Dim col1 = dataTable3.Columns.Add("CompanyName")
                        Dim col2 = dataTable3.Columns.Add("Email")
                        sheet.Cells("A1:F11").ToDataTable(Sub(o)
                                                              o.FirstRowIsColumnNames = True
                                                              o.Mappings.Add(0, col1)
                                                              o.Mappings.Add(1, col2)
                                                          End Sub, dataTable3)
                        PrintDataTable(dataTable3)

                    End Using
                End Using
            End Using
            Console.WriteLine("Sample 2.4 finished.")
            Console.WriteLine()
        End Sub

        Private Shared Sub PrintDataTable(ByVal table As DataTable)
            Console.WriteLine()
            Console.WriteLine("DATATABLE name=" & table.TableName)
            Dim cols = New StringBuilder()
            For Each col In table.Columns
                cols.AppendFormat("'{0}' ", CType(col, DataColumn).ColumnName)
            Next
            Console.WriteLine("Columns:")
            Console.WriteLine(cols.ToString())
            Console.WriteLine()

            Console.WriteLine("First 10 rows:")
            Dim r = 0

            While r < table.Rows.Count AndAlso r < 10
                For c = 0 To table.Columns.Count - 1
                    Dim col = TryCast(table.Columns(c), DataColumn)
                    Dim row = TryCast(table.Rows(r), DataRow)
                    Dim val = If(col.DataType Is GetType(String), "'" & row(col.ColumnName).ToString() & "'", row(col.ColumnName))


                    Console.Write(If(c = 0, val, ", " & val.ToString()))
                Next
                Console.WriteLine()
                r += 1
            End While
        End Sub
        Private Shared Function GetSql() As String
            Dim sb = New StringBuilder()
            sb.Append("select cu.Name as 'Company Name', sp.Name, Email as 'E-Mail', co.Name as Country, orderdate as 'Order Date', (ordervalue) as 'Order Value',tax as Tax, freight As Freight, od.currency As Currency ")
            sb.Append("from Customer cu inner join ")
            sb.Append("Orders od on cu.CustomerId=od.CustomerId inner join ")
            sb.Append("SalesPerson sp on od.salesPersonId = sp.salesPersonId inner join ")
            sb.Append("City ci on ci.cityId = sp.cityId inner join ")
            sb.Append("Country co on ci.countryId = co.countryId ")
            sb.Append("ORDER BY 1,2 desc")
            Return sb.ToString()

        End Function
    End Class
End Namespace
