Imports OfficeOpenXml
Imports System
Imports System.IO
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Module JsonExportSample
        'This sample demonstrates how to export html from a table.
        'More advanced samples using charts and json exports are available in our samples web site available 
        'here: https://samples.epplussoftware.com/JsonExport
        Public Async Function RunAsync() As Task
            Console.WriteLine("Running sample 2.6 - json export")
            Dim outputFolder = FileUtil.GetDirectoryInfo("JsonOutput")

            'Start by using the excel file generated in sample 28
            Using p = New ExcelPackage(FileUtil.GetFileInfo("Workbooks", "Tables.xlsx"))
                Dim wsSimpleTable = p.Workbook.Worksheets("SimpleTable")

                ExportTable1(outputFolder, wsSimpleTable)

                Dim wsStyleTables = p.Workbook.Worksheets("StyleTables")
                Await ExportTableWithHyperlink(outputFolder, wsStyleTables)
            End Using
            Console.WriteLine("Sample 2.6 finished.")
            Console.WriteLine()
        End Function

        Private Sub ExportTable1(ByVal outputFolder As DirectoryInfo, ByVal wsSimpleTable As ExcelWorksheet)
            Dim table1 = wsSimpleTable.Tables(0)

            'First export the table directly from the table object.
            'When exporting a table the data type is set on the column.
            Dim json = table1.ToJson(Sub(x) x.Minify = False)

            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "TableSample1_As_Table.json", True).FullName, json)

            'When exporting the range data types are set on the cell level.
            'You can alter this by AddDataTypesOn, --> x.AddDataTypesOn=eDataTypeOn.OnColumn
            json = table1.Range.ToJson(Sub(x) x.Minify = False)

            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "TableSample1_As_Range.json", True).FullName, json)
        End Sub
        Private Async Function ExportTableWithHyperlink(ByVal outputFolder As DirectoryInfo, ByVal wsStyleTables As ExcelWorksheet) As Task
            Dim table1 = wsStyleTables.Tables(0)


            Using fs = New FileStream(FileUtil.GetFileInfo(outputFolder, "TableSample2_hyperlinks.json", True).FullName, FileMode.Create, FileAccess.Write)
                Await table1.SaveToJsonAsync(fs, Sub(x)
                                                     x.AddDataTypesOn = eDataTypeOn.NoDataTypes 'Skip data types.
                                                     x.Minify = False
                                                 End Sub)
                fs.Close()
            End Using
        End Function
    End Module
End Namespace
