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
Imports OfficeOpenXml
Imports System.Data.SQLite
Imports System.Threading.Tasks
Imports OfficeOpenXml.Table
Imports EPPlusSamples.FiltersAndValidations

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Friend Class UsingAsyncAwaitSample
        ''' <summary>
        ''' Shows a few different ways to load / save asynchronous
        ''' </summary>
        Public Shared Async Function RunAsync() As Task
            Console.WriteLine("Running sample 1.3-Async-Await")
            Dim file = FileUtil.GetCleanFileInfo("1.03-AsyncAwait.xlsx")
            Using package As ExcelPackage = New ExcelPackage(file)
                Dim ws = package.Workbook.Worksheets.Add("Sheet1")

                Using sqlConn = New SQLiteConnection(ConnectionString)
                    sqlConn.Open()
                    Dim sql = OrdersSql
                    Using sqlCmd = New SQLiteCommand(sql, sqlConn)
                        Dim range = Await ws.Cells("B2").LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), True, "Table1", TableStyles.Medium10)
                        range.AutoFitColumns()
                    End Using
                End Using

                Await package.SaveAsync()
            End Using

            'Load the package async again.
            Using package = New ExcelPackage()
                Await package.LoadAsync(file)

                Dim newWs = package.Workbook.Worksheets.Add("AddedSheet2")
                Dim range = Await newWs.Cells("A1").LoadFromTextAsync(FileUtil.GetFileInfo("01-Workbook Worksheet and Ranges\03-Using Async Await", "Importfile.txt"), New ExcelTextFormat With {
                    .Delimiter = ChrW(9)
                })
                range.AutoFitColumns()

                Await package.SaveAsAsync(FileUtil.GetCleanFileInfo("1.03-AsyncAwait-LoadedAndModified.xlsx"))
            End Using
            Console.WriteLine("Sample 1.3 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()

        End Function
    End Class
End Namespace
