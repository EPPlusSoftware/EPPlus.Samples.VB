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
Imports System.Linq

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module ReadDataUsingLinq
        ''' <summary>
        ''' This sample shows how to use Linq with the Cells collection
        ''' </summary>
        ''' <paramname="outputDir">The path where sample7.xlsx is</param>
        Public Sub Run()
            Console.WriteLine("Running sample 10-Linq")
            Console.WriteLine("Now open sample 1.7 again and perform some Linq queries...")
            Console.WriteLine()

            Dim existingFile = FileUtil.GetFileInfo("1.07-PerformanceAndProtection.xlsx")
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                Dim sheet = package.Workbook.Worksheets(0)

                'Select all cells in column d between 9990 and 10000
                Dim query1 = From cell In sheet.Cells("d:d") Where TypeOf cell.Value Is Double AndAlso CDbl(cell.Value) >= 9990 AndAlso CDbl(cell.Value) <= 10000 Select cell

                Console.WriteLine("Print all cells with value between 9990 and 10000 in column D ...")
                Console.WriteLine()

                Dim count = 0
                For Each cell In query1
                    Console.WriteLine("Cell {0} has value {1:N0}", cell.Address, cell.Value)
                    count += 1
                Next

                Console.WriteLine("{0} cells found ...", count)
                Console.WriteLine()

                'Select all bold cells
                Console.WriteLine("Now get all bold cells from the entire sheet...")
                Dim query2 = From cell In sheet.Cells(sheet.Dimension.Address) Where cell.Style.Font.Bold Select cell
                'If you have a clue where the data is, specify a smaller range in the cells indexer to get better performance (for example "1:1,65536:65536" here)
                count = 0
                For Each cell In query2
                    If Not String.IsNullOrEmpty(cell.Formula) Then
                        Console.WriteLine("Cell {0} is bold and has a formula of {1:N0}", cell.Address, cell.Formula)
                    Else
                        Console.WriteLine("Cell {0} is bold and has a value of {1:N0}", cell.Address, cell.Value)
                    End If
                    count += 1
                Next

                'Here we use more than one column in the where clause. We start by searching column D, then use the Offset method to check the value of column C.
                Dim query3 = (From cell In sheet.Cells("d:d") Where TypeOf cell.Value Is Double AndAlso CDbl(cell.Value) >= 9500 AndAlso CDbl(cell.Value) <= 10000 AndAlso cell.Offset(0, -1).GetValue(Of Date)().Year = Date.Today.Year + 1 Select cell)



                Console.WriteLine()
                Console.WriteLine("Print all cells with a value between 9500 and 10000 in column D and the year of Column C is {0} ...", Date.Today.Year + 1)
                Console.WriteLine()

                count = 0
                For Each cell In query3    'The cells returned here will all be in column D, since that is the address in the indexer. Use the Offset method to print any other cells from the same row.
                    Console.WriteLine("Cell {0} has value {1:N0} Date is {2:d}", cell.Address, cell.Value, cell.Offset(0, -1).GetValue(Of Date)())
                    count += 1
                Next
                Console.WriteLine()
            End Using
        End Sub
    End Module
End Namespace
