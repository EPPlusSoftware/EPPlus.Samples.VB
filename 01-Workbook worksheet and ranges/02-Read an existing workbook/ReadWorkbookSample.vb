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
Imports System.IO
Imports OfficeOpenXml

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    ''' <summary>
    ''' Simply opens an existing file and reads some values and properties
    ''' </summary>
    Public Module ReadWorkbookSample
        Public Sub Run()
            Console.WriteLine("Running sample 1.2")
            Dim filePath = FileUtil.GetFileInfo("Workbooks", "1.02-ReadWorkbook.xlsx").FullName
            Console.WriteLine("Reading column 2 of {0}", filePath)
            Console.WriteLine()

            Dim existingFile As FileInfo = New FileInfo(filePath)
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                'Get the first worksheet in the workbook
                Dim worksheet = package.Workbook.Worksheets(0)

                Dim col = 2 'Column 2 is the item description
                For row = 2 To 4
                    Console.WriteLine(vbTab & "Cell({0},{1}).Value={2}", row, col, worksheet.Cells(row, col).Value)
                Next

                'Output the formula from row 3 in A1 and R1C1 format
                Console.WriteLine(vbTab & "Cell({0},{1}).Formula={2}", 3, 5, worksheet.Cells(3, 5).Formula)
                Console.WriteLine(vbTab & "Cell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells(3, 5).FormulaR1C1)

                'Output the formula from row 5 in A1 and R1C1 format
                Console.WriteLine(vbTab & "Cell({0},{1}).Formula={2}", 5, 3, worksheet.Cells(5, 3).Formula)
                Console.WriteLine(vbTab & "Cell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells(5, 3).FormulaR1C1)

            End Using ' the using statement automatically calls Dispose() which closes the package.

            Console.WriteLine()
            Console.WriteLine("Read workbook sample complete")
            Console.WriteLine()
        End Sub
    End Module
End Namespace
