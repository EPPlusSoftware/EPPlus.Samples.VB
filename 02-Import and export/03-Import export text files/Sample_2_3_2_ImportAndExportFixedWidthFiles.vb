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
' 04/29/2024         EPPlus Software AB           Initial release EPPlus 7
' ***********************************************************************************************
Imports OfficeOpenXml
Imports System
Imports System.IO

Namespace EPPlusSamples._02_Import_and_export._03_Import_export_text_files
    ''' <summary>
    ''' This sample shows how to load/save Fixed Width files using the LoadFromText and SaveToText methods.
    ''' </summary>
    Public Module ImportAndExportFixedWidthFiles
        Public Sub RunSample()
            Console.WriteLine("Running sample 2.3.2")
            Dim fixedWidthFile = EPPlusSamples.FileUtil.GetFileInfo("02-Import and Export\03-Import export text files", "Sample2.3.2-1.txt")
            Dim workbookFixedWidth = EPPlusSamples.FileUtil.GetFileInfo("Workbooks", "2.3.2-FixedWidthExport.xlsx").FullName
            Dim newWorkbook As FileInfo = EPPlusSamples.FileUtil.GetCleanFileInfo("2.3.2-LoadDataFromFixedWidthFiles.xlsx")

            'Import fixed width text file using column length.
            If True Then
                Console.WriteLine("Importing the file using column lengths...")
                'Create a workbook and a worksheet.
                Dim package = New ExcelPackage()
                Dim sheet = package.Workbook.Worksheets.Add("FixedWidthLengths")

                'Create the import settings object.
                Dim format As ExcelTextFormatFixedWidth = New ExcelTextFormatFixedWidth()

                'Set the length of each column.
                format.SetColumnLengths(16, 10, 16, 8, 1)
                'Skip the header row.
                format.SkipLinesBeginning = 1

                'Load the fixed width text file into.
                Dim range = sheet.Cells(CStr("A1")).LoadFromText(fixedWidthFile, format)

                'Save the excel file.
                package.SaveAs(newWorkbook)
            End If

            'Import fixed width text file using column starting position.
            If True Then
                Console.WriteLine("Importing the file using column positions...")
                'Create a workbook and a worksheet.
                Dim package = New ExcelPackage(newWorkbook)
                Dim sheet = package.Workbook.Worksheets.Add("FixedWidthPositions")

                'Create the import settings object.
                Dim format As ExcelTextFormatFixedWidth = New ExcelTextFormatFixedWidth()

                'Set the length of a row and the starting positions of each column.
                format.SetColumnPositions(51, 0, 16, 26, 42, 50)
                'Skip the header row.
                format.SkipLinesBeginning = 1

                'Load the fixed width text file into.
                Dim range = sheet.Cells(CStr("A1")).LoadFromText(fixedWidthFile, format)

                'Save the excel file.
                package.Save()
            End If



            'Export fixed width file using column length.
            If True Then
                Console.WriteLine("Exporting the file using column lengths...")
                'Load workbook and worksheet.
                Dim package = New ExcelPackage(workbookFixedWidth)
                Dim sheet = package.Workbook.Worksheets("Sheet1")

                'Create the export settings object.
                Dim format As ExcelOutputTextFormatFixedWidth = New ExcelOutputTextFormatFixedWidth()

                'Set the length of the row and the staring positions of each column
                format.SetColumnLengths(16, 10, 16, 8, 8)
                'Skip the header row.
                format.SkipLinesBeginning = 1
                'Write header
                format.Header = "Name            Date      Amount          Percent Category"

                'Export the range to fixed width text file.
                sheet.Cells(CStr("A1:E6")).SaveToText(EPPlusSamples.FileUtil.GetCleanFileInfo("2.3.2-ExportedLengthsFromEPPlus.txt"), format)
            End If



            'Export fixed width file using column starting position.
            If True Then
                Console.WriteLine("Exporting the file using column positions...")
                'Load workbook and worksheet.
                Dim package = New ExcelPackage(workbookFixedWidth)
                Dim sheet = package.Workbook.Worksheets("Sheet1")

                'Create the export settings object.
                Dim format As ExcelOutputTextFormatFixedWidth = New ExcelOutputTextFormatFixedWidth()

                'Set the length of the row and the staring positions of each column
                format.SetColumnPositions(59, 0, 16, 26, 42, 50)
                'Skip the header row.
                format.SkipLinesBeginning = 1
                'Write header
                format.Header = "Name            Date      Amount          Percent Category"

                'Export the range to fixed width text file.
                sheet.Cells(CStr("A1:E6")).SaveToText(EPPlusSamples.FileUtil.GetCleanFileInfo("2.3.2-ExportedPositionsFromEPPlus.txt"), format)
            End If
        End Sub
    End Module
End Namespace
