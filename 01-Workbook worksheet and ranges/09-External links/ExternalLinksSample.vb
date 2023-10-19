﻿Imports OfficeOpenXml
Imports System
Imports System.IO

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    ''' <summary>
    ''' This sample demonstrates how work with External references in EPPlus.
    ''' EPPlus supports adding, updating and removing external workbooks of type xlsx, xlsm and xlst. EPPlus also use the external reference cache for External workbooks. 
    ''' EPPlus will also preserve DDE and OLE links.
    ''' </summary>
    Public Module ExternalLinksSample
        Public Sub Run()
            Console.WriteLine("Running sample 1.9 - External Links")
            'Reads a workbook and calculates the formulas from the cache and from the loaded external workbook. 
            Call ReadFileWithExternalLink()

            'Sample file 1 adds external links to another workbook.
            Dim sampleFile1 = FileUtil.GetCleanFileInfo("1.09-ExternalLinks-1.xlsx")
            Using p = New ExcelPackage(sampleFile1)

                AddWorksheetWithExternalReferences(p)

                AddWorksheetWithExternalReferencesInFormulas(p)

                p.Save()
            End Using

            'Open the saved sample and break links to the fist external workbook 
            BreakLinks(sampleFile1)

            Console.WriteLine("Sample 1.9 finished.")
            Console.WriteLine()
        End Sub


        ''' <summary>
        ''' This sample shows how EPPlus works with external workbooks depending on 
        ''' </summary>
        Private Sub ReadFileWithExternalLink()
            Dim externalFile = FileUtil.GetFileInfo("Workbooks", "1.09-WorkbookWithExternalLinks.xlsx")
            Using p = New ExcelPackage(externalFile)
                'This worksheet contains references to an external workbook. 
                'First print the values saved in the package.
                Dim ws = p.Workbook.Worksheets(0)
                Console.WriteLine("Values from Excel:")
                Console.WriteLine($"Cell C1 formula : {ws.Cells("C1").Formula} with value {ws.Cells("C1").Value}")
                Console.WriteLine($"Cell C2 formula : {ws.Cells("C2").Formula} with value {ws.Cells("C2").Value}")
                Console.WriteLine($"Cell C3 formula : {ws.Cells("C3").Formula} with value {ws.Cells("C3").Value}")
                Console.WriteLine($"Cell C5 formula : {ws.Cells("C5").Formula} with value {ws.Cells("C5").Text}")
                Console.WriteLine($"Cell C6 formula : {ws.Cells("C6").Formula} with value {ws.Cells("C6").Value}")
                Console.WriteLine()

                'Now, clear the formula values and calculate the workbook again.
                'In this case, EPPlus uses the package internal saved cache for the external workbook to calculate the formulas referencing this workbook.
                ws.ClearFormulaValues()
                ws.Calculate()

                Console.WriteLine("Values after calculation in EPPLus from external reference cache:")
                Console.WriteLine($"Cell C1 formula : {ws.Cells("C1").Formula} with value {ws.Cells("C1").Value}")
                Console.WriteLine($"Cell C2 formula : {ws.Cells("C2").Formula} with value {ws.Cells("C2").Value}")
                Console.WriteLine($"Cell C3 formula : {ws.Cells("C3").Formula} with value {ws.Cells("C3").Value}")
                Console.WriteLine($"Cell C5 formula : {ws.Cells("C5").Formula} with value {ws.Cells("C5").Text}")
                Console.WriteLine($"Cell C6 formula : {ws.Cells("C6").Formula} with value {ws.Cells("C6").Value}")
                Console.WriteLine()

                'Note in the output, that Cell C6 has a different value from the Excel Calculation.
                'This is because the saved cache does not contain all information required, in this case some of the lines are hidden and should be ignored by the subtotal function. 
                'This is the same behavior as in Excel if you recalculate without the external workbook available.

                'To avoid this behavior you can load the external workbook before doing the calculate.
                'This is only an issue in special cases where the function needs information not available in the cache, as for example hidden cells and numeric formats.
                Dim externalReference = p.Workbook.ExternalLinks(0).As.ExternalWorkbook
                p.Workbook.ExternalLinks.Directories.Add(FileUtil.GetSubDirectory("Workbooks", "1.9-Data"))
                externalReference.Load()

                ws.ClearFormulaValues()
                ws.Calculate()

                Console.WriteLine("Values after calculation in EPPLus when the external package has been loaded:")
                Console.WriteLine($"Cell C1 formula : {ws.Cells("C1").Formula} with value {ws.Cells("C1").Value}")
                Console.WriteLine($"Cell C2 formula : {ws.Cells("C2").Formula} with value {ws.Cells("C2").Value}")
                Console.WriteLine($"Cell C3 formula : {ws.Cells("C3").Formula} with value {ws.Cells("C3").Value}")
                Console.WriteLine($"Cell C5 formula : {ws.Cells("C5").Formula} with value {ws.Cells("C5").Text}")
                Console.WriteLine($"Cell C6 formula : {ws.Cells("C6").Formula} with value {ws.Cells("C6").Value}")
                Console.WriteLine()
            End Using
        End Sub

        Private Sub AddWorksheetWithExternalReferences(ByVal p As ExcelPackage)
            'Add a reference to the file created by sample 28.
            Dim externalLinkFile = FileUtil.GetFileInfo("Workbooks", "Tables.xlsx")
            Dim externalWorkbook = p.Workbook.ExternalLinks.AddExternalWorkbook(externalLinkFile)

            Dim ws = p.Workbook.Worksheets.Add("Sheet1")
            'You can access individual cells like this using the index of the external reference in brackets...
            '[1] reference to the the first item in the ExternalReferences collection. This is the externalWorkbook. Index property
            ws.Cells("A1:C3").Formula = "[1]SimpleTable!A1"

            'You can also reference a value and set a format. Here we use the index property instead of hardcoding it in the formula.
            ws.Cells("F1").Formula = $"[{externalWorkbook.Index}]Slicer!F38"
            ws.Cells("F1").Style.Numberformat.Format = "yyyy-MM-dd"

            'Now, Calculate. As the workbook is loaded EPPlus will use the actual values in the package.
            ws.Calculate()

            'The cache stores cell values that are referenced when calculating formulas in the workbook, so formulas can be calculated without access to the external workbook.
            externalWorkbook.UpdateCache()
            Dim f38Value = externalWorkbook.CachedWorksheets("Slicer").CellValues("F38").Value
            Console.WriteLine($"Value of cached value in {externalWorkbook.File.Name} worksheet Slicer cell F38 is : {f38Value}")

            ws.Cells.AutoFitColumns()
            Console.WriteLine($"Cell F1 with an external link has value: {ws.Cells("F1").Value} - formatted: {ws.Cells("F1").Text}")
        End Sub
        Private Sub AddWorksheetWithExternalReferencesInFormulas(ByVal p As ExcelPackage)
            Dim externalLinkFile = FileUtil.GetFileInfo("1.01-GettingStarted.xlsx")
            Dim externalWorkbook = p.Workbook.ExternalLinks.AddExternalWorkbook(externalLinkFile)

            Dim ws = p.Workbook.Worksheets.Add("Sheet2")

            ws.Cells("A1").Value = "Quantity*Price:"
            ws.Cells("B1:B3").Formula = "[2]Inventory!C2*[2]Inventory!D2"  'Here we reference the second external reference, so index is [2]

            ws.Cells("B4").Formula = "Sum(B1:B3)"
            ws.Cells("C4").Formula = "[2]Inventory!E5"

            ws.Cells("A4").Value = "SUM:"
            ws.Cells("A4").AddComment("Sum of external cells matches the sum from cell E5 in the original workbook", "EPPlus Software")
            ws.Cells("B4:C4").Style.Font.Bold = True
            ws.Cells("B4:C4").Style.Numberformat.Format = "#,##0"

            ws.Calculate()

            'If you only want to update the cache you can use externalWorkbook.UpdateCache();
            'Note: The cache is updated when the package is saved if externalWorkbook.CacheStatue is eExternalWorkbookCacheStatus.NotUpdated.
            '      If the cache status is LoadedFromPackage or Updated, you must make sure to update the cache, if necessary, before saving the packge.
            externalWorkbook.UpdateCache()

            ws.Cells.AutoFitColumns()
        End Sub
        Private Sub BreakLinks(ByVal sampleFile1 As FileInfo)
            'If you want to break links to a workbook, you simply remove it from the ExteralLinks collection.
            'This will remove any formulas referencing the workbook and leave the values in the cells. Defined names referencing an external workbook will be set to #REF!

            Using p = New ExcelPackage(sampleFile1)
                Console.WriteLine($"Now break links to the workbook {p.Workbook.ExternalLinks(0).As.ExternalWorkbook.File.FullName}")
                p.Workbook.ExternalLinks.RemoveAt(0)

                Dim newFile = FileUtil.GetFileInfo("1.09-ExternalLinks-Link1Removed.xlsx")
                p.SaveAs(newFile)
            End Using
        End Sub

    End Module
End Namespace
