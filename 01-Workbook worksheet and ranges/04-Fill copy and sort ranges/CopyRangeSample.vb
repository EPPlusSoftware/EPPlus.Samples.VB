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
' 10/13/2021         EPPlus Software AB           Initial release EPPlus 5
' ***********************************************************************************************

Imports OfficeOpenXml

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module CopyRangeSample
        'This sample demonstrates how to copy entire worksheet, ranges and how to exclude different cell properties.
        Public Sub Run()
            Using p = New ExcelPackage()
                Dim sourceFile = FileUtil.GetFileInfo("Workbooks", "1.03-Salesreport.xlsx")
                Dim sourcePackage = New ExcelPackage(sourceFile)
                Dim sourceWs = sourcePackage.Workbook.Worksheets(0)

                'Copy the entire source worksheet to a new worksheet.
                CopyEntireWorksheet(p, sourceWs)
                'Copy a range from the source worksheet to the new worksheet.
                'This samples demonstrates how to exclude different options to exclude different parts of the cell properties
                CopyRange(p, sourceWs)
                'Copy a range with values only, removing any formula.
                CopyValues(p)
                'Copy styles 
                CopyStyles(p, sourceWs)

                p.SaveAs(FileUtil.GetCleanFileInfo("1.04-CopyRangeSamples.xlsx"))
            End Using
        End Sub

        Private Sub CopyValues(ByVal p As ExcelPackage)
            Dim ws = p.Workbook.Worksheets.Add("CopyValues")
            'Add some numbers and formulas and calculate the worksheet
            ws.Cells("A1:A10").FillNumber(1)
            ws.Cells("B1:B9").Formula = "A1+A2"
            ws.Cells("B10").Formula = "Sum(B1:B9)"
            ws.Calculate()

            'Now, copy the values starting at cell D1 without the formulas.
            ws.Cells("A1:B10").Copy(ws.Cells("D1"), ExcelRangeCopyOptionFlags.ExcludeFormulas)
        End Sub

        Private Sub CopyEntireWorksheet(ByVal p As ExcelPackage, ByVal sourceWs As ExcelWorksheet)
            'To copy the entire worksheet just add the source worksheet as parameter 2 when adding the new worksheet.
            p.Workbook.Worksheets.Add("CopySalesReport", sourceWs)
        End Sub

        Private Sub CopyRange(ByVal p As ExcelPackage, ByVal sourceWs As ExcelWorksheet)
            Dim ws = p.Workbook.Worksheets.Add("CopyRangeOfReport")

            'Use the first 10 rows of the sales report in sample 8 as the source.
            Dim sourceRange = sourceWs.Cells("A1:G10")

            'Copy the source full range starting from C1.
            'Copy always start from the top left cell of the destination range and copies the full range.
            sourceRange.Copy(ws.Cells("C1"))

            'Copy the same source range to C15 and exclude the hyperlinks.
            'We also remove the Hyperlink style from the range containing the hyperlinks.
            sourceRange.Copy(ws.Cells("C15"), ExcelRangeCopyOptionFlags.ExcludeHyperLinks)
            ws.Cells("D19:D24").StyleName = "Normal"

            'Copy the values only, excluding merged cells, styles and hyperlinks.
            sourceRange.Copy(ws.Cells("C30"), ExcelRangeCopyOptionFlags.ExcludeMergedCells, ExcelRangeCopyOptionFlags.ExcludeStyles, ExcelRangeCopyOptionFlags.ExcludeHyperLinks)

            'Copy styles and merged cells, excluding values and hyperlinks.
            sourceRange.Copy(ws.Cells("C45"), ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeHyperLinks)

            ws.Cells.AutoFitColumns()
        End Sub
        Private Sub CopyStyles(ByVal p As ExcelPackage, ByVal sourceWs As ExcelWorksheet)
            Dim ws = p.Workbook.Worksheets.Add("CopyStyles")

            'Create a new random report 
            FillRangeWithRandomData(ws)

            'Copy the styles from the sales report.
            'If the destination range is larger that the source range styles are filled down and right using the last column/row the source range of the source range.
            sourceWs.Cells("A1:G5").CopyStyles(ws.Cells("A1:G50"))

            ws.Cells.AutoFitColumns()
        End Sub

        Private Sub FillRangeWithRandomData(ByVal ws As ExcelWorksheet)
            ws.Cells("A1").Value = "EPPlus"
            ws.Cells("A2").Value = "New Random Report"
            ws.Cells("A4").Value = "Color"
            ws.Cells("B4").Value = "Category"
            ws.Cells("C4").Value = "Country"
            ws.Cells("D4").Value = "Id"
            ws.Cells("E4").Value = "Date"
            ws.Cells("F4").Value = "Amount"
            ws.Cells("G4").Value = "Currency"

            ws.Cells("A5:A50").FillList(New String() {"Red", "Green", "Blue", "Pink", "Black"})
            ws.Cells("B5:B50").FillList(New String() {"New", "Old"})
            ws.Cells("C5:C50").FillList(New String() {"Usa", "France", "India"})

            ws.Cells("D5:D50").FillNumber(1, 10)
            ws.Cells("E5:E50").FillDateTime(Date.Today)
            ws.Cells("F5:F50").FillNumber(1000, 50)
            ws.Cells("G5:G50").FillList(New String() {"USD", "EUR", "INR"})
        End Sub
    End Module
End Namespace
