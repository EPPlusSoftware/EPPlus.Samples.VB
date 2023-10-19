Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class BlanksErrorsAndDuplicates
        Public Shared Sub Run(ByVal pck As ExcelPackage)
            Dim sheet = pck.Workbook.Worksheets.Add("BlanksAndErrors")

            Dim address = "A1:A20"

            ' -------------------------------------------------------------------
            ' Create a ContainsBlanks rule
            ' -------------------------------------------------------------------
            Dim containsBlanks = sheet.ConditionalFormatting.AddContainsBlanks(address)

            containsBlanks.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.Goldenrod)

            ' -------------------------------------------------------------------
            ' Create a NotContainsBlanks rule
            ' -------------------------------------------------------------------
            Dim noBlanks = sheet.ConditionalFormatting.AddNotContainsBlanks(address)

            noBlanks.Style.Border.Top.Style = ExcelBorderStyle.Double
            noBlanks.Style.Border.Top.Color.Color = Color.ForestGreen

            sheet.Cells("A3:A6").Formula = "Row()"

            ' -------------------------------------------------------------------
            ' Create a ContainsErrors rule
            ' -------------------------------------------------------------------
            Dim containsErrors = sheet.ConditionalFormatting.AddContainsErrors(address)

            'Add a few incorrect formulas
            sheet.Cells("A7").Formula = "I an Invalid Formula"
            sheet.Cells("A8").Formula = "SUM(1,""Nonsense"")"
            'Will show up appropriately but prompts excel to update links on opening the file
            'sheet.Cells["A9"].Formula = "SUM(1,nonExistent!J12)";

            containsErrors.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Red)

            containsErrors.Priority = 1

            ' -------------------------------------------------------------------
            ' Create a NotContainsErrors rule
            ' -------------------------------------------------------------------
            Dim noErrors = sheet.ConditionalFormatting.AddNotContainsErrors(address)

            noErrors.Style.Border.Right.Style = ExcelBorderStyle.Double
            noErrors.Style.Border.Right.Color.Color = Color.Purple

            ' -------------------------------------------------------------------
            ' Create a DuplicateValues rule
            ' -------------------------------------------------------------------
            Dim duplicateValues = sheet.ConditionalFormatting.AddDuplicateValues(address)

            duplicateValues.Style.Fill.Style = eDxfFillStyle.PatternFill
            duplicateValues.Style.Fill.PatternType = ExcelFillStyle.Solid
            duplicateValues.Style.Fill.BackgroundColor.Color = Color.DarkOrange

        End Sub
    End Class
End Namespace
