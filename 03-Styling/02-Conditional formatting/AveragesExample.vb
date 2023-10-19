Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class AveragesExample
        Public Shared Sub Run(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("AverageExamples")

            sheet.Cells("A1:B21").Formula = "ROW()"

            ' -------------------------------------------------------------------
            ' Create an Above Average rule
            ' -------------------------------------------------------------------
            Dim above = sheet.ConditionalFormatting.AddAboveAverage(New ExcelAddress("A1:B21"))

            'Properties allow you to change the formatting of conditional formattings
            'Multiple font properties can be changed to alter the apperance of the formatting
            above.Style.Font.Bold = True
            above.Style.Font.Color.Color = Color.Red
            above.Style.Font.Strike = True

            ' -------------------------------------------------------------------
            ' Create an Above Or Equal Average rule
            ' -------------------------------------------------------------------
            Dim aboveOrEqual = sheet.ConditionalFormatting.AddAboveOrEqualAverage(New ExcelAddress("A1:A21"))

            'Other properties like style can change background color
            aboveOrEqual.Style.Fill.PatternType = ExcelFillStyle.Solid
            aboveOrEqual.Style.Fill.BackgroundColor.Color = Color.DarkBlue

            ' -------------------------------------------------------------------
            ' Create a Below Average rule
            ' -------------------------------------------------------------------
            Dim belowAverage = sheet.ConditionalFormatting.AddBelowAverage(New ExcelAddress("A1:B21"))

            belowAverage.Style.Fill.PatternType = ExcelFillStyle.Solid
            belowAverage.Style.Fill.BackgroundColor.Color = Color.DarkRed

            ' -------------------------------------------------------------------
            ' Create a Below Or Equal Average rule
            ' -------------------------------------------------------------------
            Dim belowOrEqual = sheet.ConditionalFormatting.AddBelowOrEqualAverage(New ExcelAddress("A1:B21"))
            belowOrEqual.Style.Font.Color.Color = Color.White

            belowOrEqual.Style.Fill.PatternType = ExcelFillStyle.Solid
            belowOrEqual.Style.Fill.BackgroundColor.Color = Color.DarkGreen

            'Not that when two properties conflict like belowEqual and aboveEqual on the background color the one with the lowest priority number "wins"
            'Test switching them around and watch the A11 cell closely.
            belowOrEqual.Priority = 2
            aboveOrEqual.Priority = 1

            sheet.Cells.AutoFitColumns()
        End Sub
    End Class
End Namespace
