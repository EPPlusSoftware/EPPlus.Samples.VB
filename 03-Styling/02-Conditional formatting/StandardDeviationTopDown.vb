Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class StandardDeviationTopDown
        Public Shared Sub Run(ByVal pck As ExcelPackage)
            Dim worksheet = pck.Workbook.Worksheets.Add("StdDev_TopBottom")

            worksheet.Cells("B1:B43").Formula = "ROW()"

            ' -------------------------------------------------------------------
            ' Create a Above StdDev rule
            ' -------------------------------------------------------------------
            Dim zeroDeviation = worksheet.ConditionalFormatting.AddAboveStdDev(New ExcelAddress("B1:B43"))
            zeroDeviation.StdDev = 0

            zeroDeviation.Style.Font.Bold = True

            ' -------------------------------------------------------------------
            ' Create a Below StdDev rule
            ' -------------------------------------------------------------------
            Dim twoDeviation = worksheet.ConditionalFormatting.AddBelowStdDev(New ExcelAddress("B1:B43"))

            twoDeviation.StdDev = 2
            twoDeviation.Style.Fill.PatternType = ExcelFillStyle.Solid
            twoDeviation.Style.Fill.BackgroundColor.Color = Color.ForestGreen

            'Make a single cell actually exist at stdev 2
            worksheet.Cells("B14").Value = -177

            ' -------------------------------------------------------------------
            ' Create a Bottom rule
            ' -------------------------------------------------------------------
            Dim bottomRank4 = worksheet.ConditionalFormatting.AddBottom(New ExcelAddress("B1:B43"))

            bottomRank4.Rank = 4

            bottomRank4.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.MediumVioletRed)

            ' -------------------------------------------------------------------
            ' Create a Bottom Percent rule
            ' -------------------------------------------------------------------
            Dim bottomPercent15 = worksheet.ConditionalFormatting.AddBottomPercent(New ExcelAddress("B1:B43"))

            bottomPercent15.Rank = 15

            bottomPercent15.Style.Fill.PatternType = ExcelFillStyle.Solid
            bottomPercent15.Style.Fill.BackgroundColor.Color = Color.DeepSkyBlue

            ' -------------------------------------------------------------------
            ' Create a Top rule
            ' -------------------------------------------------------------------
            Dim top = worksheet.ConditionalFormatting.AddTop(New ExcelAddress("B1:B43"))
            top.Style.Fill.PatternType = ExcelFillStyle.Solid
            top.Style.Fill.BackgroundColor.Color = Color.MediumPurple

            ' -------------------------------------------------------------------
            ' Create a Top Percent rule
            ' -------------------------------------------------------------------
            Dim topPercent = worksheet.ConditionalFormatting.AddTopPercent(New ExcelAddress("B1:B43"))

            topPercent.Style.Border.Left.Style = ExcelBorderStyle.Thin
            topPercent.Style.Border.Left.Color.Theme = eThemeSchemeColor.Text2
            topPercent.Style.Border.Bottom.Style = ExcelBorderStyle.DashDot
            topPercent.Style.Border.Bottom.Color.SetColor(ExcelIndexedColor.Indexed8)
            topPercent.Style.Border.Right.Style = ExcelBorderStyle.Thin
            topPercent.Style.Border.Right.Color.Color = Color.Blue
            topPercent.Style.Border.Top.Style = ExcelBorderStyle.Hair
            topPercent.Style.Border.Top.Color.Auto = True

            worksheet.Cells.AutoFitColumns()
        End Sub
    End Class
End Namespace
