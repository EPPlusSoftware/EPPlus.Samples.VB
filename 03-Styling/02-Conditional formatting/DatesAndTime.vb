Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class DatesAndTime
        Public Shared Sub Run(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("DatesAndTimeExamples")

            AddDatesAndTimesToSheet(sheet)

            ' -------------------------------------------------------------------
            ' Create a Last 7 Days rule
            ' -------------------------------------------------------------------
            'ExcelAddress timePeriodAddress = new ExcelAddress("B21:E34 A11:A20");
            Dim timePeriodAddress As ExcelAddress = New ExcelAddress("A1:K40")

            Dim last7Days = sheet.ConditionalFormatting.AddLast7Days("A1:A40")

            last7Days.Style.Fill.PatternType = ExcelFillStyle.LightTrellis
            last7Days.Style.Fill.PatternColor.Color = Color.BurlyWood
            last7Days.Style.Fill.BackgroundColor.Color = Color.LightCyan

            ' -------------------------------------------------------------------
            ' Create a Last Week rule
            ' -------------------------------------------------------------------
            Dim lastWeek = sheet.ConditionalFormatting.AddLastWeek("B1:B40")
            'lastWeek.Style.NumberFormat.Format = "YYYY";
            lastWeek.Style.Fill.PatternType = ExcelFillStyle.Solid
            lastWeek.Style.Fill.BackgroundColor.Color = Color.Orange

            ' -------------------------------------------------------------------
            ' Create a This Week rule
            ' -------------------------------------------------------------------
            Dim thisWeek = sheet.ConditionalFormatting.AddThisWeek("B1:B40")

            thisWeek.Style.Fill.PatternType = ExcelFillStyle.Solid
            thisWeek.Style.Fill.BackgroundColor.Color = Color.YellowGreen

            ' -------------------------------------------------------------------
            ' Create a Next Week rule
            ' -------------------------------------------------------------------
            Dim nextWeek = sheet.ConditionalFormatting.AddNextWeek("B1:B40")

            nextWeek.Style.Fill.PatternType = ExcelFillStyle.Solid
            nextWeek.Style.Fill.BackgroundColor.Color = Color.ForestGreen

            ' -------------------------------------------------------------------
            ' Create a Today rule
            ' -------------------------------------------------------------------
            Dim today = sheet.ConditionalFormatting.AddToday("C1:C40")

            today.Style.Fill.PatternType = ExcelFillStyle.Solid
            today.Style.Fill.BackgroundColor.Color = Color.Gold

            today.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Gold)

            ' -------------------------------------------------------------------
            ' Create a Tomorrow rule
            ' -------------------------------------------------------------------
            Dim tomorrow = sheet.ConditionalFormatting.AddTomorrow("C1:C40")

            tomorrow.Style.Fill.PatternType = ExcelFillStyle.Solid
            tomorrow.Style.Fill.BackgroundColor.Color = Color.LightSkyBlue

            tomorrow.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.Violet)

            ' -------------------------------------------------------------------
            ' Create a Yesterday rule
            ' -------------------------------------------------------------------
            Dim yesterday = sheet.ConditionalFormatting.AddYesterday("C1:C40")

            yesterday.Style.Fill.PatternType = ExcelFillStyle.Solid
            yesterday.Style.Fill.BackgroundColor.Color = Color.DimGray

            yesterday.Style.Border.BorderAround(ExcelBorderStyle.Dashed, Color.DarkRed)

            ' -------------------------------------------------------------------
            ' Create a Last Month rule
            ' -------------------------------------------------------------------
            Dim lastMonth = sheet.ConditionalFormatting.AddLastMonth("E1:E40")

            'lastMonth.Style.NumberFormat.Format = "YYYY";
            lastMonth.Style.Fill.PatternType = ExcelFillStyle.Solid
            lastMonth.Style.Fill.BackgroundColor.Color = Color.OrangeRed

            ' -------------------------------------------------------------------
            ' Create a This Month rule
            ' -------------------------------------------------------------------
            Dim thisMonth = sheet.ConditionalFormatting.AddThisMonth("F1:F40")

            thisMonth.Style.Fill.PatternType = ExcelFillStyle.Solid
            thisMonth.Style.Fill.BackgroundColor.Color = Color.LightGoldenrodYellow

            ' -------------------------------------------------------------------
            ' Create a Next Month rule
            ' -------------------------------------------------------------------
            Dim nextMonth = sheet.ConditionalFormatting.AddNextMonth("G1:G40")

            nextMonth.Style.Fill.PatternType = ExcelFillStyle.Solid
            nextMonth.Style.Fill.BackgroundColor.Color = Color.ForestGreen

            sheet.Cells.AutoFitColumns()
        End Sub

        Private Shared Sub AddDatesAndTimesToSheet(ByVal sheet As ExcelWorksheet)
            Dim startOffset As Integer = Date.Now.DayOfWeek
            Dim lastWeekDate = Date.Now.AddDays(-7 - startOffset)
            Dim year = $"{Date.Now.Year}"

            Dim lastMonth = $"{year}-{Date.Now.AddMonths(-1).Month}-"
            Dim thisMonth = $"{year}-{Date.Now.Month}-"
            Dim nextMonth = $"{year}-{Date.Now.AddMonths(+1).Month}-"

            For i = 1 To 10
                sheet.Cells(i, 1).Value = lastWeekDate.AddDays(i - 1).ToShortDateString()
                sheet.Cells(i + 7, 1).Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString()
                sheet.Cells(i + 14, 1).Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString()

                sheet.Cells(i, 2).Value = lastWeekDate.AddDays(i - 1).ToShortDateString()
                sheet.Cells(i + 7, 2).Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString()
                sheet.Cells(i + 14, 2).Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString()

                sheet.Cells(i, 3).Value = lastWeekDate.AddDays(i - 1).ToShortDateString()
                sheet.Cells(i + 7, 3).Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString()
                sheet.Cells(i + 14, 3).Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString()

                sheet.Cells(i, 5).Value = lastMonth & $"{i + 10}"
                sheet.Cells(i + 7, 5).Value = thisMonth & $"{i + 10}"
                sheet.Cells(i + 14, 5).Value = nextMonth & $"{i + 10}"

                sheet.Cells(i, 6).Value = lastMonth & $"{i + 10}"
                sheet.Cells(i + 7, 6).Value = thisMonth & $"{i + 10}"
                sheet.Cells(i + 14, 6).Value = nextMonth & $"{i + 10}"

                sheet.Cells(i, 7).Value = lastMonth & $"{i + 10}"
                sheet.Cells(i + 7, 7).Value = thisMonth & $"{i + 10}"
                sheet.Cells(i + 14, 7).Value = nextMonth & $"{i + 10}"
            Next
        End Sub
    End Class
End Namespace
