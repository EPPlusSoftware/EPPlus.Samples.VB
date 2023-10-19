Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class RadarChartSample
        Inherits ChartSampleBase
        Public Shared Sub Add(ByVal package As ExcelPackage)
            Dim ws = package.Workbook.Worksheets.Add("RadarChart")

            Dim dt = GetCarDataTable()
            Dim fullRange = ws.Cells("A1").LoadFromDataTable(dt, True)
            Dim range = fullRange.SkipRows(1)
            range.AutoFitColumns()

            Dim chart = ws.Drawings.AddRadarChart("RadarChart1", eRadarChartType.RadarFilled)
            For col = 1 To fullRange.Columns - 1
                Dim serie = chart.Series.Add(range.TakeSingleColumn(col), range.TakeSingleColumn(0))
                serie.HeaderAddress = fullRange.TakeSingleCell(0, col)
            Next

            chart.Legend.Position = eLegendPosition.Top
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.RadarChartStyle4)

            'If you want to apply custom styling do that after setting the chart style so its not overwritten.
            chart.Legend.Effect.SetPresetShadow(ePresetExcelShadowType.OuterTopLeft)

            chart.SetPosition(0, 0, 6, 0)
            chart.To.Column = 17
            chart.To.Row = 30
        End Sub

    End Class
End Namespace
