Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class ScatterChartSample
        Inherits ChartSampleBase
        Public Shared Sub Add(ByVal package As ExcelPackage)
            'Adda a scatter chart on the data with one serie per row. 
            Dim ws = package.Workbook.Worksheets.Add("Scatter Chart")

            Dim fullRange = CreateIceCreamData(ws)
            Dim range = fullRange.SkipRows(1)

            Dim chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter)
            chart.SetPosition(1, 0, 3, 0)
            chart.To.Column = 18
            chart.To.Row = 20
            chart.XAxis.Format = "yyyy-mm"
            chart.XAxis.Title.Text = "Period"
            chart.XAxis.MajorGridlines.Width = 1
            chart.YAxis.Format = "$#,##0"
            chart.YAxis.Title.Text = "Sales"

            chart.Legend.Position = eLegendPosition.Bottom

            Dim serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie.HeaderAddress = ws.Cells("A1")
            Dim tr = serie.TrendLines.Add(eTrendLine.MovingAverage)
            tr.Name = "Icecream Sales-Monthly Average"
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ScatterChartStyle12)
        End Sub
    End Class
End Namespace
