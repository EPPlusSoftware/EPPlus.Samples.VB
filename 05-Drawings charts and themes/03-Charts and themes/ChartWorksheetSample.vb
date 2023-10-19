Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class ChartWorksheetSample
        Inherits ChartSampleBase
        Public Shared Sub Add(ByVal package As ExcelPackage)
            Dim wsData = LoadBubbleChartData(package)

            'Add a bubble chart worksheet on the data with one serie per row. 
            Dim wsChart = package.Workbook.Worksheets.AddChart("Bubble Chart", eChartType.Bubble)
            Dim chart = wsChart.Chart.As.Chart.BubbleChart
            For row = 2 To 7
                Dim serie = chart.Series.Add(wsData.Cells(row, 2), wsData.Cells(row, 3), wsData.Cells(row, 4))
                serie.HeaderAddress = wsData.Cells(row, 1)
            Next

            chart.DataLabel.Position = eLabelPosition.Center
            chart.DataLabel.ShowSeriesName = True
            chart.DataLabel.ShowBubbleSize = True
            chart.Title.Text = "Sales per Region"
            chart.XAxis.Title.Text = "Total Sales"
            chart.XAxis.Title.Font.Size = 12
            chart.XAxis.MajorGridlines.Width = 1
            chart.YAxis.Title.Text = "Sold Units"
            chart.YAxis.Title.Font.Size = 12
            chart.Legend.Position = eLegendPosition.Bottom

            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.BubbleChartStyle10)
        End Sub
    End Class
End Namespace
