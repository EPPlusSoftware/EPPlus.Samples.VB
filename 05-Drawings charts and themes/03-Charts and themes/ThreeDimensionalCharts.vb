Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class ThreeDimensionalCharts
        Inherits ChartSampleBase
        Public Shared Async Function Add3DCharts(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("3D Charts")

            Dim fullRange = Await LoadFromDatabase(ws)
            Dim range = fullRange.SkipRows(1)

            'Add a column chart
            Dim chart = ws.Drawings.AddBarChart("column3dChart", eBarChartType.ColumnClustered3D)
            Dim serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie.Header = "Order Value"
            chart.SetPosition(0, 0, 6, 0)
            chart.SetSize(1200, 400)
            chart.Title.Text = "Column Chart 3D"

            'Set style 9 and Colorful Palette 3. 
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Column3dChartStyle9, ePresetChartColors.ColorfulPalette3)

            'Add a line chart
            Dim lineChart = ws.Drawings.AddLineChart("line3dChart", eLineChartType.Line3D)
            Dim lineSerie = lineChart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            lineSerie.Header = "Order Value"
            lineChart.SetPosition(21, 0, 6, 0)
            lineChart.SetSize(1200, 400)
            lineChart.Title.Text = "Line 3D"
            'Set Line3D Style 1
            lineChart.StyleManager.SetChartStyle(ePresetChartStyle.Line3dChartStyle1)

            'Add a bar chart
            chart = ws.Drawings.AddBarChart("bar3dChart", eBarChartType.BarStacked3D)
            serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie.Header = "Order Value"
            serie = chart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0))
            serie.Header = "Tax"
            serie = chart.Series.Add(range.TakeSingleColumn(3), range.TakeSingleColumn(0))
            serie.Header = "Freight"

            chart.SetPosition(42, 0, 6, 0)
            chart.SetSize(1200, 600)
            chart.Title.Text = "Bar Chart 3D"
            'Set the color
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.StackedBar3dChartStyle7, ePresetChartColors.ColorfulPalette1)

            range.AutoFitColumns(0)
        End Function
    End Class
End Namespace
