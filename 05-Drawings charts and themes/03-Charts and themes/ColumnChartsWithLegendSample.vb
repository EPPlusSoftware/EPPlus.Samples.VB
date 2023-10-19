Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Drawing
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class ColumnChartWithLegendSample
        Inherits ChartSampleBase
        Public Shared Async Function Add(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("ColumnCharts")

            Dim fullRange = Await LoadFromDatabase(ws)
            Dim range = fullRange.SkipRows(1)

            'Add a line chart
            Dim chart = ws.Drawings.AddBarChart("ColumnChartWithLegend", eBarChartType.ColumnStacked)
            Dim serie1 = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie1.Header = "Order Value"
            Dim serie2 = chart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0))
            serie2.Header = "Tax"
            Dim serie3 = chart.Series.Add(range.TakeSingleColumn(3), range.TakeSingleColumn(0))
            serie3.Header = "Freight"
            chart.SetPosition(0, 0, 6, 0)
            chart.SetSize(1200, 400)
            chart.Title.Text = "Column chart"

            'Set style 10
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ColumnChartStyle10)

            chart.Legend.Entries(0).Font.Fill.Color = Color.Red
            chart.Legend.Entries(1).Font.Fill.Color = Color.Green
            chart.Legend.Entries(2).Deleted = True

            range.AutoFitColumns(0)
        End Function
    End Class
End Namespace
