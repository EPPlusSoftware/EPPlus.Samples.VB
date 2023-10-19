Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class LineChartsSample
        Inherits ChartSampleBase
        Public Shared Async Function Add(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("LineCharts")

            Dim fullRange = Await LoadFromDatabase(ws)
            Dim range = fullRange.SkipRows(1) ' remove the headers row

            'Add a line chart
            Dim chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line)
            'var serie = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            Dim serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie.Header = "Order Value"
            chart.SetPosition(0, 0, 6, 0)
            chart.SetSize(1200, 400)
            chart.Title.Text = "Line Chart With Droplines"
            chart.AddDropLines()
            chart.DropLine.Border.Width = 2
            'Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12)

            'Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithErrorBars", eLineChartType.Line)
            serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie.Header = "Order Value"
            chart.SetPosition(21, 0, 6, 0)
            chart.SetSize(1200, 400)   'Make this chart wider to make room for the datatable.
            chart.Title.Text = "Line Chart With Error Bars"
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Percentage)
            serie.ErrorBars.Value = 5
            chart.PlotArea.CreateDataTable()

            'Set style 2
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle2)

            'Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithUpDownBars", eLineChartType.Line)
            Dim serie1 = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie1.Header = "Order Value"
            Dim serie2 = chart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0))
            serie2.Header = "Tax"
            Dim serie3 = chart.Series.Add(range.TakeSingleColumn(3), range.TakeSingleColumn(0))
            serie3.Header = "Freight"
            chart.SetPosition(42, 0, 6, 0)
            chart.SetSize(1200, 400)
            chart.Title.Text = "Line Chart With Up/Down Bars"
            chart.AddUpDownBars(True, True)

            'Set style 10, Note: As this is a line chart with multiple series, we use the enum for multiple series. Charts with multiple series usually has a subset of of the chart styles in Excel.
            'Another option to set the style is to use the Excel Style number, in this case 236: chart.StyleManager.SetChartStyle(236)
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.LineChartStyle9)
            range.AutoFitColumns(0)


            'Add a line chart with high/low Bars
            chart = ws.Drawings.AddLineChart("LineChartWithHighLowLines", eLineChartType.Line)
            serie1 = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0))
            serie1.Header = "Order Value"
            serie2 = chart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0))
            serie2.Header = "Tax"
            serie3 = chart.Series.Add(range.TakeSingleColumn(3), range.TakeSingleColumn(0))
            serie3.Header = "Freight"
            chart.SetPosition(63, 0, 6, 0)
            chart.SetSize(1200, 400)
            chart.Title.Text = "Line Chart With High/Low Lines"
            chart.AddHighLowLines()

            'Set the style using the Excel ChartStyle number. The chart style must exist in the ExcelChartStyleManager.StyleLibrary[]. 
            'Styles can be added and removed from this library. By default it is loaded with the styles for EPPlus supported chart types.
            chart.StyleManager.SetChartStyle(237)
            range.AutoFitColumns(0)
        End Function
    End Class
End Namespace
