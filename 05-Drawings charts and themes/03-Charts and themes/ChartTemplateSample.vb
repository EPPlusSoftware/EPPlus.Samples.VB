Imports OfficeOpenXml
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class ChartTemplateSample
        Inherits ChartSampleBase
        Public Shared Async Function AddAreaChart(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("Area chart from template")
            Dim range = Await LoadFromDatabase(ws)

            'Adds an Area chart from a template file. The crtx file has it's own theme, so it does not change with the theme.
            'The As property provides an easy type cast for drawing objects
            Dim areaChart = ws.Drawings.AddChartFromTemplate(FileUtil.GetFileInfo("05-Drawings charts and themes\03-Charts and themes", "AreaChartStyle3.crtx"), "areaChart").As.Chart.AreaChart
            Dim areaSerie = areaChart.Series.Add(ws.Cells(2, 2, 16, 2), ws.Cells(2, 1, 16, 1))
            areaSerie.Header = "Order Value"
            areaChart.SetPosition(1, 0, 6, 0)
            areaChart.SetSize(1200, 400)
            areaChart.Title.Text = "Area Chart"

            range.AutoFitColumns(0)
        End Function
    End Class
End Namespace
