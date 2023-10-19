Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.ChartEx
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class SunburstAndTreemapChartSample
        Inherits ChartSampleBase
        Public Shared Async Function Add(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("Sunburst & Treemap Chart")
            Dim range = Await LoadSalesFromDatabase(ws)

            Dim sunburstChart = ws.Drawings.AddSunburstChart("SunburstChart1")
            Dim sbSerie = sunburstChart.Series.Add(ws.Cells(2, 4, range.Rows, 4), ws.Cells(2, 1, range.Rows, 3))
            sbSerie.HeaderAddress = ws.Cells("D1")
            sunburstChart.SetPosition(1, 0, 6, 0)
            sunburstChart.SetSize(800, 800)
            sunburstChart.Title.Text = "Sales"
            sunburstChart.Legend.Add()
            sunburstChart.Legend.Position = eLegendPosition.Bottom
            sbSerie.DataLabel.Add(True, True)
            sunburstChart.StyleManager.SetChartStyle(ePresetChartStyle.SunburstChartStyle3)


            Dim treemapChart = ws.Drawings.AddTreemapChart("TreemapChart1")
            Dim tmSerie = treemapChart.Series.Add(ws.Cells(2, 4, range.Rows, 4), ws.Cells(2, 1, range.Rows, 3))
            treemapChart.Title.Font.Fill.Style = eFillStyle.SolidFill
            treemapChart.Title.Font.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Background2)
            tmSerie.HeaderAddress = ws.Cells("D1")
            treemapChart.SetPosition(1, 0, 19, 0)
            treemapChart.SetSize(1000, 800)
            treemapChart.Title.Text = "Sales"
            treemapChart.Legend.Add()
            treemapChart.Legend.Position = eLegendPosition.Right
            tmSerie.DataLabel.Add(True, True)
            tmSerie.ParentLabelLayout = eParentLabelLayout.Banner
            treemapChart.StyleManager.SetChartStyle(ePresetChartStyle.TreemapChartStyle6)
        End Function
    End Class
End Namespace
