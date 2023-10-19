Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.ChartEx
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class RegionMapChartSample
        Inherits ChartSampleBase
        Public Shared Async Function Add(ByVal package As ExcelPackage) As Task
            Dim ws = package.Workbook.Worksheets.Add("RegionMapChart")

            Dim range = Await LoadSalesFromDatabase(ws)

            'Region map charts 
            Dim regionChart = ws.Drawings.AddRegionMapChart("RegionMapChart")
            regionChart.Title.Text = "Sales"
            regionChart.SetPosition(1, 0, 6, 0)
            regionChart.SetSize(1200, 600)

            'Set the series address. EPPlus will not create the actual map data for the chart. Excel will do that when the chart is rendered.
            Dim rmSerie = regionChart.Series.Add(ws.Cells(2, 4, range.End.Row, 4), ws.Cells(2, 1, range.End.Row, 3))
            rmSerie.HeaderAddress = ws.Cells("D1")
            rmSerie.ColorBy = eColorBy.Value   'Set how to color the series. This value is set in the select data dialog in Excel.

            'Color settings only apply when ColorBy is set to Value
            rmSerie.Colors.NumberOfColors = eNumberOfColors.ThreeColor
            rmSerie.Colors.MinColor.Color.SetSchemeColor(eSchemeColor.Accent3)
            rmSerie.Colors.MidColor.Color.SetHslColor(180, 50, 50)
            rmSerie.Colors.MidColor.ValueType = eColorValuePositionType.Number
            rmSerie.Colors.MidColor.PositionValue = 500
            rmSerie.Colors.MaxColor.Color.SetRgbPercentageColor(75, 25, 25)
            rmSerie.Colors.MaxColor.ValueType = eColorValuePositionType.Number
            rmSerie.Colors.MaxColor.PositionValue = 1500

            rmSerie.ProjectionType = eProjectionType.Mercator
            regionChart.Legend.Add()
            regionChart.Legend.Position = eLegendPosition.Top
            regionChart.StyleManager.SetChartStyle(ePresetChartStyle.RegionMapChartStyle2)
        End Function
    End Class
End Namespace
