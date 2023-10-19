Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class BoxWhiskerHistogramChartSample
        Public Shared Sub Add(ByVal package As ExcelPackage)
            Dim ws = package.Workbook.Worksheets.Add("BoxAndWhiskerChart")
            AddBoxWhiskerData(ws)

            Dim boxWhiskerChart = ws.Drawings.AddBoxWhiskerChart("BoxAndWhisker1")
            Dim bwSerie1 = boxWhiskerChart.Series.Add(ws.Cells(2, 1, 11, 1), Nothing)
            bwSerie1.HeaderAddress = ws.Cells("A1")
            Dim bwSerie2 = boxWhiskerChart.Series.Add(ws.Cells(2, 2, 11, 2), Nothing)
            bwSerie2.HeaderAddress = ws.Cells("B1")
            Dim bwSerie3 = boxWhiskerChart.Series.Add(ws.Cells(2, 3, 11, 3), Nothing)
            bwSerie3.HeaderAddress = ws.Cells("C1")
            boxWhiskerChart.SetPosition(1, 0, 6, 0)
            boxWhiskerChart.SetSize(800, 800)
            boxWhiskerChart.Title.Text = "Number series"
            boxWhiskerChart.XAxis.Deleted = True               'Don't show the X-Axis
            boxWhiskerChart.StyleManager.SetChartStyle(ePresetChartStyle.BoxWhiskerChartStyle3)

            Dim histogramChart = ws.Drawings.AddHistogramChart("Pareto", True)
            histogramChart.SetPosition(1, 0, 19, 0)
            histogramChart.SetSize(800, 800)
            histogramChart.Title.Text = "Histogram with Pareto line"
            Dim hgSerie = histogramChart.Series.Add(ws.Cells(2, 3, 15, 3), Nothing)
            hgSerie.HeaderAddress = ws.Cells("C1")
            hgSerie.Binning.Size = 4
            histogramChart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramChartStyle2)
        End Sub
        Private Shared Sub AddBoxWhiskerData(ByVal ws As ExcelWorksheet)
            ws.Cells("A1").Value = "Primes"
            ws.Cells("A2").Value = 2
            ws.Cells("A3").Value = 3
            ws.Cells("A4").Value = 5
            ws.Cells("A5").Value = 7
            ws.Cells("A6").Value = 11
            ws.Cells("A7").Value = 13
            ws.Cells("A8").Value = 17
            ws.Cells("A9").Value = 19
            ws.Cells("A10").Value = 23
            ws.Cells("A11").Value = 29
            ws.Cells("A12").Value = 31
            ws.Cells("A13").Value = 37
            ws.Cells("A14").Value = 41
            ws.Cells("A15").Value = 43

            ws.Cells("B1").Value = "Even"
            ws.Cells("B2").Value = 2
            ws.Cells("B3").Value = 4
            ws.Cells("B4").Value = 6
            ws.Cells("B5").Value = 8
            ws.Cells("B6").Value = 10
            ws.Cells("B7").Value = 12
            ws.Cells("B8").Value = 14
            ws.Cells("B9").Value = 16
            ws.Cells("B10").Value = 18
            ws.Cells("B11").Value = 20
            ws.Cells("B12").Value = 22
            ws.Cells("B13").Value = 24
            ws.Cells("B14").Value = 26
            ws.Cells("B15").Value = 28

            ws.Cells("C1").Value = "Random"
            ws.Cells("C2").Value = 2
            ws.Cells("C3").Value = 3
            ws.Cells("C4").Value = 7
            ws.Cells("C5").Value = 12
            ws.Cells("C6").Value = 15
            ws.Cells("C7").Value = 18
            ws.Cells("C8").Value = 19
            ws.Cells("C9").Value = 23
            ws.Cells("C10").Value = 25
            ws.Cells("C11").Value = 30
            ws.Cells("C12").Value = 35
            ws.Cells("C13").Value = 37
            ws.Cells("C14").Value = 40
            ws.Cells("C15").Value = 42
        End Sub
    End Class
End Namespace
