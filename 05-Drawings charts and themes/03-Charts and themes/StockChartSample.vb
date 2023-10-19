Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports System
Imports System.Collections.Generic

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Friend Class StockChartSample
        Inherits ChartSampleBase
        Public Shared Sub Add(ByVal package As ExcelPackage)
            'Adda a scatter chart on the data with one serie per row. 
            Dim ws = package.Workbook.Worksheets.Add("Stock Chart")

            CreateStockData(ws)

            Dim chart = ws.Drawings.AddStockChart("StockChart1", eStockChartType.StockVHLC, ws.Cells("A2:E11"))
            chart.SetPosition(1, 0, 6, 0)
            chart.To.Column = 28
            chart.To.Row = 25
            'The first chart type is the bar chart containng one serie
            chart.PlotArea.ChartTypes(0).Series(0).HeaderAddress = ws.Cells("B1")
            'The second chart type is the stock chart containing three series
            chart.Series(0).HeaderAddress = ws.Cells("C1")
            chart.Series(1).HeaderAddress = ws.Cells("D1")
            chart.Series(2).HeaderAddress = ws.Cells("E1")
            chart.Series(2).TrendLines.Add(eTrendLine.MovingAverage) 'Add a moving average trend line on the close price.
            chart.Legend.Position = eLegendPosition.Right

            chart.Title.Text = "Fiction Inc"
            chart.StyleManager.SetChartStyle(ePresetChartStyle.StockChartStyle10)
        End Sub
        Public Shared Sub CreateStockData(ByVal ws As ExcelWorksheet)
            Dim list = New List(Of TradingData)() From {
    New TradingData() With {
        .[Date] = New DateTime(2019, 12, 30),
        .Volume = 1000,
        .LowPrice = 99,
        .HighPrice = 101,
        .ClosePrice = 100
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 2),
        .Volume = 700,
        .LowPrice = 97.4,
        .HighPrice = 100,
        .ClosePrice = 98.7
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 3),
        .Volume = 400,
        .LowPrice = 98.4,
        .HighPrice = 99.3,
        .ClosePrice = 99.1
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 6),
        .Volume = 1100,
        .LowPrice = 99.1,
        .HighPrice = 105.6,
        .ClosePrice = 105.6
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 7),
        .Volume = 900,
        .LowPrice = 104.3,
        .HighPrice = 105.6,
        .ClosePrice = 104.8
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 8),
        .Volume = 1500,
        .LowPrice = 100.3,
        .HighPrice = 104.8,
        .ClosePrice = 101.1
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 9),
        .Volume = 1200,
        .LowPrice = 101.1,
        .HighPrice = 111.3,
        .ClosePrice = 111.3
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 10),
        .Volume = 900,
        .LowPrice = 111.3,
        .HighPrice = 115.3,
        .ClosePrice = 114.4
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 13),
        .Volume = 800,
        .LowPrice = 107.4,
        .HighPrice = 114.4,
        .ClosePrice = 108.1
    },
    New TradingData() With {
        .[Date] = New DateTime(2020, 1, 14),
        .Volume = 1150,
        .LowPrice = 105.4,
        .HighPrice = 110.1,
        .ClosePrice = 110.1
    }
}
            ws.Cells("A1").LoadFromCollection(list, True)
            ws.Cells("A1:E1").Style.Font.Bold = True
            ws.Cells("A2:A11").Style.Numberformat.Format = "yyyy-MM-dd"
            ws.Cells("B2:B11").Style.Numberformat.Format = "#,##0"
            ws.Cells("C2:E11").Style.Numberformat.Format = "$#,##0.00"
            ws.Cells.AutoFitColumns()
        End Sub
    End Class

    Friend Class TradingData
        Public Property [Date] As Date
        Public Property Volume As Double
        Public Property LowPrice As Double
        Public Property HighPrice As Double
        Public Property ClosePrice As Double
    End Class
End Namespace
