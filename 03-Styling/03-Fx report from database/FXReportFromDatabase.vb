' ***********************************************************************************************
' Required Notice: Copyright (C) EPPlus Software AB. 
' This software is licensed under PolyForm Noncommercial License 1.0.0 
' and may only be used for noncommercial purposes 
' https://polyformproject.org/licenses/noncommercial/1.0.0/
' 
' A commercial license to use this software can be purchased at https://epplussoftware.com
' ************************************************************************************************
' Date               Author                       Change
' ************************************************************************************************
' 01/27/2020         EPPlus Software AB           Initial release EPPlus 5
' ***********************************************************************************************
Imports System
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports System.Data.SQLite
Imports System.Linq

Namespace EPPlusSamples.Styling
    Friend Class FxReportFromDatabase
        Public Class FxRates
            Public Property [Date] As Date?
            Public Property UsdSek As Double
            Public Property UsdEur As Double
            Public Property UsdInr As Double
            Public Property UsdCny As Double
            Public Property UsdDkk As Double
        End Class
        ''' <summary>
        ''' This sample creates a new workbook from a template file containing a chart and populates it with Exchange rates from 
        ''' the database and set the three series on the chart.
        ''' </summary>
        ''' <paramname="connectionString">Connectionstring to the db</param>
        ''' <paramname="template">the template</param>
        ''' <paramname="outputdir">output dir</param>
        ''' <returns></returns>
        Public Shared Sub Run()
            Console.WriteLine("Running sample 3.3")
            Dim template = FileUtil.GetFileInfo("Workbooks", "3.3-GraphTemplate.xlsx")

            Using p As ExcelPackage = New ExcelPackage(template, True)
                'Set up the headers
                Dim ws = p.Workbook.Worksheets(0)
                ws.Cells("A20").Value = "Date"
                ws.Cells("B20").Value = "EOD Rate"
                ws.Cells("B20:F20").Merge = True
                ws.Cells("G20").Value = "Change"
                ws.Cells("G20:K20").Merge = True
                ws.Cells("B20:K20").Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous
                Using row = ws.Cells("A20:G20")
                    row.Style.Fill.PatternType = ExcelFillStyle.Solid
                    row.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93))
                    row.Style.Font.Color.SetColor(Color.White)
                    row.Style.Font.Bold = True
                End Using
                ws.Cells("B21").Value = "USD/SEK"
                ws.Cells("C21").Value = "USD/EUR"
                ws.Cells("D21").Value = "USD/INR"
                ws.Cells("E21").Value = "USD/CNY"
                ws.Cells("F21").Value = "USD/DKK"
                ws.Cells("G21").Value = "USD/SEK"
                ws.Cells("H21").Value = "USD/EUR"
                ws.Cells("I21").Value = "USD/INR"
                ws.Cells("J21").Value = "USD/CNY"
                ws.Cells("K21").Value = "USD/DKK"
                Using row = ws.Cells("A21:K21")
                    row.Style.Fill.PatternType = ExcelFillStyle.Solid
                    row.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228))
                    row.Style.Font.Color.SetColor(Color.Black)
                    row.Style.Font.Bold = True
                End Using

                Dim startRow = 22
                'Connect to the database and fill the data
                Using sqlConn = New SQLiteConnection(ConnectionString)
                    Dim row = startRow
                    sqlConn.Open()
                    Dim sql = GetSql()
                    Using sqlCmd = New SQLiteCommand(sql, sqlConn)
                        Using sqlReader = sqlCmd.ExecuteReader()
                            ' get the data and fill rows 22 onwards
                            While sqlReader.Read()
                                ws.Cells(row, 1).Value = sqlReader(0)
                                ws.Cells(row, 2).Value = sqlReader(1)
                                ws.Cells(row, 3).Value = sqlReader(2)
                                ws.Cells(row, 4).Value = sqlReader(3)
                                ws.Cells(row, 5).Value = sqlReader(4)
                                ws.Cells(row, 6).Value = sqlReader(5)
                                row += 1
                            End While
                        End Using
                        'Set the numberformat
                        ws.Cells(startRow, 1, row - 1, 1).Style.Numberformat.Format = "yyyy-mm-dd"
                        ws.Cells(startRow, 2, row - 1, 6).Style.Numberformat.Format = "#,##0.0000"
                        'Set the Formulas 
                        ws.Cells(startRow + 1, 7, row - 1, 11).Formula = $"B${startRow}/B{startRow + 1}-1"
                        ws.Cells(startRow, 7, row - 1, 11).Style.Numberformat.Format = "0.00%"
                    End Using

                    'Set the series for the chart. The series must exist in the template or the program will crash.
                    Dim chart = ws.Drawings("SampleChart").As.Chart.LineChart 'We know the chart is a linechart, so we can use the As.Chart.LineChart Property directly
                    chart.Title.Text = "Exchange rate %"
                    chart.Series(0).Header = "USD/SEK"
                    chart.Series(0).XSeries = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 1, row - 1, 1)
                    chart.Series(0).Series = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 7, row - 1, 7)

                    chart.Series(1).Header = "USD/EUR"
                    chart.Series(1).XSeries = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 1, row - 1, 1)
                    chart.Series(1).Series = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 8, row - 1, 8)

                    chart.Series(2).Header = "USD/INR"
                    chart.Series(2).XSeries = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 1, row - 1, 1)
                    chart.Series(2).Series = "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 9, row - 1, 9)

                    Dim serie = chart.Series.Add("'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 10, row - 1, 10), "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 1, row - 1, 1))
                    serie.Header = "USD/CNY"
                    serie.Marker.Style = eMarkerStyle.None

                    serie = chart.Series.Add("'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 11, row - 1, 11), "'" & ws.Name & "'!" & ExcelCellBase.GetAddress(startRow + 1, 1, row - 1, 1))
                    serie.Header = "USD/DKK"
                    serie.Marker.Style = eMarkerStyle.None

                    chart.Legend.Position = eLegendPosition.Bottom

                    'Set Font bold on USD/EUR in the legend.
                    chart.Legend.Entries(1).Font.Bold = True

                    'Set the chart style
                    chart.StyleManager.SetChartStyle(236)

                    Dim query = (From cell In ws.Cells("A22:A" & (row - 1).ToString()) Select New FxRates With {
.[Date] = cell.GetCellValue(Of Date?)(),
.UsdSek = cell.GetCellValue(Of Double)(1),
.UsdEur = cell.GetCellValue(Of Double)(2),
.UsdInr = cell.GetCellValue(Of Double)(3),
.UsdCny = cell.GetCellValue(Of Double)(4),
.UsdDkk = cell.GetCellValue(Of Double)(5)
}).ToList()
                End Using

                'Get the documet as a byte array from the stream and save it to disk.  (This is useful in a webapplication) ... 
                Dim bin = p.GetAsByteArray()

                Dim file = FileUtil.GetCleanFileInfo("3.3-FxReportFromDatabase.xlsx")
                IO.File.WriteAllBytes(file.FullName, bin)
                Console.WriteLine("Sample 3.3 created: {0}", file.FullName)
                Console.WriteLine()
            End Using
        End Sub

        Private Shared Function GetSql() As String
            Dim sb = New StringBuilder()
            sb.Append("SELECT date,")
            sb.Append("SUM(Case when CodeTo = 'SEK' Then rate Else 0 END) AS [SEK], ")
            sb.Append("SUM(Case when CodeTo = 'EUR' Then rate Else 0 END) AS [EUR], ")
            sb.Append("SUM(Case when CodeTo = 'INR' Then rate Else 0 END) AS [INR], ")
            sb.Append("SUM(Case when CodeTo = 'CNY' Then rate Else 0 END) AS [CNY], ")
            sb.Append("SUM(Case when CodeTo = 'DKK' Then rate Else 0 END) AS [DKK], ")
            sb.Append("SUM(Case when CodeTo = 'USD' Then rate Else 0 END) AS [USD] ")
            sb.Append("FROM CurrencyRate ")
            sb.Append("where [CodeFrom]='USD' AND CodeTo in ('SEK', 'EUR', 'INR','CNY','DKK') ")
            sb.Append("GROUP BY date ")
            sb.Append("ORDER BY date")
            Return sb.ToString()
        End Function
    End Class
End Namespace
