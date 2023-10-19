Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing.Chart.Style
Imports OfficeOpenXml.Style
Imports System
Imports System.Data.SQLite

Namespace EPPlusSamples.FormulaCalculation
    Public Module DynamicArrayFromTableWithChart
        Public Sub Run()
            Using package = New ExcelPackage()
                Dim sheet1 = package.Workbook.Worksheets.Add("Data")

                ' load currency rates from database
                Using conn = New SQLiteConnection(ConnectionString)
                    conn.Open()
                    Dim command = conn.CreateCommand()
                    command.CommandText = "SELECT codeFrom as 'From Currency', codeTo as 'To Currency', date as Date, rate as Rate FROM CurrencyRate"
                    Dim reader = command.ExecuteReader()
                    Dim tableRange = sheet1.Cells("A1").LoadFromDataReader(reader, True, "currencyTable", Table.TableStyles.Medium1)
                    tableRange.SkipRows(1).TakeSingleColumn(2).Style.Numberformat.Format = "yyyy-MM-dd"

                    Dim sheet2 = package.Workbook.Worksheets.Add("Add dynamic formula")
                    sheet2.Cells("A1").Formula = "CONCATENATE(""USD-"",B3)"
                    sheet2.Cells("A1").Style.Font.Bold = True
                    ' add input field for currency
                    sheet2.Cells("A3").Value = "Currency"
                    Dim validation = sheet2.Cells("B3").DataValidation.AddListDataValidation()
                    validation.Formula.Values.Add("CNY")
                    validation.Formula.Values.Add("DKK")
                    validation.Formula.Values.Add("INR")
                    validation.Formula.Values.Add("EUR")
                    validation.Formula.Values.Add("SEK")
                    sheet2.Cells("B3").Value = "DKK"
                    sheet2.Cells("B3").Style.Border.BorderAround(ExcelBorderStyle.Medium)
                    sheet2.Cells("B3").Style.Fill.PatternType = ExcelFillStyle.Solid
                    sheet2.Cells("B3").Style.Fill.BackgroundColor.SetColor(Drawing.Color.LightGray)
                    sheet2.Cells("B3").Style.Font.Bold = True

                    'Set a dynamic formula to the table headers.
                    sheet2.Cells("A5").Formula = "Data!A1:D1"
                    sheet2.Cells("A5:D5").Style.Font.Bold = True

                    ' Here we use the FILTER function to get all USD-DKK rates
                    ' from the imported table.                    
                    sheet2.Cells("A6").Formula = $"FILTER(currencyTable[], currencyTable[To Currency]=B3)"
                    ' Dynamic array formulas must always be calculated before saving the workbook.
                    sheet2.Calculate()

                    ' The FormulaAddress property contains the range used by the dynamic
                    ' array formula after calculation. The variable fa will be used to refer
                    ' to address of the dynamic array formulas result range.
                    Dim fa = sheet2.Cells("A6").FormulaRange
                    fa.TakeSingleColumn(2).Style.Numberformat.Format = "yyyy-MM-dd"
                    ' Now let's add a chart for the filtered array (initially showing USD-DKK rates)
                    Dim chart = sheet2.Drawings.AddLineChart("Dynamic Chart", eLineChartType.Line)
                    chart.Title.LinkedCell = sheet2.Cells("B3")
                    Dim series = chart.Series.Add(fa.TakeSingleColumn(3), fa.TakeSingleColumn(2))
                    series.Header = "Rate"

                    'Add conditional formatting for each currency in the filtered data.
                    AddConditionalNumberFormat(sheet2.Cells("D5:D1000"), "$B5=""CNY""", "[$¥-804]#,##0.00")
                    AddConditionalNumberFormat(sheet2.Cells("D5:D1000"), "$B5=""DKK""", "#,##0.00\ [$kr.-406]")
                    AddConditionalNumberFormat(sheet2.Cells("D5:D1000"), "$B5=""EUR""", "#,##0.00\ [$€-1]")
                    AddConditionalNumberFormat(sheet2.Cells("D5:D1000"), "$B5=""INR""", "[$₹-4009]\ #,##0.00")
                    AddConditionalNumberFormat(sheet2.Cells("D5:D1000"), "$B5=""SEK""", "#,##0.00\ [$kr-41D]")

                    chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle7)

                    chart.SetPosition(1, 0, 6, 0)
                    chart.SetSize(200)
                    sheet1.Cells.AutoFitColumns()
                    sheet2.Cells.AutoFitColumns()

                    sheet2.Select("B3", True)
                End Using
                package.SaveAs(FileUtil.GetCleanFileInfo("6.2-DynamicArrayFormulasWithChart.xlsx"))
            End Using

        End Sub

        Private Sub AddConditionalNumberFormat(ByVal range As ExcelRangeBase, ByVal formula As String, ByVal numberFormat As String)
            Dim cf1 = range.ConditionalFormatting.AddExpression()
            cf1.Formula = formula
            cf1.Style.NumberFormat.Format = numberFormat
        End Sub
    End Module
End Namespace
