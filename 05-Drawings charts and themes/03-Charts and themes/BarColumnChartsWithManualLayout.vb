Imports OfficeOpenXml
Imports System
Imports OfficeOpenXml.Drawing.Chart
Imports System.Drawing
Imports OfficeOpenXml.Table

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class BarColumnChartsWithManualLayout
        Public Shared Sub Add(package As ExcelPackage)
            Dim cSheet = package.Workbook.Worksheets.Add("ColumnChartSheet")

            Dim range = cSheet.Cells("A1:C3")
            Dim table = cSheet.Tables.Add(range, "DataTable")
            table.ShowHeader = False
            table.SyncColumnNames(ApplyDataFrom.ColumnNamesToCells, True)

            range.Formula = "ROW() + COLUMN()"

            cSheet.Calculate()

            Dim sChart = cSheet.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked)

            sChart.Series.Add(cSheet.Cells("A1:A3"))
            sChart.Series.Add(cSheet.Cells("B1:B3"))
            sChart.Series.Add(cSheet.Cells("C1:C3"))

            Dim highestSeriesOfColumns = sChart.Series(2)

            Dim dataLabelRulesForEntireRow = highestSeriesOfColumns.DataLabel

            Dim topColumnInStack = dataLabelRulesForEntireRow.DataLabels.Add(0)
            topColumnInStack.Fill.Style = Drawing.eFillStyle.SolidFill
            topColumnInStack.Fill.SolidFill.Color.SetRgbColor(Color.MediumPurple)

            Dim middleColumnInSecondStack = sChart.Series(1).DataLabel.DataLabels.Add(1)

            Dim bottomColumnInFirstStack = sChart.Series(0).DataLabel.DataLabels.Add(0)
            bottomColumnInFirstStack.Fill.Style = Drawing.eFillStyle.SolidFill
            bottomColumnInFirstStack.Fill.SolidFill.Color.SetRgbColor(Color.CornflowerBlue)

            Dim lastTop = sChart.Series(2).DataLabel.DataLabels.Add(2)

            SetShowValues(topColumnInStack)
            SetShowValues(middleColumnInSecondStack)
            SetShowValues(bottomColumnInFirstStack)
            SetShowValues(lastTop)

            Dim manualLayoutTop = topColumnInStack.Layout.ManualLayout
            Dim middleColumnLayout = middleColumnInSecondStack.Layout.ManualLayout
            Dim bottomColumnLayout = bottomColumnInFirstStack.Layout.ManualLayout
            Dim lastTopLayout = lastTop.Layout.ManualLayout

            'Set x and y position (units are in percent of chart width/height)
            'Left means pushing 'from' the left. Same with top. It's the position of the left side of the element box.
            manualLayoutTop.Left = -5
            manualLayoutTop.Top = -25

            'Set textbox width
            manualLayoutTop.Width = 5
            manualLayoutTop.Height = 10

            'Set x only
            middleColumnLayout.Left = 10

            '''By default left and top are offsets to the starting Position of the element. AKA: dataLabel.Position
            '''To make positioning easier you can also define starting position as an offset from the Left or Top edge of the chart:
            bottomColumnLayout.TopMode = eLayoutMode.Edge
            bottomColumnLayout.LeftMode = eLayoutMode.Edge

            'This will put the label in the top left corner of the chart itself.
            bottomColumnLayout.Left = 0
            bottomColumnLayout.Top = 0

            'Note that when in edge mode, negative inputs are nonsensical as 0 is the starting point and negative values would be outside the chart.
            'Forcing a label outside the chart resets it to its Position attribute in excel.
            lastTopLayout.TopMode = eLayoutMode.Edge
            lastTopLayout.Top = -5
        End Sub

        Private Shared Sub SetShowValues(item As ExcelChartDataLabelItem)
            item.ShowLegendKey = False
            item.ShowValue = True
            item.ShowCategory = False
            item.ShowSeriesName = False
            item.ShowPercent = False
            item.ShowBubbleSize = False
            item.ShowLeaderLines = True
        End Sub
    End Class
End Namespace
