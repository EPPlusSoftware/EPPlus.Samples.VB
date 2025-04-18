﻿' ***********************************************************************************************
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
Imports OfficeOpenXml
Imports System.Threading.Tasks
Imports System.IO

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Friend Class ChartsAndThemesSample
        ''' <summary>
        ''' Sample 5.3 - Creates various charts and apply a theme if supplied in the parameter themeFile.
        ''' </summary>
        Public Shared Async Function RunAsync(ByVal xlFile As FileInfo, ByVal themeFile As FileInfo) As Task(Of String)
            Using package = New ExcelPackage()
                'Load a theme file if set. Thmx files can be exported from Excel. This will change the appearance for the workbook.                
                If themeFile IsNot Nothing Then
                    package.Workbook.ThemeManager.Load(themeFile)
                    ' * Themes can also be altered. For example, uncomment this code to set the Accent1 to a blue color **
                    'package.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetRgbColor(Color.FromArgb(32, 78, 224));
                End If

                ' *******************************************************************************************************
                '  About chart styles: 
                '  
                '  Chart styles can be applied to charts using the Chart.StyleManager.SetChartMethod method.
                '  The chart styles can either be set by the two enums ePresetChartStyle and ePresetChartStyleMultiSeries or by setting the Chart Style Number.
                '  
                '  Note: Chart styles in Excel changes depending on many parameters (like number of series, axis types and more), so the enums will not always reflect the style index in Excel. 
                '  The enums are for the most common scenarios.
                '  If you want to reflect a specific style please use the Chart Style Number for the chart in Excel. 
                '  The chart style number can be fetched by recording a macro in Excel and click the style you want to apply.
                '  
                '  Chart style do not alter visibility of chart objects like data labels or chart titles like Excel do. That must be set in code before setting the style.
                ' *******************************************************************************************************

                'The first method adds a worksheet with four 3D charts with different styles. The last chart applies an exported chart template file (*.crtx) to the chart.
                Await ThreeDimensionalCharts.Add3DCharts(package)

                'This method adds four line charts with different chart elements like up-down bars, error bars, drop lines and high-low lines.
                Await LineChartsSample.Add(package)

                'Adds a scatter chart with a moving average trendline.
                ScatterChartSample.Add(package)

                'Adds a column chart with a legend where we style and remove individual legend items.
                Await ColumnChartWithLegendSample.Add(package)

                'Adds a bubble-chartsheet
                ChartWorksheetSample.Add(package)

                'Adds a radar chart
                RadarChartSample.Add(package)

                'Adds a Volume-High-Low-Close stock chart
                StockChartSample.Add(package)

                'Adds a sunburst and a treemap chart 
                Await SunburstAndTreemapChartSample.Add(package)

                'Adds a box & whisker and a histogram chart 
                BoxWhiskerHistogramChartSample.Add(package)

                ' Adds a waterfall chart
                WaterfallChartSample.Add(package)

                ' Adds a funnel chart
                FunnelChartSample.Add(package)

                Await RegionMapChartSample.Add(package)

                'Add an area chart using a chart template (chrx file)
                Await ChartTemplateSample.AddAreaChart(package)

                'Add a stackedColumn chart with custom labels
                BarColumnChartsWithManualLayout.Add(package)

                'Save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile)
                Return xlFile.FullName
            End Using
        End Function
    End Class
End Namespace
