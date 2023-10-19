Imports EPPlusSamples.DrawingsChartsAndThemes
Imports System
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Friend Class DrawingsChartsAndThemesSample
        Public Shared Async Function RunAsync() As Task
            'Sample 5.1 - Shapes & Images
            Call ShapesAndImagesSample.Run()

            'Sample 5.2
            'Open the file in sample 1.1 and add a pie chart.
            Call OpenWorkbookAndAddDataAndChartSample.Run()

            Console.WriteLine("Running sample 5.3-Theme and Chart styling")
            'Sample 5.3 - Themes and Chart styling
            'Run the sample with the default office theme
            Await ChartsAndThemesSample.RunAsync(FileUtil.GetFileInfo("5.3-ChartsAndThemes.xlsx"), Nothing)

            'Run the sample with the integral theme. Themes can be exported as thmx files from Excel and can then be applied to a package.
            Await ChartsAndThemesSample.RunAsync(FileUtil.GetFileInfo("5.3-ChartsAndThemes-IntegralTheme.xlsx"), FileUtil.GetFileInfo("05-Drawings charts and themes\03-Charts and themes", "integral.thmx"))
            Console.WriteLine("Sample 5.3 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()

            'Sample 5.4 - Shows how to add sparkline charts.
            Call SparkLinesSample.Run()

            'Sample 26 - Form Controls & Drawing Groups
            Call FormControlsSample.Run()
        End Function
    End Class
End Namespace
