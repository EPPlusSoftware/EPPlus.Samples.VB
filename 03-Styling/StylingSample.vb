Imports EPPlusSamples.Styling
Imports System.IO
Imports System.Reflection

Namespace EPPlusSamples
    Public Class StylingBasics
        Public Shared Sub Run()
            'Sample 2.1 - Basic styling            
            Call BasicStyleSample.Run()

            ' Sample 2.2 - Conditional Formatting
            ConditionalFormattingSamples.Run()

            ' Sample 2.3 - Creates a workbook based on a template.
            ' Populates a range with data and set the series of a linechart.
            Call FxReportFromDatabase.Run()

            'Sample 2.4
            'Creates an advanced report on a directory in the filesystem.
            'Parameter 2 is the directory to report. Parameter 3 is how deep the scan will go. Parameter 4 Skips Icons if set to true (The icon handling is slow)
            'This example demonstrates how to use outlines, tables,comments, shapes, pictures and charts.                

            Dim directoryToList = New DirectoryInfo(Assembly.GetEntryAssembly().Location).Parent
            CreateAFileSystemReport.Run(directoryToList, 5, True)
        End Sub
    End Class
End Namespace
