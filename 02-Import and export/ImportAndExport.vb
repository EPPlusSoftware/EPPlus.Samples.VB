Imports EPPlusSamples.ImportAndExport
Imports EPPlusSamples.LoadDataFromCsvFilesIntoTables
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Module ImportAndExportSamples
        Public Async Function RunAsync() As Task
            ToCollectionSample.Run()

            Call LoadingDataFromCollection.Run()

            'Sample 2.3 Loads two csv files into tables and creates an area chart and a Column/Line chart on the data.
            'This sample also shows how to use a secondary axis.
            Await ImportAndExportCsvFilesSample.RunAsync()

            'Sample 2.4 - Import and Export DataTable
            DataTableSample.Run()

            ' Sample 2.5 - Html Export
            'This sample shows basic html export functionality.
            'For more advanced samples using charts see https://samples.epplussoftware.com/HtmlExport
            HtmlTableExportSample.Run()
            Await HtmlRangeExportSample.RunAsync()

            'Sample 32 - Json Export
            'This sample shows the json export functionality.
            'For more a samples exporting to chart librays see https://samples.epplussoftware.com/JsonExport
            Await JsonExportSample.RunAsync()

            ' Sample 2.7 - ToCollection and ToCollectionWithMappings
            ' This sample shows how to export data from a worksheet
            ' to a IEnumerable<T> where T is a class.
            ToCollectionSample.Run()
        End Function
    End Module
End Namespace
