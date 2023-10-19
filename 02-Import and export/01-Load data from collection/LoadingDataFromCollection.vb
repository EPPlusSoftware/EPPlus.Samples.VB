Imports EPPlusSamples.LoadingData
Imports System

Namespace EPPlusSamples.ImportAndExport
    Public Module LoadingDataFromCollection
        Public Sub Run()
            'Sample 4 - Shows a few ways to load data (Datatable, IEnumerable and more).
            Console.WriteLine("Running sample 4 - LoadingDataWithTables")
            Call LoadingDataWithTablesSample.Run()
            Console.WriteLine("Sample 2.1 (LoadingDataWithTables) created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()

            'Sample 4 - Shows how to load dynamic/ExpandoObject 
            Call LoadingDataWithDynamicObjects.Run()
            Console.WriteLine("Sample 2.1 (LoadingDataWithDynamicObjects) created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()

            ' Sample 4 - LoadFromCollectionWithAttributes
            Call LoadingDataFromCollectionWithAttributes.Run()
            Console.WriteLine("Sample 2.1 (LoadingDataFromCollectionWithAttributes) created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub
    End Module
End Namespace
