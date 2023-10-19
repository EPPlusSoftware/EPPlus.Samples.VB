Imports EPPlusSamples.PivotTables
Imports EPPlusSamples.TablesPivotTablesAndSlicers
Imports System
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Module TablesPivotTableAndSlicersSample
        Public Async Function RunAsync() As Task
            'Sample 7.1 - Custom Named Table, Pivot Table and Slicer styles
            Console.WriteLine("Running sample 7.1 - Working with tables")
            Await TablesSample.RunAsync()
            Console.WriteLine("Sorting tables sample...")
            Await SortingTablesSample.RunAsync()
            Console.WriteLine("Sample 7.1 finished.")
            Console.WriteLine()

            'Sample 7.2 - Table slicers and Pivot table slicers
            Call SlicerSample.Run()

            'sample 7.3 - pivot tables
            'This sample demonstrates how to create and work with pivot tables.
            Call PivotTablesSample.Run()
            'The second class demonstrates how to style you pivot table.
            Call PivotTablesStylingSample.Run()

            'Sample 7.4 - Custom Named Table, Pivot Table and Slicer styles
            Call CustomTableSlicerStyleSample.Run()
        End Function
    End Module
End Namespace
