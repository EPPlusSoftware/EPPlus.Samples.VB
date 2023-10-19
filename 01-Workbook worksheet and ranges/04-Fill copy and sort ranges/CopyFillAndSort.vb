Imports System

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module CopyFillAndSort
        Public Sub Run()
            Console.WriteLine("Running sample 1.4 - Copy, Fill and Sort Ranges")
            Call CopyRangeSample.Run()
            Call FillRangeSample.Run()
            Call SortingRangesSample.Run()
            Console.WriteLine("Sample 1.4 finished.")
            Console.WriteLine()
        End Sub
    End Module
End Namespace
