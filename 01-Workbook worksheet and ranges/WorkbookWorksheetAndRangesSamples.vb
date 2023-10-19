﻿Imports EPPlusSamples.WorkbookWorksheetAndRanges
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Module WorkbookWorksheetAndRangesSamples
        Public Async Function RunAsync() As Task
            ' Sample 1.1 - Simply creates a new workbook from scratch
            ' containing a worksheet that adds a few numbers together 
            Call CreateASimpleWorkbook.Run()

            ' Sample 1.2 - Simply reads some values from the file generated by sample 1
            ' and outputs them to the console
            Call ReadWorkbookSample.Run()

            'Sample 1.3 - Load and save using async methods
            Await UsingAsyncAwaitSample.RunAsync()

            'Sample 1.4 - Copy, Fill and Sort Ranges
            Call CopyFillAndSort.Run()

            'Sample 1.5 - Notes/Comments and Threaded comments
            Call CommentsSample.Run()

            ' Sample 1.6 - creates a workbook from scratch 
            'Shows how to use Ranges, Styling, Namedstyles and Hyperlinks
            Call SalesReportFromDatabase.Run()

            'Sample 1.7
            'This sample shows the performance capabilities of the component and shows sheet protection.
            'Load X(param 2) rows with five columns
            PerformanceAndProtectionSample.Run(65534)

            'Sample 1.8 - Linq
            'Opens Sample 1.7 and perform some Linq queries
            Call ReadDataUsingLinq.Run()

            'Sample 1.9 - Add references to external workbooks
            Call ExternalLinksSample.Run()

            'Sample 1.10 - Ignore cell errors using the IngnoreErrors Collection
            Call IgnoreErrorsSample.Run()
        End Function
    End Module
End Namespace