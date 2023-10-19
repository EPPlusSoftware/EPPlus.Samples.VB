Imports EPPlusSamples.FiltersAndValidations
Imports System
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Class FiltersAndValidation
        Public Shared Async Function RunAsync() As Task
            'Sample 12 - Data validation
            Call DataValidationSample.Run()

            'Sample 13 - Filter
            Console.WriteLine("Running sample 13-Filter")
            Await Filter.RunAsync()
            Console.WriteLine("Sample 13 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Function
    End Class
End Namespace
