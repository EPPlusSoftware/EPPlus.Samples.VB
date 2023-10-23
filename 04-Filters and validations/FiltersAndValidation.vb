Imports EPPlusSamples.FiltersAndValidations
Imports System
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Class FiltersAndValidation
        Public Shared Async Function RunAsync() As Task
            'Sample 4.2 - Data validation
            Call DataValidationSample.Run()

            'Sample 4.2 - Filter
            Console.WriteLine("Running sample 4.2-Filter")
            Await Filter.RunAsync()
            Console.WriteLine("Sample 4.2 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Function
    End Class
End Namespace
