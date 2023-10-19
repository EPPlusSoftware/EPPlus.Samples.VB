Imports EPPlusSamples.FormulaCalculation
Imports System

Namespace EPPlusSamples
    Public Module FormulaCalculationSample
        Public Sub Run()
            'Sample 6 Calculate - Shows how to calculate formulas in the workbook.
            Console.WriteLine("Sample 6.1 - Calculate formulas")
            CalculateFormulasSample.Run()
            Console.WriteLine("Sample 6.1 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()

            Console.WriteLine("Sample 6.2 - Dynamic array formulas")
            Call UsingArrayformulas.Run()
            Call DynamicArrayFromTableWithChart.Run()
            Console.WriteLine("Sample 6.2 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub
    End Module
End Namespace
