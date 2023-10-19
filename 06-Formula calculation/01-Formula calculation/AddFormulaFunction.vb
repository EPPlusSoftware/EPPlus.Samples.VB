' ***********************************************************************************************
' Required Notice: Copyright (C) EPPlus Software AB. 
' This software is licensed under PolyForm Noncommercial License 1.0.0 
' and may only be used for noncommercial purposes 
' https://polyformproject.org/licenses/noncommercial/1.0.0/
' 
' A commercial license to use this software can be purchased at https://epplussoftware.com
' ************************************************************************************************
' Date               Author                       Change
' ************************************************************************************************
' 01/27/2020         EPPlus Software AB           Initial release EPPlus 5
' ***********************************************************************************************
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports OfficeOpenXml
Imports OfficeOpenXml.FormulaParsing
Imports OfficeOpenXml.FormulaParsing.Excel.Functions
Imports OfficeOpenXml.FormulaParsing.Excel.Functions.Text
Imports OfficeOpenXml.FormulaParsing.FormulaExpressions

Namespace EPPlusSamples.FormulaCalculation
    ''' <summary>
    ''' This sample shows how to add functions to the FormulaParser of EPPlus.
    ''' 
    ''' For further details on how to build functions, have a look in the EPPlus.FormulaParsing.Excel.Functions namespace
    ''' </summary>
    Public Module AddFormulaFunction
        Public Sub Run()
            Console.WriteLine("Sample 6.1 - AddFormulaFunction")
            Console.WriteLine()
            Using package = New ExcelPackage()
                ' add your function module to the parser
                package.Workbook.FormulaParserManager.LoadFunctionModule(New MyFunctionModule())

                ' Note that if you dont want to write a module, you can also
                ' add new functions to the parser this way:
                ' package.Workbook.FormulaParserManager.AddOrReplaceFunction("sum.addtwo", new SumAddTwo());
                ' package.Workbook.FormulaParserManager.AddOrReplaceFunction("seanconneryfy", new SeanConneryfy());


                'Override the buildin Text function to handle swedish date formatting strings. Excel has localized date format strings with is now supported by EPPlus.
                package.Workbook.FormulaParserManager.AddOrReplaceFunction("text", New TextSwedish())

                ' add a worksheet with some dummy data
                Dim ws = package.Workbook.Worksheets.Add("Test")
                ws.Cells("A1").Value = 1
                ws.Cells("A2").Value = 2
                ws.Cells("P3").Formula = "SUM(A1:A2)"
                ws.Cells("B1").Value = "Hello"
                ws.Cells("C1").Value = New DateTime(2013, 12, 31)
                ws.Cells("C2").Formula = "Text(C1,""åååå-MM-dd"")"   'Swedish formatting
                ' use the added "sum.addtwo" function
                ws.Cells("A4").Formula = "TAXES.VAT(A1:A2,P3)"
                ' use the other function "seanconneryfy"
                ws.Cells("B2").Formula = "REVERSESTRING(B1)"

                ' calculate
                ws.Calculate()

                ' show result
                Console.WriteLine("TAXES.VAT(A1:A2,P3) evaluated to {0}", ws.Cells("A4").Value)
                Console.WriteLine("REVERSESTRING(B1) evaluated to {0}", ws.Cells("B2").Value)
            End Using
        End Sub
    End Module

    Friend Class MyFunctionModule
        Inherits FunctionsModule
        Public Sub New()
            Functions.Add("taxes.vat", New CalculateVat())
            Functions.Add("reversestring", New ReverseString())
        End Sub
    End Class

    ''' <summary>
    ''' A simple function that calculates 25% VAT on the sum of a range.
    ''' </summary>
    Friend Class CalculateVat
        Inherits ExcelFunction
        ''' <summary>
        ''' Sets the minimum number of parameters for the function.
        ''' </summary>
        Public Overrides ReadOnly Property ArgumentMinLength As Integer
            Get
                Return 1
            End Get
        End Property
        Public Overrides Function Execute(ByVal arguments As IList(Of FunctionArgument), ByVal context As ParsingContext) As CompileResult
            Const VatRate = 0.25

            ' Helper method that converts function arguments to an enumerable of doubles
            Dim errorValue As ExcelErrorValue = Nothing
            Dim numbers = MyBase.ArgsToDoubleEnumerable(arguments, context, errorValue)

            If errorValue Is Nothing Then
                ' Do the work
                Dim result = 0R
                Enumerable.ToList(numbers).ForEach(Sub(x) result += x * VatRate)

                ' return the result
                Return CreateResult(result, DataType.Decimal)
            Else
                'return errorValue.AsCompileResult;
                Return New CompileResult(errorValue, DataType.ExcelError)
            End If
        End Function
    End Class
    ''' <summary>
    ''' This function handles Swedish formatting strings.
    ''' </summary>
    Friend Class TextSwedish
        Inherits ExcelFunction
        Public Overrides ReadOnly Property ArgumentMinLength As Integer
            Get
                Return 2
            End Get
        End Property
        Public Overrides Function Execute(ByVal arguments As IList(Of FunctionArgument), ByVal context As ParsingContext) As CompileResult
            'Replace swedish year format with invariant for parameter 2.
            Dim format = arguments(1).Value.ToString().Replace("åååå", "yyyy")
            Dim newArgs = New List(Of FunctionArgument) From {
                arguments.ElementAt(0)
            }
            newArgs.Add(New FunctionArgument(format))

            'Use the build-in Text function.
            Dim func = New Text()
            Return func.Execute(newArgs, context)
        End Function
    End Class

    ''' <summary>
    ''' Reverses a string
    ''' </summary>
    Friend Class ReverseString
        Inherits ExcelFunction
        Public Overrides ReadOnly Property ArgumentMinLength As Integer
            Get
                Return 1
            End Get
        End Property
        Public Overrides Function Execute(ByVal arguments As IList(Of FunctionArgument), ByVal context As ParsingContext) As CompileResult
            ' Get the first arg
            Dim input = ArgToString(arguments, 0)

            ' reverse the string
            Dim charArr = input.ToCharArray()
            Array.Reverse(charArr)

            ' return the result
            Return CreateResult(New String(charArr), DataType.String)
        End Function
    End Class
End Namespace
