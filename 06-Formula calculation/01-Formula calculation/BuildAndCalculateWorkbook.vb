﻿' ***********************************************************************************************
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
Imports OfficeOpenXml

Namespace EPPlusSamples.FormulaCalculation
    Public Module BuildAndCalculateWorkbook
        Public Sub Run()
            Console.WriteLine("Sample 6.1 - Build and calculate workbook")

            Using package = New ExcelPackage()
                Dim ws1 = package.Workbook.Worksheets.Add("ws1")
                ' Add some values to sum
                ws1.Cells("A1").Formula = "(2*2)/2"
                ws1.Cells("A2").Value = 4
                ws1.Cells("A3").Value = 6
                ws1.Cells("A4").Formula = "SUM(A1:A3)"

                ' calculate all formulas on  the worksheet
                ws1.Calculate()

                ' Print the calculated value
                Console.WriteLine("SUM(A1:A3) evaluated to {0}", ws1.Cells("A4").Value)

                ' Add another worksheet
                Dim ws2 = package.Workbook.Worksheets.Add("ws2")
                ws2.Cells("A1").Value = 3
                ws2.Cells("A2").Formula = "SUM(A1,ws1!A4)"

                ' calculate all formulas in the entire workbook
                package.Workbook.Calculate()

                ' Print the calculated value
                Console.WriteLine("SUM(A1,ws1!A4) evaluated to {0}", ws2.Cells("A2").Value)

                ' Calculate a range
                ws1.Cells("B1").Formula = "IF(TODAY()<DATE(2014,6,1),""BEFORE"" &"" FIRST"",CONCATENATE(""FIRST"","" OF"","" JUNE 2014 OR LATER""))"
                ws1.Cells("B1").Calculate()

                ' Print the calculated value
                Console.WriteLine("IF(TODAY()<DATE(2014,6,1),""BEFORE"" &"" FIRST"",CONCATENATE(""FIRST"","" OF"","" JUNE 2014 OR LATER"")) evaluated to {0}", ws1.Cells("B1").Value)

                ' Evaluate a formula string (without calculate depending cells).
                ' That means that if A1 contains a formula that hasn't been calculated it take the value from a1, blank or zero if it's a new formula.
                ' In this case A1 has been calculated (2), so everything should be ok!
                Const formula = "(2+4)*ws1!A1"
                Dim result = package.Workbook.FormulaParserManager.Parse(formula)

                ' Print the calculated value
                Console.WriteLine("(2+4)*ws1!A2 evaluated to {0}", result)

                ' Evaluate a formula string (Calculate depending cells)
                ' A1 will be recalculated.
                Dim result2 = ws1.Calculate("(2+4)*A1")

            End Using
        End Sub
    End Module
End Namespace
