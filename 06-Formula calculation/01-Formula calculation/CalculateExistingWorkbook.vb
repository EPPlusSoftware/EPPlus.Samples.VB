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
Imports OfficeOpenXml
Imports System
Imports System.IO

Namespace EPPlusSamples.FormulaCalculation
    ''' <summary>
    ''' This sample demonstrates the formula calculation engine of EPPlus by opening an existing
    ''' workbook and calculate the formulas in it.
    ''' </summary>
    Public Module CalculateExistingWorkbook
        Private Sub RemoveCalculatedFormulaValues(ByVal workbook As ExcelWorkbook)
            For Each worksheet In workbook.Worksheets
                For Each cell In worksheet.Cells
                    ' if there is a formula in the cell, the following code keeps the formula but clears the calculated value.
                    If Not String.IsNullOrEmpty(cell.Formula) Then
                        Dim formula = cell.Formula
                        cell.Value = Nothing
                        cell.Formula = formula
                    End If
                Next
            Next
        End Sub

        Public Sub Run()
            'var resourceStream = GetResource("EPPlusSampleApp.Core.FormulaCalculation.FormulaCalcSample.xlsx");
            Dim filePath = FileUtil.GetFileInfo("06-Formula calculation\01-Formula calculation", "FormulaCalcSample.xlsx").FullName
            Using package = New ExcelPackage(New FileInfo(filePath))
                ' Read the value from the workbook. This is calculated by Excel.
                Dim totalSales As Double? = package.Workbook.Worksheets("Sales").Cells("E10").GetValue(Of Double?)()
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.Value)

                ' This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook)

                ' totalSales from cell C10 should now be empty
                totalSales = package.Workbook.Worksheets("Sales").Cells("E10").GetValue(Of Double?)()
                Console.WriteLine("Total sales read from Cell E10: {0}", If(totalSales.HasValue, totalSales.Value.ToString(), "null"))


                ' ************** 1. Calculate the entire workbook **************
                package.Workbook.Calculate()

                ' totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets("Sales").Cells("E10").GetValue(Of Double?)()
                Console.WriteLine("Total sales read from Cell E10: {0}", If(totalSales.HasValue, totalSales.Value.ToString(), "null"))

                ' ************** 2. Calculate a worksheet **************

                ' This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook)

                package.Workbook.Worksheets("Sales").Calculate()

                ' totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets("Sales").Cells("E10").GetValue(Of Double?)()
                Console.WriteLine("Total sales read from Cell E10: {0}", If(totalSales.HasValue, totalSales.Value.ToString(), "null"))

                ' ************** 3. Calculate a range **************

                ' This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook)

                package.Workbook.Worksheets("Sales").Cells("E10").Calculate()

                ' totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets("Sales").Cells("E10").GetValue(Of Double?)()
                Console.WriteLine("Total sales read from Cell E10: {0}", If(totalSales.HasValue, totalSales.Value.ToString(), "null"))
            End Using

        End Sub
    End Module
End Namespace
