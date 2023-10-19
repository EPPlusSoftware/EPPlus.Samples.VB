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
' 10/03/2023         EPPlus Software AB           Initial release EPPlus 7
' ***********************************************************************************************
Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.Data.SQLite

Namespace EPPlusSamples.FormulaCalculation
    Public Class UsingArrayformulas
        Public Shared Sub Run()
            Dim file = FileUtil.GetCleanFileInfo("6.2-ArrayFormulas.xlsx")
            Using xlPackage As ExcelPackage = New ExcelPackage(file)
                DynamicArrayFormulasSample(xlPackage)
                LegacyArrayFormulasSample(xlPackage)
                ImpicitIntersectionSample(xlPackage)
                UseDynamicArrayFormulaOutputRange(xlPackage)
                xlPackage.Save()
            End Using
        End Sub

        Private Shared Sub DynamicArrayFormulasSample(ByVal xlPackage As ExcelPackage)
            Dim worksheet = LoadData(xlPackage, "Dynamic Array Formulas")
            '***********************************************************************
            ' A dynamic array formula is always set in a single cell 
            ' that can return an array that spills bottom-right.
            '***********************************************************************

            'Create a number of different dynamic array formulas
            worksheet.Cells("J1").Value = "New Ordernumber"
            worksheet.Cells("J2").Formula = "G2:G251+10000"   'Create a dynamic array formula in cell J2 that spills downwards, adding 10000 to order number.

            worksheet.Cells("L1").Value = "Is Last Quater"
            worksheet.Cells("L2").Formula = "Month(F2:F251) > 9"   'Returns a boolean if the last quarter of the year.

            worksheet.Cells("N1").Value = "Sorted Name"
            worksheet.Cells("O1").Value = "Company Name"
            worksheet.Cells("P1").Value = "Country"

            worksheet.Cells("N2").Formula = "SORT(CHOOSECOLS(A2:F251,2,1,3,6))"

            worksheet.Cells("J1:AA1").Style.Font.Bold = True

            worksheet.Calculate()

            'When we have calculated the workbook we can get the dynamic formulas output range via the FormulaRange property or the Worksheet's GetFormulaRange method.
            Dim sortedRange = worksheet.Cells("N2").FormulaRange
            sortedRange.Offset(0, 3, sortedRange.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd"
            worksheet.Cells.AutoFitColumns()
        End Sub
        ''' <summary>
        ''' This sample shows how to create and calculate legacy array formulas.
        ''' Note: If not needed for backward compatibility, do not use legacy array formulas, use dynamic array formulas instead.
        ''' </summary>
        ''' <paramname="xlPackage"></param>
        Private Shared Sub LegacyArrayFormulasSample(ByVal xlPackage As ExcelPackage)
            Dim worksheet = LoadData(xlPackage, "Legacy Array Formulas")

            'This sample created the same legacy array formulas as the dynamic array formulas sample.
            'In most cases you only use legacy array formulas if you need them for backward compatibility.
            worksheet.Cells("J1").Value = "New Ordernumber"
            worksheet.Cells("J2:J251").CreateArrayFormula("G2:G251+10000")   'Create a dynamic array formula in cell J2 that spills downwards, adding 10000 to order number.

            worksheet.Cells("L1").Value = "Is Last Quater"
            worksheet.Cells("L2:L251").CreateArrayFormula("Month(F2:F251) > 9")   'Returns a boolean if the last quarter of the year.

            worksheet.Cells("N1").Value = "Sorted Name"
            worksheet.Cells("O1").Value = "Company Name"
            worksheet.Cells("P1").Value = "Country"

            'Using dynamic functions in a legacy array formula works, but in this case it's a lot easier to use a dynamic array formula as you in many cases don't know the size of the output array.
            worksheet.Cells("N2:P251").CreateArrayFormula("SORT(CHOOSECOLS(A2:F251,2,1,3,6))")

            worksheet.Cells("J1:AA1").Style.Font.Bold = True

            worksheet.Calculate()

            'When we have calculated the workbook we can get the array formulas output range via the FormulaRange property or the Worksheet's GetFormulaRange method.
            Dim sortedRange = worksheet.Cells("N2").FormulaRange
            sortedRange.Offset(0, 3, sortedRange.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd"
            worksheet.Cells.AutoFitColumns()
        End Sub

        Private Shared Function LoadData(ByVal xlPackage As ExcelPackage, ByVal worksheetName As String) As ExcelWorksheet
            Dim worksheet = xlPackage.Workbook.Worksheets.Add(worksheetName)
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand(OrdersSql, sqlConn)
                    Dim reader = sqlCmd.ExecuteReader()
                    worksheet.Cells("A1").LoadFromDataReader(reader, True, $"OrderDataTable{xlPackage.Workbook.Worksheets.Count + 1}", TableStyles.Dark1)
                End Using
                worksheet.Cells("E2:E251,G2:G251").Style.Numberformat.Format = "#,##0"
                worksheet.Cells("F2:F251").Style.Numberformat.Format = "yyyy-MM-dd"
            End Using

            Return worksheet
        End Function
        Private Shared Sub ImpicitIntersectionSample(ByVal xlPackage As ExcelPackage)
            Dim worksheet = LoadData(xlPackage, "Implicit Intersection")

            'If you creates a shared formula (a formula shared over more than one cell.) that returns an array, implicit intersection will allways be used.
            'This means that the value in the column that intesects with the formula will be returned. An @ will be added before the formula in Excel.
            worksheet.Cells("J2:J251").Formula = "$E$2:$E$251+10000"

            'By default EPPlus creates a dynamic array formula if you set a single cell to a forumla that outputs an array. To use Implicit Itersection instead set the UseImplicitItersection on the cell.
            'Note that this property has to be set after the formula is set.
            worksheet.Cells("L2").Formula = "$E$2:$E$251+10000"
            worksheet.Cells("L2").UseImplicitItersection = True

            worksheet.Calculate()

            worksheet.Cells.AutoFitColumns()

        End Sub

        Private Shared Sub UseDynamicArrayFormulaOutputRange(ByVal xlPackage As ExcelPackage)
            Dim worksheet = xlPackage.Workbook.Worksheets.Add("Using dynamic formula output")
            worksheet.Cells("A1").Formula = "RANDARRAY(5,5,1,10,TRUE)"
            worksheet.Cells("A7").Value = "SUM:"
            ' this corresponds to SUM(A1#) in Excel...
            worksheet.Cells("B7").Formula = "SUM(ANCHORARRAY(A1))"
            worksheet.Cells("A8").Value = "AVERAGE:"
            worksheet.Cells("B8").Formula = "AVERAGE(ANCHORARRAY(A1))"
            worksheet.Calculate()
        End Sub
    End Class
End Namespace
