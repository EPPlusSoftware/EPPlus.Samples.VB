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
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Table.PivotTable
Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml.Table.PivotTable.Calculation

Namespace EPPlusSamples.PivotTables
    ''' <summary>
    ''' This class shows how to calculate pivottables and fetch data via the CalculatedData propety or the GetPivotData method.
    ''' </summary>
    Public Module PivotTablesCalculationSample
        Public Sub Run()
            Console.WriteLine("Running sample 7.3-Pivot Table Calculation")

            Dim templateFile As FileInfo = EPPlusSamples.FileUtil.GetFileInfo("7.2-PivotTables.xlsx")
            If templateFile.Exists = False Then
                Console.WriteLine("Template file 7.2-PivotTables.xlsx does not exist. Please make sure the sample PivotTablesSample.Run() sample has executed to create this file.")
            End If
            Using pck As ExcelPackage = New ExcelPackage(templateFile)
                Dim pt1 = pck.Workbook.Worksheets("PivotSimple").PivotTables(0)
                Dim pt2 = pck.Workbook.Worksheets("PivotDateGrp").PivotTables(0)
                Dim pt3 = pck.Workbook.Worksheets("PivotWithPageField").PivotTables(0)
                Dim pt4 = pck.Workbook.Worksheets("PivotWithSlicer").PivotTables(0)
                Dim pt5 = pck.Workbook.Worksheets("PivotWithCalculatedField").PivotTables(0)
                Dim pt6 = pck.Workbook.Worksheets("PivotWithCaptionFilter").PivotTables(0)
                Dim pt7 = pck.Workbook.Worksheets("PivotWithShowAsFields").PivotTables(0)
                Dim pt8 = pck.Workbook.Worksheets("PivotSorting")

                'Use calculate on the pivot table to calculate the values that can be accessed via the CalculatedData property or the GetPivotData method.
                'If no calculation has been performed, EPPlus will call this method when using these properties, but will not refresh the pivot cache unless it does not exist.
                'If you have altered data in the pivot source, make sure to call this method with true, to update the pivot cache.
                'Also make sure to calculate any formulas that the pivot table source contains before calculating the data.
                StandardPivotTableSample(pt1)
                DateGroupSample(pt2)
                PageFieldFilterSample(pt3)
                SlicerFilterSample(pt4)
                CalculatedFieldSample(pt5)
                CaptionFilterSample(pt6)
                ShowAsSample(pt7)
                GetPivotDataMethodSample(pck.Workbook.Worksheets("PivotSorting"))
            End Using
        End Sub

        Private Sub StandardPivotTableSample(pt As ExcelPivotTable)
            pt.Calculate(True)

            Dim tot = pt.CalculatedData.GetValue() 'Get the grant total from the pivot table.
            Dim capVerde = pt.CalculatedData.SelectField("Country", "Cape verde").GetValue()

            Console.WriteLine($"The calculated grand total for pivot table {pt.Name} is: {tot:N0}")
            Console.WriteLine($"The total for Cap Verde for pivot table {pt.Name} is: {capVerde:N0}")
            Console.WriteLine()
        End Sub

        Private Sub DateGroupSample(pt As ExcelPivotTable)
            Dim hellenKuhlman = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").GetValue("OrderValue")

            Dim hellenKuhlman2017Q3Tax = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").SelectField("Years", "2017").SelectField("Quarters", "Q3").GetValue("Tax")

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}")
            Console.WriteLine($"The value for Tax for Hellen Kuhlman,Q3 2017 for pivot table {pt.Name} is: {hellenKuhlman2017Q3Tax:N2}")
            Console.WriteLine()
        End Sub

        Private Sub PageFieldFilterSample(pt As ExcelPivotTable)
            Dim hellenKuhlman As Object = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").GetValue("OrderValue")
            Dim hellenKuhlman2017Q4Tax = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").SelectField("OrderDate", "Qtr4").GetValue("Tax")

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}. This value has been filtered by the page field.")
            Console.WriteLine($"The value for Tax for Hellen Kuhlman,Q4 2017 for pivot table {pt.Name} is: {hellenKuhlman2017Q4Tax:N2}")
            Console.WriteLine()
        End Sub

        Private Sub SlicerFilterSample(pt As ExcelPivotTable)
            Dim hellenKuhlman As Object = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").GetValue("OrderValue")
            Dim walschSum = pt.CalculatedData.SelectField("Name", "Walsh LLC").GetValue("OrderValue")

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}. This value has been filtered by the page field.")
            Console.WriteLine($"The value for OrderValue for Walsh LLC,Q4 2017 for pivot table {pt.Name} is: {walschSum:N2}. It is filtered out by the slicer.")
            Console.WriteLine()
        End Sub

        Private Sub CalculatedFieldSample(pt As ExcelPivotTable)
            Dim sengerOrderValue = pt.CalculatedData.SelectField("CompanyName", "Senger LLC").GetValue("OrderValue")

            Dim sengerTax = pt.CalculatedData.SelectField("CompanyName", "Senger LLC").GetValue("Tax")

            Dim sengerFreight = pt.CalculatedData.SelectField("CompanyName", "Senger LLC").GetValue("Freight")

            Dim sengerTotal = pt.CalculatedData.SelectField("CompanyName", "Senger LLC").GetValue("Total")

            Dim grandTotal = pt.CalculatedData.GetValue("Total")

            Console.WriteLine($"Calculated Fields: The value of field OrderValue for Senger LLC for pivot table {pt.Name} is: {sengerOrderValue:N0}.")
            Console.WriteLine($"Calculated Fields: The value of field Tax for Senger LLC for pivot table {pt.Name} is: {sengerTax:N0}.")
            Console.WriteLine($"Calculated Fields: The value of field Freight for Senger LLC for pivot table {pt.Name} is: {sengerFreight:N0}.")
            Console.WriteLine($"Calculated Fields: The value of field Total for Senger LLC for pivot table {pt.Name} is: {sengerTotal:N0}. This field uses the formula [OrderValue]+[Tax]+[Freight]")

            Console.WriteLine($"Calculated Fields: The grand value for OrderValue for  2017 for pivot table {pt.Name} is: {grandTotal:N2}.")
            Console.WriteLine()
        End Sub
        Private Sub CaptionFilterSample(pt As ExcelPivotTable)
            'Sabryna Schulist
            Dim sabrynaSchulistOrderValue = pt.CalculatedData.SelectField("Name", "Sabryna Schulist").GetValue("OrderValue")

            Dim orderDate = New DateTime(2017, 8, 27, 1, 57, 0)

            Dim sabrynaSchulistDateTimeTax = pt.CalculatedData.SelectField("Name", "Sabryna Schulist").SelectField("OrderDate", orderDate).GetValue("Tax")

            'Chelsey Powlowski - is filtered out by the caption filter as the name startes with "C". #REF! will be returened
            Dim chelseyPowlowskiOrderValue = pt.CalculatedData.SelectField("Name", "Chelsey Powlowski").GetValue("OrderValue")

            'Get the grand total for field OrderValue
            Dim grandTotalOrderValue = pt.CalculatedData.GetValue("OrderValue")

            Console.WriteLine($"Caption Filters: The value of field OrderValue for Sabryna Schulist for pivot table {pt.Name} is: {sabrynaSchulistOrderValue:N0}.")
            Console.WriteLine($"Caption Filters: The value of field Tax Sabryna Schulist, {orderDate} for pivot table {pt.Name} is: {sabrynaSchulistDateTimeTax:N0}.")

            Console.WriteLine($"Caption Filters: The value of field OrderValue for Chelsey Powlowski for pivot table {pt.Name} is: {chelseyPowlowskiOrderValue:N0}. This value has been filtered out by the caption filter")
            Console.WriteLine($"Caption Filters: The grand total of field OrderValue for pivot table {pt.Name} is: {grandTotalOrderValue:N0}. All Name's starting with ""C"" has been filtered out by the caption filter")
            Console.WriteLine()
        End Sub
        Private Sub ShowAsSample(pt As ExcelPivotTable)
            Dim wizaHauckEUR = pt.CalculatedData.SelectField("CompanyName", "Wiza-Hauck").SelectField("Currency", "EUR").GetValue("Order value")

            'SelectField("Name", "Kianna Bradtke").
            Dim wizaHauckEURPercentOfTotal = pt.CalculatedData.SelectField("CompanyName", "Wiza-Hauck").SelectField("Currency", "EUR").GetValue("Order value % of total")

            Dim wizaHauckEURCountDifferance = pt.CalculatedData.SelectField("CompanyName", "Wiza-Hauck").SelectField("Currency", "EUR").GetValue("Count Difference From Previous")

            Console.WriteLine($"Show as: The value of field Order value for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEUR:N0}.")
            Console.WriteLine($"Show as: The value of field Order value % of total for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEURPercentOfTotal:P1}.")
            Console.WriteLine($"Show as: The value of field Order value From Previous for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEURCountDifferance:N0}.")
            Console.WriteLine()
        End Sub
        ''' <summary>
        ''' This sample shows how to use the ExcelPivotTable.GetPivotData method as an option to use the ExcelPivotTable.CalculatedData property.
        ''' </summary>
        ''' <paramname="ws">The worksheet containing the pivot tables</param>
        Private Sub GetPivotDataMethodSample(ws As ExcelWorksheet)
            Dim pt1 = ws.PivotTables(0)
            Dim pt2 = ws.PivotTables(1)
            Dim pt3 = ws.PivotTables(2)

            Dim grandTotal = pt1.GetPivotData("OrderValue")
            Dim tajikistanTotal = pt2.GetPivotData("OrderValue",
                New List(Of PivotDataFieldItemSelection)() From {
                New PivotDataFieldItemSelection("Country", "Tajikistan")
            })
            Dim equatorialGuinea = pt3.GetPivotData("OrderValue", New List(Of PivotDataFieldItemSelection)() From {
                New PivotDataFieldItemSelection("Country", "Equatorial Guinea")
            })

            Dim equatorialGuineaChelseyPowlowski = pt3.GetPivotData("OrderValue", New List(Of PivotDataFieldItemSelection)() From {New PivotDataFieldItemSelection("Country", "Equatorial Guinea"), New PivotDataFieldItemSelection("Name", "Chelsey Powlowski")})

            Console.WriteLine($"GetPivotData method: The grand total for pivot table {pt1.Name} is: {grandTotal:N0}.")
            Console.WriteLine($"GetPivotData method: The value of field OrderValue for Tajikistan for pivot table {pt2.Name} is: {tajikistanTotal:N0}.")
            Console.WriteLine($"GetPivotData method: The value of field OrderValue for Equatorial Guinea, Chelsey Powlowski for pivot table {pt3.Name} is: {equatorialGuineaChelseyPowlowski:N0}.")
            Console.WriteLine()
        End Sub
    End Module
End Namespace
