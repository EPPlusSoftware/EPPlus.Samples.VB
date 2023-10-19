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
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.DataValidation
Imports OfficeOpenXml.DataValidation.Contracts

Namespace EPPlusSamples.FiltersAndValidations
    ''' <summary>
    ''' This sample shows how to use data validation
    ''' </summary>
    Friend Class DataValidationSample
        Public Shared Sub Run()
            Console.WriteLine("Running sample 4.1")
            Dim output = FileUtil.GetCleanFileInfo("4.1-DataValidation.xlsx")
            Using package = New ExcelPackage(output)
                AddIntegerValidation(package)
                AddListValidationFormula(package)
                AddListValidationValues(package)
                AddTimeValidation(package)
                AddDateTimeValidation(package)
                ReadExistingValidationsFromPackage(package)
                package.SaveAs(output)
            End Using
            Console.WriteLine("Sample 4.1 created {0}", output.FullName)
            Console.WriteLine()
        End Sub

        ''' <summary>
        ''' Adds integer validation
        ''' </summary>
        ''' <paramname="file"></param>
        Private Shared Sub AddIntegerValidation(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("integer")
            ' add a validation and set values
            Dim validation = sheet.DataValidations.AddIntegerValidation("A1:A2")
            ' Alternatively:
            'var validation = sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop
            validation.PromptTitle = "Enter a integer value here"
            validation.Prompt = "Value should be between 1 and 5"
            validation.ShowInputMessage = True
            validation.ErrorTitle = "An invalid value was entered"
            validation.Error = "Value must be between 1 and 5"
            validation.ShowErrorMessage = True
            validation.Operator = ExcelDataValidationOperator.between
            validation.Formula.Value = 1
            validation.Formula2.Value = 5

            Console.WriteLine("Added sheet for integer validation")
        End Sub

        ''' <summary>
        ''' Adds a list validation where the list source is a formula
        ''' </summary>
        ''' <paramname="package"></param>
        Private Shared Sub AddListValidationFormula(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("list formula")
            sheet.Cells("B1").Style.Font.Bold = True
            sheet.Cells("B1").Value = "Source values"
            sheet.Cells("B2").Value = 1
            sheet.Cells("B3").Value = 2
            sheet.Cells("B4").Value = 3

            ' add a validation and set values
            Dim validation = sheet.DataValidations.AddListValidation("A1")
            ' Alternatively:
            ' var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
            validation.ShowErrorMessage = True
            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning
            validation.ErrorTitle = "An invalid value was entered"
            validation.Error = "Select a value from the list"
            validation.Formula.ExcelFormula = "B2:B4"

            Console.WriteLine("Added sheet for list validation with formula")

        End Sub

        ''' <summary>
        ''' Adds a list validation where the selectable values are set
        ''' </summary>
        ''' <paramname="package"></param>
        Private Shared Sub AddListValidationValues(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("list values")

            ' add a validation and set values
            Dim validation = sheet.DataValidations.AddListValidation("A1")
            validation.ShowErrorMessage = True
            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning
            validation.ErrorTitle = "An invalid value was entered"
            validation.Error = "Select a value from the list"
            For i = 1 To 5
                validation.Formula.Values.Add(i.ToString())
            Next
            Console.WriteLine("Added sheet for list validation with values")

        End Sub

        ''' <summary>
        ''' Adds a time validation
        ''' </summary>
        ''' <paramname="package"></param>
        Private Shared Sub AddTimeValidation(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("time")
            ' add a validation and set values
            Dim validation = sheet.DataValidations.AddTimeValidation("A1")
            ' Alternatively:
            ' var validation = sheet.Cells["A1"].DataValidation.AddTimeDataValidation();
            validation.ShowErrorMessage = True
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop
            validation.ShowInputMessage = True
            validation.PromptTitle = "Enter time in format HH:MM:SS"
            validation.Prompt = "Should be greater than 13:30:10"
            validation.Operator = ExcelDataValidationOperator.greaterThan
            Dim time = validation.Formula.Value
            time.Hour = 13
            time.Minute = 30
            time.Second = 10
            Console.WriteLine("Added sheet for time validation")
        End Sub

        Private Shared Sub AddDateTimeValidation(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("datetime")
            ' add a validation and set values
            Dim validation = sheet.DataValidations.AddDateTimeValidation("A1")
            ' Alternatively:
            ' var validation = sheet.Cells["A1"].DataValidation.AddDateTimeDataValidation();
            validation.ShowErrorMessage = True
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop
            validation.Error = "Invalid date!"
            validation.ShowInputMessage = True
            validation.Prompt = "Enter a date greater than todays date here"
            validation.Operator = ExcelDataValidationOperator.greaterThan
            validation.Formula.Value = Date.Now.Date
            Console.WriteLine("Added sheet for date time validation")

        End Sub

        ''' <summary>
        ''' shows details about all existing validations in the entire workbook
        ''' </summary>
        ''' <paramname="package"></param>
        Private Shared Sub ReadExistingValidationsFromPackage(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("Package validations")
            ' print headers
            sheet.Cells("A1:E1").Style.Font.Bold = True
            sheet.Cells("A1").Value = "Type"
            sheet.Cells("B1").Value = "Address"
            sheet.Cells("C1").Value = "Operator"
            sheet.Cells("D1").Value = "Formula1"
            sheet.Cells("E1").Value = "Formula2"

            Dim row = 2
            For Each otherSheet In package.Workbook.Worksheets
                If otherSheet Is sheet Then
                    Continue For
                End If
                For Each dv In otherSheet.DataValidations
                    sheet.Cells("A" & row.ToString()).Value = dv.ValidationType.Type.ToString()
                    sheet.Cells("B" & row.ToString()).Value = dv.Address.Address
                    If dv.AllowsOperator Then
                        sheet.Cells("C" & row.ToString()).Value = CType(dv, IExcelDataValidationWithOperator).Operator.ToString()
                    End If
                    ' type casting is needed to get validationtype-specific values
                    Select Case dv.ValidationType.Type
                        Case eDataValidationType.Whole
                            DataValidationSample.PrintWholeValidationDetails(sheet, dv.As.IntegerValidation, row)
                        Case eDataValidationType.List
                            DataValidationSample.PrintListValidationDetails(sheet, dv.As.ListValidation, row)
                        Case eDataValidationType.Time
                            ' the rest of the types are not supported in this sample, but I hope you get the picture...
                            DataValidationSample.PrintTimeValidationDetails(sheet, dv.As.TimeValidation, row)
                        Case Else
                    End Select
                    row += 1
                Next
            Next
        End Sub

        Private Shared Sub PrintWholeValidationDetails(ByVal sheet As ExcelWorksheet, ByVal wholeValidation As IExcelDataValidationInt, ByVal row As Integer)
            sheet.Cells("D" & row.ToString()).Value = If(wholeValidation.Formula.Value.HasValue, wholeValidation.Formula.Value.Value.ToString(), wholeValidation.Formula.ExcelFormula)
            sheet.Cells("E" & row.ToString()).Value = If(wholeValidation.Formula2.Value.HasValue, wholeValidation.Formula2.Value.Value.ToString(), wholeValidation.Formula2.ExcelFormula)
        End Sub

        Private Shared Sub PrintListValidationDetails(ByVal sheet As ExcelWorksheet, ByVal listValidation As IExcelDataValidationList, ByVal row As Integer)
            Dim value = String.Empty
            ' if formula is used - show it...
            If Not String.IsNullOrEmpty(listValidation.Formula.ExcelFormula) Then
                value = listValidation.Formula.ExcelFormula
            Else
                ' otherwise - show the values from the list collection
                Dim sb = New StringBuilder()
                For Each listValue In listValidation.Formula.Values
                    If sb.Length > 0 Then
                        sb.Append(",")
                    End If
                    sb.Append(listValue)
                Next
                value = sb.ToString()
            End If
            sheet.Cells("D" & row.ToString()).Value = value
        End Sub

        Private Shared Sub PrintTimeValidationDetails(ByVal sheet As ExcelWorksheet, ByVal validation As IExcelDataValidationTime, ByVal row As Integer)
            Dim value1 = String.Empty
            If Not String.IsNullOrEmpty(validation.Formula.ExcelFormula) Then
                value1 = validation.Formula.ExcelFormula
            Else
                value1 = String.Format("{0}:{1}:{2}", validation.Formula.Value.Hour, validation.Formula.Value.Minute, If(validation.Formula.Value.Second, 0))
            End If
            sheet.Cells("D" & row.ToString()).Value = value1
        End Sub
    End Class
End Namespace
