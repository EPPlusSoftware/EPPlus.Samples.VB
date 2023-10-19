﻿Imports OfficeOpenXml
Imports OfficeOpenXml.ConditionalFormatting
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class RemovalAndCasting
        Public Shared Sub Run(ByVal package As ExcelPackage)
            Dim worksheet = package.Workbook.Worksheets.Add("RemovalAndCasting")

            worksheet.Cells("A1:A5").Formula = "ROW()"

            Dim signs = worksheet.ConditionalFormatting.AddThreeIconSet("A1:A5", eExcelconditionalFormatting3IconsSetType.Signs)
            worksheet.ConditionalFormatting.AddDatabar("A1:A5", Color.Green)

            ' -----------------------------------------------------------
            ' Removing Conditional Formatting rules
            ' -----------------------------------------------------------
            ' Remove one Rule by its object
            'worksheet.ConditionalFormatting.Remove(signs);

            ' Remove one Rule by index
            'worksheet.ConditionalFormatting.RemoveAt(1);

            ' Remove one Rule by its Priority
            'worksheet.ConditionalFormatting.RemoveByPriority(2);

            ' Remove all the Rules
            'worksheet.ConditionalFormatting.RemoveAll();

            ' set some document properties
            package.Workbook.Properties.Title = "Conditional Formatting"
            package.Workbook.Properties.Author = "Eyal Seagull"
            package.Workbook.Properties.Comments = "This sample demonstrates how to add Conditional Formatting to an Excel 2007 worksheet using EPPlus"

            ' set some custom property values
            package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Eyal Seagull")
            package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus")

            'Getting a rule from the collection as a typed rule
            If worksheet.ConditionalFormatting(0).Type = eExcelConditionalFormattingRuleType.ThreeIconSet Then
                Dim iconRule = worksheet.ConditionalFormatting(0).As.ThreeIconSet 'Type cast the rule as an iconset rule.    
                'Do something with the iconRule...
            End If
            If worksheet.ConditionalFormatting(1).Type = eExcelConditionalFormattingRuleType.DataBar Then

                Dim dataBarRule = worksheet.ConditionalFormatting(1).As.DataBar 'Type cast the rule as an iconset rule.
                'Do something with the databarRule...
            End If
        End Sub
    End Class
End Namespace
