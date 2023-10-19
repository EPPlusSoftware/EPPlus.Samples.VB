Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.ConditionalFormatting

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class ColorScaleExample
        Public Shared Sub Run(ByVal pck As ExcelPackage)
            Dim sheet = pck.Workbook.Worksheets.Add("ColorScales")

            sheet.Cells("A1:B20").Formula = "ROW()"

            Dim twoScale = sheet.ConditionalFormatting.AddTwoColorScale("A1:A20")
            Dim threeScale = sheet.ConditionalFormatting.AddThreeColorScale("B1:B20")

            twoScale.LowValue.Color = Color.CadetBlue
            twoScale.HighValue.Color = ColorTranslator.FromHtml("#FF63BE7B")

            threeScale.LowValue.Color = Color.DarkRed
            threeScale.MiddleValue.Color = Color.Orange
            threeScale.HighValue.Color = Color.ForestGreen
            'ColorSettings attribute allow you to use tint
            threeScale.HighValue.ColorSettings.Tint = 0.80

            'It can also be used for alternative ways to set color. Note: Only last applied colorsetting matters.
            'Except for Tint which works with all.

            'threeScale.LowValue.ColorSettings.Theme = eThemeSchemeColor.Accent3;
            'threeScale.MiddleValue.ColorSettings.Auto = true;
            'threeScale.HighValue.ColorSettings.Index = 3;

            threeScale.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile
            threeScale.MiddleValue.Value = 50

            sheet.Cells.AutoFitColumns()
        End Sub
    End Class
End Namespace
