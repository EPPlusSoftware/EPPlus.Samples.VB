Imports OfficeOpenXml
Imports OfficeOpenXml.ConditionalFormatting

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class IconsetsExample
        Public Shared Sub Run(ByVal pck As ExcelPackage)
            Dim sheet = pck.Workbook.Worksheets.Add("Iconsets")
            'Fill sheet with data similar to wiki example
            SheetInitalizeData(sheet)

            Dim cfRule = sheet.ConditionalFormatting.AddThreeIconSet(sheet.Cells("D7:D13"), eExcelconditionalFormatting3IconsSetType.TrafficLights2)
            'Normally order is red yellow green we want high values to be red so we reverse
            cfRule.Reverse = True
            cfRule.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula
            cfRule.Icon2.Formula = "$E$2 * ($E$4 +1)"
            cfRule.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula
            cfRule.Icon3.Formula = "$E$3"

            'Below only available epplus7 onwards
            'Switching out icons in a pre-existing iconset
            Dim customIcons = sheet.ConditionalFormatting.AddFiveIconSet(sheet.Cells("G1:G13"), eExcelconditionalFormatting5IconsSetType.Quarters)

            'Switch icons
            customIcons.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCross
            customIcons.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.HalfGoldStar
            customIcons.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDiamond

            'Add data to cells so icons show up.
            sheet.Cells("G1:G13").Formula = "ROW()"

            sheet.Cells.AutoFitColumns()
        End Sub

        Private Shared Sub SheetInitalizeData(ByVal sheet As ExcelWorksheet)
            sheet.Cells("D7:D13").Formula = "INT(RAND()*100)"

            sheet.Cells("B2").Value = "Speed Limit"
            sheet.Cells("B3").Value = "Drivers license suspended"
            sheet.Cells("B4").Value = "Tolerance"

            sheet.Cells("E2").Value = 50
            sheet.Cells("E3").Value = 80
            sheet.Cells("E4").Value = 0.08

            sheet.Cells("C6").Value = "Driver"
            sheet.Cells("D6").Value = "Speed"

            sheet.Cells("C7").Value = "Peter"
            sheet.Cells("C8").Value = "Maria"
            sheet.Cells("C9").Value = "John"
            sheet.Cells("C10").Value = "Bob"
            sheet.Cells("C11").Value = "Anna"
            sheet.Cells("C12").Value = "Cecilia"
            sheet.Cells("C13").Value = "Joe"
        End Sub
    End Class
End Namespace
