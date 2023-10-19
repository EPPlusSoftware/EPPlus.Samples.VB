Imports OfficeOpenXml
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class StartsEndContainsExample
        Public Shared Sub Run(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("StartEndContains")

            sheet.Cells("E11").Value = "SearchMe will find me but the notContains will not. Nope NotMe never."
            sheet.Cells("E12").Value = "SearchMe will find me but not EndText because I'm a EndTextFaker"
            sheet.Cells("E13").Value = "NotMe won't be found by much"
            sheet.Cells("E14").Value = "This will be found by notContains and is also EndText"
            sheet.Cells("E15").Value = "This will be found by notContains and is also EndText"
            sheet.Cells("E16").Value = "SearchMe To be found by all and let the end be EndText"

            ' -------------------------------------------------------------------
            ' Create a BeginsWith rule
            ' -------------------------------------------------------------------
            Dim cellIsAddress As ExcelAddress = New ExcelAddress("E11:E20")
            Dim beginsWith = sheet.ConditionalFormatting.AddBeginsWith(cellIsAddress)

            beginsWith.Text = "SearchMe"

            beginsWith.Style.Font.Bold = True

            ' -------------------------------------------------------------------
            ' Create an EndsWith rule
            ' -------------------------------------------------------------------
            Dim EndText = sheet.ConditionalFormatting.AddEndsWith(cellIsAddress)

            EndText.Text = "EndText"

            EndText.Style.Font.Color.Color = Color.DarkRed

            ' -------------------------------------------------------------------
            ' Create a ContainsText rule
            ' -------------------------------------------------------------------
            Dim ContainsText = sheet.ConditionalFormatting.AddContainsText(cellIsAddress)

            ContainsText.Text = "Me"
            ContainsText.Style.Font.Italic = True

            ' -------------------------------------------------------------------
            ' Create a NotContainsText rule
            ' -------------------------------------------------------------------
            Dim notContains = sheet.ConditionalFormatting.AddNotContainsText(cellIsAddress)

            notContains.Text = "NotMe"
            notContains.Style.Border.BorderAround(Style.ExcelBorderStyle.MediumDashed, Color.Red)
        End Sub
    End Class
End Namespace
