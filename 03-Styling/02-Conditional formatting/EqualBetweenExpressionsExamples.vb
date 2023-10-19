Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Namespace EPPlusSamples.ConditionalFormatting
    Friend Class EqualBetweenExpressionsExamples
        Public Shared Sub Run(ByVal package As ExcelPackage)
            Dim sheet = package.Workbook.Worksheets.Add("ExpressionExamples")

            Dim range = "B1:B30"

            sheet.Cells("B1:B2").Value = 3
            sheet.Cells("B3").Value = 6
            sheet.Cells("B4").Value = 35
            sheet.Cells("B5").Value = 36
            sheet.Cells("B6").Value = 37
            sheet.Cells("B7").Value = 38
            sheet.Cells("B8").Value = 68
            sheet.Cells("B8").Value = 444
            sheet.Cells("B9").Value = 1000
            sheet.Cells("B10").Value = 25
            sheet.Cells("B11").Value = 43
            sheet.Cells("B12").Value = 43
            sheet.Cells("B13").Value = 43

            ' -------------------------------------------------------------------
            ' Create a Between rule
            ' -------------------------------------------------------------------
            Dim betweenRule = sheet.ConditionalFormatting.AddBetween(range)

            betweenRule.Formula = "IF(A1>5,10,20)"
            betweenRule.Formula2 = "IF(A1>5,30,50)"

            betweenRule.Style.Border.Right.Style = ExcelBorderStyle.Thick
            betweenRule.Style.Border.Right.Color.Color = Color.Goldenrod

            ' -------------------------------------------------------------------
            ' Create an Equal rule
            ' -------------------------------------------------------------------
            Dim equal = sheet.ConditionalFormatting.AddEqual(range)

            equal.Formula = "6"

            equal.Style.Border.Left.Style = ExcelBorderStyle.MediumDashed
            equal.Style.Border.Left.Color.Color = Color.Purple

            ' -------------------------------------------------------------------
            ' Create an NotEqual rule
            ' -------------------------------------------------------------------
            Dim notEqual = sheet.ConditionalFormatting.AddNotEqual("A10:A11")

            notEqual.Formula = "14"

            notEqual.Style.Border.BorderAround(ExcelBorderStyle.DashDotDot, Color.Firebrick)

            sheet.Cells("A10").Value = 14
            sheet.Cells("A11").Value = 10

            ' -------------------------------------------------------------------
            ' Create an Expression rule
            ' -------------------------------------------------------------------
            Dim customExpression = sheet.ConditionalFormatting.AddExpression(range)

            customExpression.Formula = "B1=B2"
            customExpression.Style.Font.Bold = True

            ' -------------------------------------------------------------------
            ' Create a GreaterThan rule
            ' -------------------------------------------------------------------
            Dim greater = sheet.ConditionalFormatting.AddGreaterThan(range)

            greater.Formula = "SE(B1<10,10,65)"

            greater.Style.Fill.PatternType = ExcelFillStyle.Solid
            greater.Style.Fill.BackgroundColor.Color = Color.DarkOrchid

            ' -------------------------------------------------------------------
            ' Create a GreaterThanOrEqual rule
            ' -------------------------------------------------------------------
            Dim greaterEqual = sheet.ConditionalFormatting.AddGreaterThanOrEqual(range)

            greaterEqual.Formula = "40"

            greaterEqual.Priority = 1
            greaterEqual.Style.Border.BorderAround(ExcelBorderStyle.Double, Color.Red)

            ' -------------------------------------------------------------------
            ' Create a LessThan rule
            ' -------------------------------------------------------------------
            Dim lessThan = sheet.ConditionalFormatting.AddLessThan(range)

            lessThan.Formula = "36"
            lessThan.Style.Font.Strike = True

            ' -------------------------------------------------------------------
            ' Create a LessThanOrEqual rule
            ' -------------------------------------------------------------------
            Dim lessThanEqual = sheet.ConditionalFormatting.AddLessThanOrEqual(range)

            lessThanEqual.Formula = "37"
            lessThanEqual.Style.Font.Italic = True

            ' -------------------------------------------------------------------
            ' Create a NotBetween rule
            ' -------------------------------------------------------------------
            Dim notBetween = sheet.ConditionalFormatting.AddNotBetween(range)

            notBetween.Style.Font.Color.Color = Color.ForestGreen

            notBetween.Formula = "333"
            notBetween.Formula2 = "999"
        End Sub
    End Class
End Namespace
