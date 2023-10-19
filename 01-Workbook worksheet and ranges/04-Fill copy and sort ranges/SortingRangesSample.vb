Imports OfficeOpenXml
Imports System

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    ''' <summary>
    ''' This sample demonstrates how to sort ranges with EPPlus.
    ''' </summary>
    Public Module SortingRangesSample
        Private _letters As String() = New String() {"A", "B", "C", "D"}
        Private _tShirtSizes As String() = New String() {"S", "M", "L", "XL", "XXL"}
        Private Const StartRow As Integer = 3
        Public Sub Run()
            Using p = New ExcelPackage()
                CreateWorksheetsAndLoadData(p)
                Dim sheet1 = p.Workbook.Worksheets(0)
                Dim sheet2 = p.Workbook.Worksheets(1)

                ' Sort the range by column 0, then by column 1 descending
                sheet1.Cells("A3:D17").Sort(Sub(x) x.SortBy.Column(0).ThenSortBy.Column(1, eSortOrder.Descending))

                ' Sort the range left to right by row 0 (using a custom list), then by row 1
                sheet2.Cells("A3:K5").Sort(Sub(x) x.SortLeftToRightBy.Row(0).UsingCustomList("S", "M", "L", "XL", "XXL").ThenSortBy.Row(1))

                p.SaveAs(FileUtil.GetCleanFileInfo("1.04-SortingRanges.xlsx"))
            End Using
        End Sub

        Private Sub CreateWorksheetsAndLoadData(ByVal p As ExcelPackage)
            Dim rnd = New Random(CInt(Date.UtcNow.ToOADate()))

            Dim sheet1 = p.Workbook.Worksheets.Add("Sort top down")
            sheet1.Cells("A1").Value = "To view the sort state in Excel 2019 with english localization, select the range A3:D17, right click and chose 'Sort' followed by 'Custom sort'"
            ' create random data for this sheet
            For row = StartRow To StartRow + 15 - 1
                For col = 1 To 4
                    If col = 1 Then
                        Dim ix = rnd.Next(0, _letters.Length - 1)
                        sheet1.Cells(row, 1).Value = _letters(ix)
                    ElseIf col = 4 Then
                        ' Add a formula in the right most column to demonstrate that the formulas will be shifted when sorted.
                        sheet1.Cells(row, 4).Formula = $"SUM(B{row}:C{row})"
                    Else
                        sheet1.Cells(row, col).Value = rnd.Next(14, 555)
                    End If
                Next
            Next

            Dim sheet2 = p.Workbook.Worksheets.Add("Sort left to right")
            sheet2.Cells("A1").Value = "To view the sort state in Excel 2019 with english localization, select the range A3:K5, right click and chose 'Sort' followed by 'Custom sort'"
            ' create random data for this sheet
            For col = 1 To 11
                For row = 3 To 5
                    If row = 3 Then
                        Dim ix = rnd.Next(0, _tShirtSizes.Length - 1)
                        sheet2.Cells(row, col).Value = _tShirtSizes(ix)
                        sheet2.Cells(row, col).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    Else
                        sheet2.Cells(row, col).Value = rnd.Next(14, 555)
                    End If
                Next
            Next
        End Sub
    End Module
End Namespace
