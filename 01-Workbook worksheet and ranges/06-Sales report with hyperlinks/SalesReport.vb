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
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports System.Data.SQLite
Imports System.Text

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module SalesReportFromDatabase
        ''' <summary>
        ''' Creates a workbook and populates using data from a SQLite database
        ''' </summary>
        ''' <paramname="outputDir">The output directory</param>
        ''' <paramname="templateDir">The location of the sample template</param>
        ''' <paramname="connectionString">The connection string to the SQLite database</param>
        Public Sub Run()
            Console.WriteLine("Running sample 1.6")

            Dim file = FileUtil.GetCleanFileInfo("1.06-Salesreport.xlsx")
            Using xlPackage As ExcelPackage = New ExcelPackage(file)
                Dim worksheet = xlPackage.Workbook.Worksheets.Add("Sales")
                Dim namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink")
                namedStyle.Style.Font.UnderLine = True
                namedStyle.Style.Font.Color.SetColor(Color.Blue)
                namedStyle.BuildInId = 8 'This is the id for the build in HyperLink style.

                Const startRow = 5
                Dim row = startRow
                'Create Headers and format them 
                worksheet.Cells("A1").Value = "Fiction Inc."
                Using r = worksheet.Cells("A1:G1")
                    r.Merge = True
                    r.Style.Font.SetFromFont("Britannic Bold", 22, False, True)
                    r.Style.Font.Color.SetColor(Color.White)
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93))
                End Using
                worksheet.Cells("A2").Value = "Sales Report"
                Using r = worksheet.Cells("A2:G2")
                    r.Merge = True
                    r.Style.Font.SetFromFont("Britannic Bold", 18, False, True)
                    r.Style.Font.Color.SetColor(Color.Black)
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228))
                End Using

                worksheet.Cells("A4").Value = "Company"
                worksheet.Cells("B4").Value = "Sales Person"
                worksheet.Cells("C4").Value = "Country"
                worksheet.Cells("D4").Value = "Order Id"
                worksheet.Cells("E4").Value = "OrderDate"
                worksheet.Cells("F4").Value = "Order Value"
                worksheet.Cells("G4").Value = "Currency"
                worksheet.Cells("A4:G4").Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells("A4:G4").Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228))
                worksheet.Cells("A4:G4").Style.Font.Bold = True


                'Lets connect to the sample database for some data
                Using sqlConn = New SQLiteConnection(ConnectionString)
                    sqlConn.Open()
                    Dim sql = GetSql()
                    Using sqlCmd = New SQLiteCommand(sql, sqlConn)
                        Using sqlReader = sqlCmd.ExecuteReader()
                            ' get the data and fill rows 5 onwards
                            While sqlReader.Read()
                                Dim col = 1
                                ' our query has the columns in the right order, so simply
                                ' iterate through the columns
                                For i = 0 To sqlReader.FieldCount - 1
                                    ' use the email address as a hyperlink for column 1
                                    If Equals(sqlReader.GetName(i), "email") Then
                                        ' insert the email address as a hyperlink for the name
                                        Dim hyperlink As String = "mailto:" & sqlReader.GetValue(i).ToString()
                                        worksheet.Cells(row, 2).Hyperlink = New Uri(hyperlink, UriKind.Absolute)
                                    Else
                                        ' do not bother filling cell with blank data (also useful if we have a formula in a cell)
                                        If sqlReader.GetValue(i) IsNot Nothing Then worksheet.Cells(row, col).Value = sqlReader.GetValue(i)
                                        col += 1
                                    End If
                                Next
                                row += 1
                            End While
                            sqlReader.Close()

                            worksheet.Cells(startRow, 2, row - 1, 2).StyleName = "HyperLink"
                            worksheet.Cells(startRow, 5, row - 1, 5).Style.Numberformat.Format = "yyyy/mm/dd"
                            worksheet.Cells(startRow, 6, row - 1, 6).Style.Numberformat.Format = "[$$-409]#,##0"

                            'Set column width
                            worksheet.Columns(1, 3).Width = 35
                            worksheet.Columns(2).Width = 28
                            worksheet.Columns(4).Width = 10
                            worksheet.Columns(5, 7).Width = 12
                        End Using
                    End Using
                    sqlConn.Close()

                    ' lets set the header text 
                    worksheet.HeaderFooter.OddHeader.CenteredText = "Fiction Inc. Sales Report"
                    ' add the page number to the footer plus the total number of pages
                    worksheet.HeaderFooter.OddFooter.RightAlignedText = String.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages)
                    ' add the sheet name to the footer
                    worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName
                    ' add the file path to the footer
                    worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath & ExcelHeaderFooter.FileName
                End Using

                ' We can also add some document properties to the spreadsheet 

                ' Set some core property values
                xlPackage.Workbook.Properties.Title = "Sales Report"
                xlPackage.Workbook.Properties.Author = "Jan Källman"
                xlPackage.Workbook.Properties.Subject = "Sales Report Samples"
                xlPackage.Workbook.Properties.Keywords = "Office Open XML"
                xlPackage.Workbook.Properties.Category = "Sales Report  Samples"
                xlPackage.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel file from scratch using EPPlus"

                ' Set some extended property values
                xlPackage.Workbook.Properties.Company = "Fiction Inc."
                xlPackage.Workbook.Properties.HyperlinkBase = New Uri("https://EPPlusSoftware.com")

                ' Set some custom property values
                xlPackage.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman")
                xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "1")
                xlPackage.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus")

                ' save the new spreadsheet
                xlPackage.Save()
            End Using

            Console.WriteLine("Sample 1.6 created: {0}", file.FullName)
            Console.WriteLine()
        End Sub

        Private Function GetSql() As String
            Dim sb = New StringBuilder()

            sb.Append("select cu.Name as CompanyName, sp.Name, Email, co.Name as Country, OrderId, orderdate, ordervalue, currency ")
            sb.Append("from Customer cu inner join ")
            sb.Append("Orders od on cu.CustomerId=od.CustomerId inner join ")
            sb.Append("SalesPerson sp on od.salesPersonId = sp.salesPersonId inner join ")
            sb.Append("City ci on ci.cityId = sp.cityId inner join ")
            sb.Append("Country co on ci.countryId = co.countryId ")
            sb.Append("ORDER BY 1,2 desc")

            Return sb.ToString()
        End Function
    End Module
End Namespace
