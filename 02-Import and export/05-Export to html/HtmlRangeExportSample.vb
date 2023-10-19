Imports OfficeOpenXml
Imports OfficeOpenXml.Export.HtmlExport
Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks

Namespace EPPlusSamples
    Public Module HtmlRangeExportSample
        'This sample demonstrates how to copy entire worksheet, ranges and how to exclude different cell properties.
        'More advanced samples using charts and json exports are available in our samples web site available 
        'here: https://samples.epplussoftware.com/HtmlExport, https://samples.epplussoftware.com/JsonExport
        Public Async Function RunAsync() As Task
            Dim outputFolder = FileUtil.GetDirectoryInfo("HtmlOutput")

            Await ExportGettingStartedAsync(outputFolder)

            ExportSalesReport(outputFolder)

            Await ExcludeCssAsync(outputFolder)

            ExportMultipleRanges(outputFolder)
        End Function

        Private Async Function ExportGettingStartedAsync(ByVal outputFolder As DirectoryInfo) As Task
            'Start by using the excel file generated in sample 8
            Using p = New ExcelPackage(FileUtil.GetFileInfo("1.01-GettingStarted.xlsx"))
                Dim ws = p.Workbook.Worksheets("Inventory")
                'Will create the html exporter for min and max bounds of the worksheet (ws.Dimensions)
                Dim exporter = ws.Cells.CreateHtmlExporter()

                'Get the html and styles in one call. 
                Dim html = Await exporter.GetSinglePageAsync()
                Await File.WriteAllTextAsync(FileUtil.GetFileInfo(outputFolder, "2.5-Range-GettingStarted.html", True).FullName, html)
            End Using
        End Function

        Private Sub ExportSalesReport(ByVal outputFolder As DirectoryInfo)
            'Start by using the excel file generated in sample 8
            Using p = New ExcelPackage(FileUtil.GetFileInfo("1.06-Salesreport.xlsx"))
                Dim ws = p.Workbook.Worksheets("Sales")
                Dim exporter = ws.Cells.CreateHtmlExporter()   'Will create the html exporter for min and max bounds of the worksheet (ws.Dimensions)
                exporter.Settings.HeaderRows = 4               'We have three header rows.
                exporter.Settings.TableId = "my-table"         'We can set an id of the worksheet if we want to use it in css or javascript.

                'By default EPPlus include the normal font in the css for the table. This can be tuned off and replaces by your own settings.
                exporter.Settings.Css.IncludeNormalFont = False
                'AdditionalCssElements is a collection where you can add your own styles for the table. You can also clear default styles set by EPPlus.
                exporter.Settings.Css.AdditionalCssElements.Add("font-family", "verdana")

                'EPPlus will not set column width and row heights by default, as this doesn't go well with todays responsive designs.
                'If you want fixed widths/heights set the proprties below to true...
                'Note that individual width and height are set direcly on the colspan-elements and tr-elements.
                'Default width and heights are set via the classes epp-dcw and epp-drh (with default StyleClassPrefix.).
                exporter.Settings.SetColumnWidth = True
                exporter.Settings.SetRowHeight = True

                'Get the html...
                Dim htmlTable = exporter.GetHtmlString()
                '...and the styles
                Dim cssTable = exporter.GetCssString()

                'EPPlus will not add the Excel grid lines, but you can easily add your own in the css...
                cssTable += "#my-table th,td {border:solid thin lightgray}"

                Dim html = $"<html><head><style type=""text/css"">{cssTable}</style></head>{htmlTable}</html>"

                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "2.5-Range-Salesreport.html", True).FullName, html)
            End Using
        End Sub
        Private Async Function ExcludeCssAsync(ByVal outputFolder As DirectoryInfo) As Task
            'Start by using the excel file generated in sample 20
            Using p = New ExcelPackage(FileUtil.GetFileInfo("Workbooks", "2.5-CreateAFileSystemReport.xlsx"))
                Dim ws = p.Workbook.Worksheets(0)
                Dim range = ws.Cells(1, 1, 5, ws.Dimension.End.Column)

                Dim exporter = range.CreateHtmlExporter()
                'Css can be excluded on style level, if you don't want some style or you want to add your own.
                exporter.Settings.Css.CssExclude.Font = eFontExclude.Bold Or eFontExclude.Italic Or eFontExclude.Underline

                Dim html = Await exporter.GetSinglePageAsync()
                Await File.WriteAllTextAsync(FileUtil.GetFileInfo(outputFolder, "2.5-Range-ExcludeCss.html", True).FullName, html)
            End Using
        End Function
        Private Sub ExportMultipleRanges(ByVal outputFolder As DirectoryInfo)
            'Start by using the excel file generated in sample 15
            Using p = New ExcelPackage(FileUtil.GetFileInfo("Workbooks", "2.5-ChartsAndThemes.xlsx"))
                'Now we will use the sample 15 and read two ranges from two different worksheets and combine them to use the same CSS.
                'To do so we create an HTML exporter on the workbook level and adds the ranges we want to use.
                Dim ws3D = p.Workbook.Worksheets("3D Charts")
                Dim wsStock = p.Workbook.Worksheets("Stock Chart")

                'We mark the top and bottom two values with red and green.
                ws3D.Cells("B13,B7").Style.Fill.SetBackground(Color.Green)
                ws3D.Cells("B14,B5").Style.Fill.SetBackground(Color.Red)

                'We mark the top and bottom two rows with red and green.
                wsStock.Cells("A3:E4").Style.Fill.SetBackground(Color.Green)
                wsStock.Cells("A7:E8").Style.Fill.SetBackground(Color.Red)

                Dim tblChartData = p.Workbook.Worksheets("ChartData").Tables(0)

                'Create the exporter. The workbook exporter exports ranges and tables, if the range corresponds to the table range.. 
                Dim rngExporter = p.Workbook.CreateHtmlExporter(ws3D.Cells("A1:D16"), wsStock.Cells("A1:E11"), tblChartData.Range)    'If you want to export a table, the exact table range must be used.

                'Get the html for the ranges in the HTML. The argument index referece to the ranges supplied when creating the exporter. 
                Dim html3D = rngExporter.GetHtmlString(0)
                Dim htmlStock = rngExporter.GetHtmlString(1)
                Dim tblHtml = rngExporter.GetHtmlString(2)
                Dim css = rngExporter.GetCssString()

                'We also exports a table and merge the css the range css.
                Dim htmlTemplate = "<html>" & vbCrLf & "<head>" & vbCrLf & "<style type=""text/css"">" & vbCrLf & "{0}</style></head>" & vbCrLf & "<body>" & vbCrLf & "{1}<hr>{2}<hr>{3}</body>" & vbCrLf & "</html>"
                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "2.5-Range-MultipleRanges.html", True).FullName, String.Format(htmlTemplate, css, html3D, htmlStock, tblHtml))
            End Using
        End Sub
    End Module
End Namespace
