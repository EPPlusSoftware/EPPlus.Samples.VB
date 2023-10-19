Imports OfficeOpenXml
Imports OfficeOpenXml.Export.HtmlExport
Imports System
Imports System.IO

Namespace EPPlusSamples
    Public Module HtmlTableExportSample
        'This sample demonstrates how to export html from a table.
        'More advanced samples using charts and json exports are available in our samples web site available 
        'here: https://samples.epplussoftware.com/HtmlExport, https://samples.epplussoftware.com/JsonExport
        Public Sub Run()
            Console.WriteLine("Running sample 2.5 - Html export-tables")
            Dim outputFolder = FileUtil.GetDirectoryInfo("HtmlOutput")
            'Start by using the excel file generated in sample 28
            Using p = New ExcelPackage(FileUtil.GetFileInfo("Workbooks", "Tables.xlsx"))
                Dim wsSimpleTable = p.Workbook.Worksheets("SimpleTable")

                ExportSimpleTable1(outputFolder, wsSimpleTable)
                ExportSimpleTable2(outputFolder, wsSimpleTable)

                Dim wsStyleTables = p.Workbook.Worksheets("StyleTables")
                ExportStyleTables(outputFolder, wsStyleTables)

                'This samples exports the filtered table from the slicer sample.
                Dim wsSlicer = p.Workbook.Worksheets("Slicer")
                ExportSlicerTables1(outputFolder, wsSlicer)

                'Exports three tables and combine the html and css 
                ExportMultipleTables(outputFolder)
            End Using
        End Sub
        Private Sub ExportSimpleTable1(ByVal outputFolder As DirectoryInfo, ByVal wsSimpleTable As ExcelWorksheet)
            Dim table1 = wsSimpleTable.Tables(0)
            'Create the exporter for the table.
            Dim htmlExporter = table1.CreateHtmlExporter()

            'EPPlus will minify the css and html by default, but for this sample we want it easier to read.
            htmlExporter.Settings.Minify = False

            ' The GetSinglePage method generates en single page. You can also add a string parameter with your own HTML where where the styles and table html is inserted.
            Dim fullHtml = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1_SinglePage.html", True).FullName, fullHtml)

            'In most cases you want to keep the html and the styles separated, so you will retrive the html and the css in separate calls...
            Dim tableHtml = htmlExporter.GetHtmlString()
            Dim tableCss = htmlExporter.GetCssString()

            'First create the html file and reference the the css.
            Dim html = $"<html><head><link rel=""stylesheet"" href=""Table-01-Table1.css""</head>{tableHtml}</html>"
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1.html", True).FullName, html)

            'The css is written to a separate file.
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1.css", True).FullName, tableCss)
        End Sub
        Private Sub ExportSimpleTable2(ByVal outputFolder As DirectoryInfo, ByVal wsSimpleTable As ExcelWorksheet)
            Dim table2 = wsSimpleTable.Tables(1)

            'Create the exporter for the table.
            Dim htmlExporter = table2.CreateHtmlExporter()
            'EPPlus will generate Accessibility and data attributes by default, but you can turn it of in the settings.
            htmlExporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = False
            htmlExporter.Settings.RenderDataAttributes = False

            Dim html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Table2.html", True).FullName, html)

            'We can also change the table style to get a different styling.
            'Here we change to Medium15...
            table2.TableStyle = Table.TableStyles.Medium15
            html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-table2_Medium15.html", True).FullName, html)

            '...Here we use Dark2...
            table2.TableStyle = Table.TableStyles.Dark2
            html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-table2_Dark2.html", True).FullName, html)
        End Sub

        Private Sub ExportStyleTables(ByVal outputFolder As DirectoryInfo, ByVal wsStyleTables As ExcelWorksheet)
            'The last row of the cell contains uncalculated cell (they calculate when opened in Excel),
            'but in EPPlus we need to calculate them first to get a result in cell A254 in the totals row.
            wsStyleTables.Calculate()

            Dim table1 = wsStyleTables.Tables(0)
            Dim htmlExporter = table1.CreateHtmlExporter()

            'This sample exports the table as well as some individually cell styles. The headers have font italic and the totals row has a custom formatted text.
            'Also note that Column 2 has hyper links create for the mail addresses.
            Dim html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Styling_table1_with_hyperlinks.html", True).FullName, html)

            Dim table2 = wsStyleTables.Tables(1)
            htmlExporter = table2.CreateHtmlExporter()

            'Table 2 contains a custom table style.
            html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Styling_table2.html", True).FullName, html)
        End Sub
        Private Sub ExportSlicerTables1(ByVal outputFolder As DirectoryInfo, ByVal wsSlicer As ExcelWorksheet)
            Dim table1 = wsSlicer.Tables(0)
            Dim htmlExporter = table1.CreateHtmlExporter()

            'This sample exports the table filtered by the selection in the slicer (that applies the filter on the table).
            'By default EPPlus will remove hidden rows.
            Dim html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-03-Slicer.html", True).FullName, html)

            'You can change this option by setting eHiddenState.Include in the settings.
            'You can also set the it to eHiddenState.IncludeButHide if you want to apply your own filtering.
            htmlExporter.Settings.HiddenRows = eHiddenState.Include
            html = htmlExporter.GetSinglePage()
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-03-Slicer_table_all_rows.html", True).FullName, html)
        End Sub
        Private Sub ExportMultipleTables(ByVal outputFolder As DirectoryInfo)
            Using p = New ExcelPackage(FileUtil.GetFileInfo("Workbooks", "2.5-LoadingData.xlsx"))
                'Now we will use the third worksheet from sample 4, that contains three tables with different styles.
                Dim wsList = p.Workbook.Worksheets("FromList")

                Dim tbl1 = wsList.Tables(0)
                Dim exporter1 = tbl1.CreateHtmlExporter()
                Dim tbl1Html = exporter1.GetHtmlString()
                Dim tbl1Css = exporter1.GetCssString()

                Dim tbl2 = wsList.Tables(1)
                Dim exporter2 = tbl2.CreateHtmlExporter()
                Dim tbl2Html = exporter2.GetHtmlString()
                'We have already exported the css once, so we don't want shared css classes to be added again.
                exporter2.Settings.Css.IncludeSharedClasses = False
                Dim tbl2Css = exporter2.GetCssString()

                Dim tbl3 = wsList.Tables(2)
                Dim exporter3 = tbl3.CreateHtmlExporter()
                Dim tbl3Html = exporter3.GetHtmlString()
                exporter3.Settings.Css.IncludeSharedClasses = False

                Dim tbl3Css = exporter3.GetCssString()

                'As the tables have different table styles we add all of the css's.
                'If multiple tables have the same table style, you should only add one of them.
                Dim css = tbl1Css & tbl2Css & tbl3Css

                Dim htmlTemplate = "<html>" & vbCrLf & "<head>" & vbCrLf & "<style type=""text/css"">" & vbCrLf & "{0}</style></head>" & vbCrLf & "<body>" & vbCrLf & "{1}<hr>{2}<hr>{3}</body>" & vbCrLf & "</html>"
                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-04-MultipleTables.html", True).FullName, String.Format(htmlTemplate, css, tbl1Html, tbl2Html, tbl3Html))
            End Using
        End Sub

    End Module
End Namespace
