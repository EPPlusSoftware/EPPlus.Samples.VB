Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml.Table
Imports OfficeOpenXml.Table.PivotTable
Imports System
Imports System.Drawing

Namespace EPPlusSamples.TablesPivotTablesAndSlicers
    ''' <summary>
    ''' This sample demonstrates how to add custom named styles for 
    ''' </summary>
    Public Module CustomTableSlicerStyleSample
        Public Sub Run()
            Console.WriteLine("Running sample 7.4 - Custom table and slicer styles")
            Using p = New ExcelPackage()
                CreateTableStyles(p)
                CreatePivotTableStyles(p)
                CreateSlicerStyles(p)

                p.SaveAs(FileUtil.GetCleanFileInfo("7.4-TableAndSlicerStyles.xlsx"))
            End Using
            Console.WriteLine("Sample 7.4 finished.")
            Console.WriteLine()
        End Sub

        Private Sub CreateTableStyles(ByVal p As ExcelPackage)
            Dim wsTables = p.Workbook.Worksheets.Add("CustomStyledTables")

            'Create a custom table style from scratch and adds a fill gradient fill style 
            Dim customTableStyle1 = "CustomTableStyle1"
            CreateCustomTableStyleFromScratch(p, customTableStyle1)

            'This samples creates a style with the build in table style Dark11 as template and set the header row and table row font to italic.
            Dim customTableStyle2 = "CustomTableStyleFromDark11"
            CreateCustomTableStyleFromBuildInTableStyle(p, customTableStyle2)

            'This samples creates a style with the build in table style Medium11 as template and set the header row and table row font to italic.
            Dim customTableStyle3 = "CustomTableAndPivotTableStyleFromDark11"
            CreateCustomTableAndPivotTableStyleFromBuildInStyle(p, customTableStyle3)


            Dim tbl1 = CreateTable(wsTables, "Table1")
            tbl1.StyleName = customTableStyle1

            Dim tbl2 = CreateTable(wsTables, "Table2", 9, 1)
            tbl2.StyleName = customTableStyle2

            Dim tbl3 = CreateTable(wsTables, "Table3", 17, 1)
            tbl3.StyleName = customTableStyle3

            wsTables.Cells.AutoFitColumns()
        End Sub

        Private Sub CreatePivotTableStyles(ByVal p As ExcelPackage)
            Dim wsPivotTable = p.Workbook.Worksheets.Add("CustomStyledPivotTables")

            'Create a pivot table style from scratch.
            Dim customPivotTableStyle1 = "CustomPivotTableStyle1"
            CreateCustomPivotTableStyleFromScratch(p, customPivotTableStyle1)

            'This samples creates a style with the build in table style Dark11 as template and set the header row and table row font to italic.
            Dim customPivotTableStyle2 = "CustomPivotTableStyleFromMedium25"
            CreateCustomPivotTableStyleFromBuildInTableStyle(p, customPivotTableStyle2)

            'Create a pivot table and use the named style we created earlier in this sample for both pivot tables and tables.
            Dim pt1 = CreatePivotTable(wsPivotTable, "PivotTable1", p.Workbook.Worksheets(0).Tables(0), wsPivotTable.Cells("A3"))
            pt1.StyleName = "CustomTableAndPivotTableStyleFromDark11"

            Dim pt2 = CreatePivotTable(wsPivotTable, "PivotTable2", p.Workbook.Worksheets(0).Tables(0), wsPivotTable.Cells("A15"))
            pt2.StyleName = customPivotTableStyle1

            Dim pt3 = CreatePivotTable(wsPivotTable, "PivotTable3", p.Workbook.Worksheets(0).Tables(0), wsPivotTable.Cells("A30"))
            pt3.StyleName = customPivotTableStyle2

        End Sub

        Private Sub CreateSlicerStyles(ByVal p As ExcelPackage)
            Dim wsSlicers = p.Workbook.Worksheets.Add("CustomStyledSlicers")
            Dim tbl = CreateTable(wsSlicers, "TableForSlicer1")

            Dim slicer1 = tbl.Columns(0).AddSlicer()
            slicer1.SetPosition(100, 300)

            'Create a slicer style from scratch.
            Dim customSlicerStyle1 = "CustomSlicerStyleConsole"
            CreateCustomSlicerStyleFromScratch(p, customSlicerStyle1)
            slicer1.StyleName = customSlicerStyle1

            Dim slicer2 = tbl.Columns(1).AddSlicer()
            slicer2.SetPosition(100, 500)

            Dim customSlicerStyle2 = "CustomSlicerStyleFromStyleDark2"
            CreateCustomSlicerStyleFromBuildInStyle(p, customSlicerStyle2)
            slicer2.StyleName = customSlicerStyle2
        End Sub

        Private Sub CreateCustomSlicerStyleFromBuildInStyle(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a custom named slicer style that applies users the build in style Dark4 as template and make some minor modifications.
            Dim customSlicerStyle = p.Workbook.Styles.CreateSlicerStyle(styleName, eSlicerStyle.Dark4)
            customSlicerStyle.WholeTable.Style.Font.Name = "Broadway"
            customSlicerStyle.HeaderRow.Style.Font.Italic = True
            customSlicerStyle.HeaderRow.Style.Border.Bottom.Color.SetColor(Color.Red)
        End Sub

        Private Sub CreateCustomSlicerStyleFromScratch(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a named style that applies to slicers with a console feel to the style.
            Dim customSlicerStyle = p.Workbook.Styles.CreateSlicerStyle(styleName)

            customSlicerStyle.WholeTable.Style.Font.Name = "Consolas"
            customSlicerStyle.WholeTable.Style.Font.Size = 12
            customSlicerStyle.WholeTable.Style.Font.Color.SetColor(Color.WhiteSmoke)
            customSlicerStyle.WholeTable.Style.Fill.BackgroundColor.SetColor(Color.Black)

            customSlicerStyle.HeaderRow.Style.Fill.BackgroundColor.SetColor(Color.LightGray)
            customSlicerStyle.HeaderRow.Style.Font.Color.SetColor(Color.Black)

            customSlicerStyle.SelectedItemWithData.Style.Fill.BackgroundColor.SetColor(Color.Gray)
            customSlicerStyle.SelectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray)

            customSlicerStyle.SelectedItemWithNoData.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(&HFF, 64, 64, 64))
            customSlicerStyle.SelectedItemWithNoData.Style.Font.Color.SetColor(Color.DarkGray)
            customSlicerStyle.SelectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray)

            customSlicerStyle.UnselectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray)
            customSlicerStyle.UnselectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray)

            customSlicerStyle.UnselectedItemWithNoData.Style.Font.Color.SetColor(Color.DarkGray)

            customSlicerStyle.HoveredSelectedItemWithData.Style.Fill.BackgroundColor.SetColor(Color.DarkGray)
            customSlicerStyle.HoveredSelectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke)

            customSlicerStyle.HoveredSelectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke)

            customSlicerStyle.HoveredUnselectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke)
            customSlicerStyle.HoveredUnselectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke)
        End Sub

#Region "Table Styles"
        Private Sub CreateCustomTableStyleFromScratch(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a named style used to tables only.
            Dim customTableStyle = p.Workbook.Styles.CreateTableStyle(styleName)

            customTableStyle.WholeTable.Style.Font.Color.SetColor(eThemeSchemeColor.Text2)
            customTableStyle.HeaderRow.Style.Font.Bold = True
            customTableStyle.HeaderRow.Style.Font.Italic = True

            customTableStyle.HeaderRow.Style.Fill.Style = eDxfFillStyle.GradientFill
            customTableStyle.HeaderRow.Style.Fill.Gradient.Degree = 90

            Dim c1 = customTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(0)
            c1.Color.SetColor(Color.LightGreen)

            Dim c3 = customTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(100)
            c3.Color.SetColor(Color.DarkGreen)
        End Sub
        Private Sub CreateCustomTableStyleFromBuildInTableStyle(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a new custom table style with the build in style Dark11 as template.
            Dim customTableStyle = p.Workbook.Styles.CreateTableStyle(styleName, TableStyles.Dark11)

            customTableStyle.HeaderRow.Style.Font.Italic = True
            customTableStyle.TotalRow.Style.Font.Italic = True

            'Set the stripe size to 2 rows for both the the row stripes
            customTableStyle.FirstRowStripe.BandSize = 2
            customTableStyle.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGreen)

            customTableStyle.SecondRowStripe.BandSize = 2
            customTableStyle.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue)
        End Sub
        Private Sub CreateCustomTableAndPivotTableStyleFromBuildInStyle(ByVal p As ExcelPackage, ByVal customTableStyle3 As String)
            'Create a named style that can be used both for tables and pivot tables. 
            'We create this style from one of the build in pivot table styles - Medium13, but table styles can also be used as a parameter for this method
            Dim customTableStyle = p.Workbook.Styles.CreateTableAndPivotTableStyle(customTableStyle3, PivotTableStyles.Medium13)

            'Set the header row and total row border to dotted.
            customTableStyle.HeaderRow.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted
            customTableStyle.HeaderRow.Style.Border.Bottom.Color.SetColor(Color.Gray)

            customTableStyle.TotalRow.Style.Border.Top.Style = ExcelBorderStyle.Dotted
            customTableStyle.TotalRow.Style.Border.Top.Color.SetColor(Color.Gray)
        End Sub
        ''' <summary>
        ''' Creates a table with random data used for this sample
        ''' </summary>
        ''' <paramname="wsTables">The worksheet </param>
        ''' <paramname="tableName">The name of the table</param>
        ''' <paramname="rowStart">Start row of the table</param>
        ''' <paramname="colStart">Start column of the table</param>
        ''' <returns></returns>
        Private Function CreateTable(ByVal wsTables As ExcelWorksheet, ByVal tableName As String, ByVal Optional rowStart As Integer = 1, ByVal Optional colStart As Integer = 1) As ExcelTable
            wsTables.Cells(rowStart, colStart).Value = "Column1"
            wsTables.Cells(rowStart, colStart + 1).Value = "Column2"
            wsTables.Cells(rowStart, colStart + 2).Value = "Column3"
            wsTables.Cells(rowStart + 1, colStart).Value = 1
            wsTables.Cells(rowStart + 1, colStart + 1).Value = 2
            wsTables.Cells(rowStart + 1, colStart + 2).Value = "Type 1"

            wsTables.Cells(rowStart + 2, colStart).Value = 2
            wsTables.Cells(rowStart + 2, colStart + 1).Value = 4
            wsTables.Cells(rowStart + 2, colStart + 2).Value = "Type 2"

            wsTables.Cells(rowStart + 3, colStart).Value = 3
            wsTables.Cells(rowStart + 3, colStart + 1).Value = 7
            wsTables.Cells(rowStart + 3, colStart + 2).Value = "Type 1"

            wsTables.Cells(rowStart + 4, colStart).Value = 4
            wsTables.Cells(rowStart + 4, colStart + 1).Value = 20
            wsTables.Cells(rowStart + 4, colStart + 2).Value = "Type 3"

            wsTables.Cells(rowStart + 5, colStart).Value = 5
            wsTables.Cells(rowStart + 5, colStart + 1).Value = 43
            wsTables.Cells(rowStart + 5, colStart + 2).Value = "Type 3"

            Dim tbl = wsTables.Tables.Add(wsTables.Cells(rowStart, colStart, rowStart + 5, colStart + 2), tableName)
            tbl.Columns(0).TotalsRowFunction = RowFunctions.Sum
            tbl.Columns(1).TotalsRowFunction = RowFunctions.Sum
            tbl.ShowTotal = True
            Return tbl
        End Function
#End Region
#Region "Pivot Table Styles"

        Private Sub CreateCustomPivotTableStyleFromScratch(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a named style that applies only to pivot tables.
            Dim customPivotTableStyle = p.Workbook.Styles.CreatePivotTableStyle(styleName)

            customPivotTableStyle.WholeTable.Style.Font.Color.SetColor(ExcelIndexedColor.Indexed22)
            customPivotTableStyle.PageFieldLabels.Style.Font.Color.SetColor(Color.Red)
            customPivotTableStyle.PageFieldValues.Style.Font.Color.SetColor(eThemeSchemeColor.Accent4)

            customPivotTableStyle.HeaderRow.Style.Font.Color.SetColor(Color.DarkGray)
            customPivotTableStyle.HeaderRow.Style.Fill.Style = eDxfFillStyle.GradientFill
            customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Degree = 180

            Dim c1 = customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(0)
            c1.Color.SetColor(Color.LightBlue)

            Dim c3 = customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(100)
            c3.Color.SetColor(Color.DarkCyan)

        End Sub
        Private Sub CreateCustomPivotTableStyleFromBuildInTableStyle(ByVal p As ExcelPackage, ByVal styleName As String)
            'Create a new custom pivot table style with the build in style Medium as template.
            Dim customPivotTableStyle = p.Workbook.Styles.CreatePivotTableStyle(styleName, PivotTableStyles.Medium25)

            'Alter the font color of the entire table to theme color Text 2
            customPivotTableStyle.WholeTable.Style.Font.Color.SetColor(eThemeSchemeColor.Text2)

            customPivotTableStyle.HeaderRow.Style.Font.Italic = True
            customPivotTableStyle.TotalRow.Style.Font.Italic = True

            customPivotTableStyle.FirstColumnSubheading.Style.Fill.BackgroundColor.SetColor(Color.LightGreen)
            customPivotTableStyle.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGray)
            customPivotTableStyle.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke)
        End Sub

        Private Function CreatePivotTable(ByVal wsPivotTables As ExcelWorksheet, ByVal pivotTableName As String, ByVal tableSource As ExcelTable, ByVal range As ExcelRange) As ExcelPivotTable
            Dim pt = wsPivotTables.PivotTables.Add(range, tableSource, pivotTableName)
            pt.RowFields.Add(pt.Fields(0))
            pt.DataFields.Add(pt.Fields(1))
            pt.PageFields.Add(pt.Fields(2))
            Return pt
        End Function
#End Region
    End Module
End Namespace
