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
Imports System.Collections.Generic
Imports System.IO
Imports System.Drawing
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.Chart
Imports System.Drawing.Imaging
Imports OfficeOpenXml.Style
Imports OfficeOpenXml.Table
Imports OfficeOpenXml.Drawing.Chart.Style

Namespace EPPlusSamples.Styling
    ''' <summary>
    ''' Sample 6 - Reads the filesystem and makes a report.
    ''' </summary>                  
    Friend Class CreateAFileSystemReport
        Public Class StatItem
            Implements IComparable(Of StatItem)
            Public Property Name As String
            Public Property Count As Integer
            Public Property Size As Long

#Region "IComparable<StatItem> Members"

            'Default compare Size
            Public Function CompareTo(ByVal other As StatItem) As Integer Implements IComparable(Of StatItem).CompareTo
                Return If(Size < other.Size, -1, If(Size > other.Size, 1, 0))
            End Function

#End Region
        End Class
        Private Shared _maxLevels As Integer

        Private Shared _extStat As Dictionary(Of String, StatItem) = New Dictionary(Of String, StatItem)()
        Private Shared _fileSize As List(Of StatItem) = New List(Of StatItem)()
        ''' <summary>
        ''' Sample 6 - Reads the filesystem and makes a report.
        ''' </summary>
        ''' <paramname="outputDir">Output directory</param>
        ''' <paramname="dir">Directory to scan</param>
        ''' <paramname="depth">How many levels?</param>
        ''' <paramname="skipIcons">Skip the icons in column A. A lot faster</param>
        Public Shared Sub Run(ByVal dir As DirectoryInfo, ByVal depth As Integer, ByVal skipIcons As Boolean)
            Console.WriteLine("Running sample 20")
            _maxLevels = depth

            Dim newFile = FileUtil.GetCleanFileInfo("2.4-CreateAFileSystemReport.xlsx")

            'Create the workbook
            Dim pck As ExcelPackage = New ExcelPackage(newFile)
            'Add the Content sheet
            Dim ws = pck.Workbook.Worksheets.Add("Content")

            ws.View.ShowGridLines = False

            ws.Columns(1).Width = 2.5
            ws.Columns(2).Width = 60
            ws.Columns(3).Width = 16
            ws.Columns(4, 5).Width = 20

            'This set the outline for column 4 and 5 and hide them
            ws.Columns(4, 5).OutlineLevel = 1
            ws.Columns(4, 5).Collapsed = True
            ws.OutLineSummaryRight = True

            'Headers
            ws.Cells("B1").Value = "Name"
            ws.Cells("C1").Value = "Size"
            ws.Cells("D1").Value = "Created"
            ws.Cells("E1").Value = "Last modified"
            ws.Cells("B1:E1").Style.Font.Bold = True

            ws.View.FreezePanes(2, 1)
            ws.Select("A2")
            'height is 20 pixels 
            Dim height = 20 * 0.75
            'Start at row 2;
            Dim row = 2

            'Load the directory content to sheet 1
            row = AddDirectory(ws, dir, row, height, skipIcons)
            ws.OutLineSummaryBelow = False

            'Format columns
            ws.Cells(1, 3, row - 1, 3).Style.Numberformat.Format = "#,##0"
            ws.Cells(1, 4, row - 1, 4).Style.Numberformat.Format = "yyyy-MM-dd hh:mm"
            ws.Cells(1, 5, row - 1, 5).Style.Numberformat.Format = "yyyy-MM-dd hh:mm"

            'Add the textbox
            Dim shape = ws.Drawings.AddShape("txtDesc", eShapeStyle.Rect)
            shape.SetPosition(1, 5, 6, 5)
            shape.SetSize(400, 200)
            shape.EditAs = eEditAs.Absolute
            shape.Text = "This example demonstrates how to create various drawing objects like pictures, shapes and charts." & vbLf & vbCrLf & vbCr & "The first sheet contains all subdirectories and files with an icon, name, size and dates." & vbCrLf & vbCrLf & "The second sheet contains statistics about extensions and the top-10 largest files."
            shape.Fill.Style = eFillStyle.SolidFill
            shape.Fill.Color = Color.DarkSlateGray
            shape.Fill.Transparancy = 20
            shape.TextAnchoring = eTextAnchoringType.Top
            shape.TextVertical = eTextVerticalType.Horizontal
            shape.TextAnchoringControl = False

            shape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight)
            shape.Effect.SetPresetGlow(ePresetExcelGlowType.Accent3_8Pt)

            ws.Calculate()
            ws.Cells(1, 2, row, 5).AutoFitColumns()

            'Add the graph sheet
            AddGraphs(pck, row, dir.FullName)


            'Add a drawing with a HyperLink to the statistics sheet.
            'We add the hyperlink as a drawing here, as we don't want it to move when we expand and collapse rows.. 
            Dim hl = ws.Drawings.AddShape("HyperLink", eShapeStyle.Rect)
            hl.Hyperlink = New ExcelHyperLink("Statistics!A1", "Statistics")
            hl.SetPosition(13, 0, 9, 0)
            hl.SetSize(70, 30)
            hl.EditAs = eEditAs.Absolute
            hl.Border.Fill.Style = eFillStyle.NoFill
            hl.Fill.Style = eFillStyle.NoFill
            hl.Text = "Statistics"
            hl.Font.UnderLine = eUnderLineType.Single
            hl.Font.Fill.Color = Color.Blue

            ' Collaps children to level 1 for each row under the root.
            ws.Rows(2).SetVisibleOutlineLevel(1)

            'Printer settings
            ws.PrinterSettings.FitToPage = True
            ws.PrinterSettings.FitToWidth = 1
            ws.PrinterSettings.FitToHeight = 0
            ws.PrinterSettings.RepeatRows = New ExcelAddress("1:1") 'Print titles
            ws.PrinterSettings.PrintArea = ws.Cells(1, 1, row - 1, 5)
            pck.Workbook.Calculate()

            'Done! save the sheet
            pck.Save()

            Console.WriteLine("Sample 20 created:", newFile.FullName)
            Console.WriteLine()
        End Sub
        ''' <summary>
        ''' This method adds the comment to the header row
        ''' </summary>
        ''' <paramname="ws"></param>
        Private Shared Sub AddComments(ByVal ws As ExcelWorksheet)
            'Add Comments using the range class
            Dim comment = ws.Cells("A3").AddComment("Jan Källman:" & vbCrLf, "JK")
            comment.Font.Bold = True
            Dim rt = comment.RichText.Add("This column contains the extensions.")
            rt.Bold = False
            comment.AutoFit = True

            'Add a comment using the Comment collection
            comment = ws.Comments.Add(ws.Cells("B3"), "This column contains the size of the files.", "JK")
            'This sets the size and position. (The position is only when the comment is visible)
            comment.From.Column = 7
            comment.From.Row = 3
            comment.To.Column = 16
            comment.To.Row = 8
            comment.BackgroundColor = Color.White
            comment.RichText.Add(vbCrLf & "To format the numbers use the Numberformat-property like:" & vbCrLf)

            ws.Cells("B3:B42").Style.Numberformat.Format = "#,##0"

            'Format the code using the RichText Collection
            Dim rc = comment.RichText.Add("//Format the Size and Count column" & vbCrLf)
            rc.FontName = "Courier New"
            rc.Color = Color.FromArgb(0, 128, 0)
            rc = comment.RichText.Add("ws.Cells[")
            rc.Color = Color.Black
            rc = comment.RichText.Add("""B3:B42""")
            rc.Color = Color.FromArgb(123, 21, 21)
            rc = comment.RichText.Add("].Style.Numberformat.Format = ")
            rc.Color = Color.Black
            rc = comment.RichText.Add("""#,##0""")
            rc.Color = Color.FromArgb(123, 21, 21)
            rc = comment.RichText.Add(";")
            rc.Color = Color.Black
        End Sub
        ''' <summary>
        ''' Add the second sheet containg the graphs
        ''' </summary>
        ''' <paramname="pck">Package</param>
        ''' <paramname="rows"></param>
        ''' <paramname="header"></param>
        Private Shared Sub AddGraphs(ByVal pck As ExcelPackage, ByVal rows As Integer, ByVal dir As String)
            Dim ws = pck.Workbook.Worksheets.Add("Statistics")
            ws.View.ShowGridLines = False

            'Set first the header and format it
            ws.Cells("A1").Value = "Statistics for "
            Using r = ws.Cells("A1:O1")
                r.Merge = True
                r.Style.Font.SetFromFont("Arial", 22)
                r.Style.Font.Color.SetColor(Color.White)
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous
                r.Style.Fill.PatternType = ExcelFillStyle.Solid
                r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93))
            End Using

            'Use the RichText property to change the font for the directory part of the cell
            Dim rtDir = ws.Cells("A1").RichText.Add(dir)
            rtDir.FontName = "Consolas"
            rtDir.Size = 18

            'Start with the Extention Size 
            Dim lst As List(Of StatItem) = New List(Of StatItem)(_extStat.Values)
            lst.Sort()

            'Add rows
            Dim row = AddStatRows(ws, lst, 2, "Extensions", "Size")

            'Add commets to the Extensions header
            AddComments(ws)

            'Add the piechart
            Dim pieChart = ws.Drawings.AddPieChart("crtExtensionsSize", ePieChartType.PieExploded3D)
            'Set top left corner to row 1 column 2
            pieChart.SetPosition(1, 0, 2, 0)
            pieChart.SetSize(400, 400)
            pieChart.Series.Add(ExcelCellBase.GetAddress(3, 2, row - 1, 2), ExcelCellBase.GetAddress(3, 1, row - 1, 1))

            pieChart.Title.Text = "Extension Size"
            'Set datalabels and remove the legend
            pieChart.DataLabel.ShowCategory = True
            pieChart.DataLabel.ShowPercent = True
            pieChart.DataLabel.ShowLeaderLines = True
            pieChart.Legend.Remove()
            pieChart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle6)

            'Resort on Count and add the rows

            lst.Sort(Function(first, second) If(first.Count < second.Count, -1, If(first.Count > second.Count, 1, 0)))
            row = AddStatRows(ws, lst, 16, "", "Count")

            'Add the Doughnut chart
            Dim doughtnutChart = TryCast(ws.Drawings.AddDoughnutChart("crtExtensionCount", eDoughnutChartType.DoughnutExploded), ExcelDoughnutChart)
            'Set position to row 1 column 7 and 16 pixels offset
            doughtnutChart.SetPosition(1, 0, 8, 16)
            doughtnutChart.SetSize(400, 400)
            doughtnutChart.Series.Add(ExcelCellBase.GetAddress(16, 2, row - 1, 2), ExcelCellBase.GetAddress(16, 1, row - 1, 1))

            doughtnutChart.Title.Text = "Extension Count"
            doughtnutChart.DataLabel.ShowPercent = True
            doughtnutChart.DataLabel.ShowLeaderLines = True
            doughtnutChart.StyleManager.SetChartStyle(ePresetChartStyle.DoughnutChartStyle8)

            'Top-10 filesize
            Call _fileSize.Sort()
            row = AddStatRows(ws, _fileSize, 29, "Files", "Size")
            Dim barChart = TryCast(ws.Drawings.AddBarChart("crtFiles", eBarChartType.BarClustered3D), ExcelBarChart)
            '3d Settings
            barChart.View3D.RotX = 0
            barChart.View3D.Perspective = 0

            barChart.SetPosition(22, 0, 2, 0)
            barChart.SetSize(800, 398)
            barChart.Series.Add(ExcelCellBase.GetAddress(30, 2, row - 1, 2), ExcelCellBase.GetAddress(30, 1, row - 1, 1))
            'barChart.Series[0].Header = "Size";
            barChart.Title.Text = "Top File size"
            barChart.StyleManager.SetChartStyle(ePresetChartStyle.Bar3dChartStyle9)
            'Format the Size and Count column
            ws.Cells("B3:B42").Style.Numberformat.Format = "#,##0"
            'Set a border around
            ws.Cells("A1:A43").Style.Border.Left.Style = ExcelBorderStyle.Thin
            ws.Cells("A1:O1").Style.Border.Top.Style = ExcelBorderStyle.Thin
            ws.Cells("O1:O43").Style.Border.Right.Style = ExcelBorderStyle.Thin
            ws.Cells("A43:O43").Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            ws.Cells(1, 1, row, 2).AutoFitColumns(1)

            'And last the printersettings
            ws.PrinterSettings.Orientation = eOrientation.Landscape
            ws.PrinterSettings.FitToPage = True
            ws.PrinterSettings.Scale = 67
        End Sub
        ''' <summary>
        ''' Add statistic-rows to the statistics sheet.
        ''' </summary>
        ''' <paramname="ws">Worksheet</param>
        ''' <paramname="lst">List with statistics</param>
        ''' <paramname="startRow"></param>
        ''' <paramname="header">Header text</param>
        ''' <paramname="propertyName">Size or Count</param>
        ''' <returns></returns>
        Private Shared Function AddStatRows(ByVal ws As ExcelWorksheet, ByVal lst As List(Of StatItem), ByVal startRow As Integer, ByVal header As String, ByVal propertyName As String) As Integer
            'Add Headers
            Dim row = startRow
            If Not Equals(header, "") Then
                ws.Cells(row, 1).Value = header
                Using r = ws.Cells(row, 1, row, 2)
                    r.Merge = True
                    r.Style.Font.SetFromFont("Arial", 16, False, True)
                    r.Style.Font.Color.SetColor(Color.White)
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                End Using
                row += 1
            End If

            Dim tblStart = row
            'Header 2
            ws.Cells(row, 1).Value = "Name"
            ws.Cells(row, 2).Value = propertyName
            Using r = ws.Cells(row, 1, row, 2)
                r.Style.Font.SetFromFont("Arial", 12, True)
            End Using

            row += 1
            'Add top 10 rows
            For i = 0 To 9
                If lst.Count - i > 0 Then
                    ws.Cells(row, 1).Value = lst(lst.Count - i - 1).Name
                    If Equals(propertyName, "Size") Then
                        ws.Cells(row, 2).Value = lst(lst.Count - i - 1).Size
                    Else
                        ws.Cells(row, 2).Value = lst(lst.Count - i - 1).Count
                    End If

                    row += 1
                End If
            Next

            'If we have more than 10 items, sum...
            Dim rest As Long = 0
            For i = 0 To lst.Count - 10 - 1
                If Equals(propertyName, "Size") Then
                    rest += lst(i).Size
                Else
                    rest += lst(i).Count
                End If
            Next
            '... and add anothers row
            If rest > 0 Then
                ws.Cells(row, 1).Value = "Others"
                ws.Cells(row, 2).Value = rest
                ws.Cells(row, 1, row, 2).Style.Fill.PatternType = ExcelFillStyle.Solid
                ws.Cells(row, 1, row, 2).Style.Fill.BackgroundColor.SetColor(Color.LightGray)
                row += 1
            End If

            Dim tbl = ws.Tables.Add(ws.Cells(tblStart, 1, row - 1, 2), Nothing)
            tbl.TableStyle = TableStyles.Medium16
            tbl.ShowTotal = True
            tbl.Columns(1).TotalsRowFunction = RowFunctions.Sum
            Return row
        End Function
        ''' <summary>
        ''' Just alters the colors in the list
        ''' </summary>
        ''' <paramname="ws">The worksheet</param>
        ''' <paramname="row">Startrow</param>
        Private Shared Sub AlterColor(ByVal ws As ExcelWorksheet, ByVal row As Integer)
            Using rowRange = ws.Cells(row, 1, row, 2)
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid
                If row Mod 2 = 1 Then
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray)
                Else
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightYellow)
                End If
            End Using
        End Sub

        Private Shared Function AddDirectory(ByVal ws As ExcelWorksheet, ByVal dir As DirectoryInfo, ByVal row As Integer, ByVal height As Double, ByVal skipIcons As Boolean) As Integer
            'Get the icon as a bitmap
            Console.WriteLine("Directory " & dir.Name)
            If Not skipIcons Then
                Dim icon = GetIcon(dir.FullName)
                ws.Rows(row).Height = height
                'Add the icon as a picture
                If icon IsNot Nothing Then
                    Using ms = New MemoryStream()
                        icon.Save(ms, ImageFormat.Bmp)
                        Dim pic As ExcelPicture = ws.Drawings.AddPicture("pic" & row.ToString(), ms)
                        pic.SetPosition(20 * (row - 1) + 2, 0)
                    End Using
                End If
            End If
            ws.Cells(row, 2).Value = dir.Name
            ws.Cells(row, 4).Value = dir.CreationTime
            ws.Cells(row, 5).Value = dir.LastAccessTime

            ws.Cells(row, 2, row, 5).Style.Font.Bold = True

            Dim prevRow = row
            row += 1
            'Add subdirectories
            For Each subDir In dir.GetDirectories()
                If ws.Rows(prevRow).OutlineLevel < _maxLevels Then
                    row = AddDirectory(ws, subDir, row, height, skipIcons)
                End If
            Next

            'Add files in the directory
            For Each file In dir.GetFiles()
                If Not skipIcons Then
                    Dim fileIcon = GetIcon(file.FullName)

                    ws.Rows(row).Height = height
                    If fileIcon IsNot Nothing Then
                        Using ms = New MemoryStream()
                            fileIcon.Save(ms, ImageFormat.Bmp)
                            Dim pic As ExcelPicture = ws.Drawings.AddPicture("pic" & row.ToString(), ms)
                            pic.SetPosition(20 * (row - 1) + 2, 0)
                        End Using
                    End If
                End If

                ws.Cells(row, 2).Value = file.Name
                ws.Cells(row, 3).Value = file.Length
                ws.Cells(row, 4).Value = file.CreationTime
                ws.Cells(row, 5).Value = file.LastAccessTime

                AddStatistics(file)

                row += 1
            Next

            'If the directory has children, group them. The Group method adds one to the Outline level.
            If prevRow < row - 1 Then
                ws.Rows(prevRow + 1, row - 1).Group()
            End If

            'Add a subtotal for the directory
            If row - 1 > prevRow Then
                ws.Cells(prevRow, 3).Formula = String.Format("SUBTOTAL(9, {0})", ExcelCellBase.GetAddress(prevRow + 1, 3, row - 1, 3))
            Else
                ws.Cells(prevRow, 3).Value = 0
            End If

            Return row
        End Function
        ''' <summary>
        ''' Add statistics to the collections 
        ''' </summary>
        ''' <paramname="file"></param>
        Private Shared Sub AddStatistics(ByVal file As FileInfo)
            'Extension
            If _extStat.ContainsKey(file.Extension) Then
                _extStat(file.Extension).Count += 1
                _extStat(file.Extension).Size += file.Length
            Else
                Dim ext = If(file.Extension.Length > 0, file.Extension.Remove(0, 1), "")
                Call _extStat.Add(file.Extension, New StatItem() With {
                    .Name = ext,
                    .Count = 1,
                    .Size = file.Length
                })
            End If

            'File top 10;
            If _fileSize.Count < 10 Then
                Call _fileSize.Add(New StatItem With {
                    .Name = file.Name,
                    .Size = file.Length
                })
                If _fileSize.Count = 10 Then
                    Call _fileSize.Sort()
                End If
            ElseIf _fileSize(0).Size < file.Length Then
                _fileSize.RemoveAt(0)
                Call _fileSize.Add(New StatItem With {
                    .Name = file.Name,
                    .Size = file.Length
                })
                Call _fileSize.Sort()
            End If
        End Sub
        ''' <summary>
        ''' Gets the icon for a file or directory
        ''' </summary>
        ''' <paramname="FileName"></param>
        ''' <returns></returns>
        Private Shared Function GetIcon(ByVal FileName As String) As Bitmap
            If File.Exists(FileName) Then
                Dim bmp = Icon.ExtractAssociatedIcon(FileName).ToBitmap()
                Return New Bitmap(bmp, New Size(16, 16))
            Else
                Return Nothing
            End If
        End Function
    End Class
End Namespace
