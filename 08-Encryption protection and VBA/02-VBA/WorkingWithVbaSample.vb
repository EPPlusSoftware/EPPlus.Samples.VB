﻿' ***********************************************************************************************
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
Imports System.Text
Imports System.IO
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml.Drawing.Chart
Imports OfficeOpenXml.Drawing

Namespace EPPlusSamples.EncryptionProtectionAndVba
    Public Module WorkingWithVbaSample
        Public Sub Run()
            'Create a macro-enabled workbook from scratch.
            Call SimpleVba()

            'Open Sample 1 and add code to change the chart to a bubble chart.
            Call AddABubbleChart()

            'Simple battleships game from scratch.
            Call CreateABattleShipsGame()
        End Sub
        Private Sub SimpleVba()
            Dim pck As ExcelPackage = New ExcelPackage()

            'Add a worksheet.
            Dim ws = pck.Workbook.Worksheets.Add("VBA Sample")
            ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect)

            'Create a vba project             
            pck.Workbook.CreateVBAProject()

            'Now add some code to update the text of the shape...
            Dim sb = New StringBuilder()

            sb.AppendLine("Private Sub Workbook_Open()")
            sb.AppendLine("    [VBA Sample].Shapes(""VBASampleRect"").TextEffect.Text = ""This text is set from VBA!""")
            sb.AppendLine("End Sub")
            pck.Workbook.CodeModule.Code = sb.ToString()

            'And Save as xlsm
            Dim fi = FileUtil.GetCleanFileInfo("8.2-01-SimpleVba.xlsm")
            pck.SaveAs(fi)
        End Sub
        Private Sub AddABubbleChart()
            Dim sample1File = FileUtil.GetFileInfo("1.01-GettingStarted.xlsx")
            'Open Sample 1 again
            Dim pck As ExcelPackage = New ExcelPackage(sample1File)
            Dim p = New ExcelPackage()
            'Create a vba project             
            pck.Workbook.CreateVBAProject()

            'Now add some code that creates a bubble chart...
            Dim sb = New StringBuilder()

            sb.AppendLine("Public Sub CreateBubbleChart()")
            sb.AppendLine("Dim co As ChartObject")
            sb.AppendLine("Set co = Inventory.ChartObjects.Add(10, 100, 400, 200)")
            sb.AppendLine("co.Chart.SetSourceData Source:=Range(""'Inventory'!$B$1:$E$5"")")
            sb.AppendLine("co.Chart.ChartType = xlBubble3DEffect         'Add a bubblechart")
            sb.AppendLine("End Sub")

            'Create a new module and set the code
            Dim [module] = pck.Workbook.VbaProject.Modules.AddModule("EPPlusGeneratedCode")
            [module].Code = sb.ToString()

            'Call the newly created sub from the workbook open event
            pck.Workbook.CodeModule.Code = "Private Sub Workbook_Open()" & vbCrLf & "CreateBubbleChart" & vbCrLf & "End Sub"

            'Optionally, Sign the code with your company certificate.
            'X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            'store.Open(OpenFlags.ReadOnly);
            'pck.Workbook.VbaProject.Signature.Certificate = store.Certificates[0];

            'And Save as xlsm
            Dim fi = FileUtil.GetCleanFileInfo("8.2-02-AddABubbleChartVba.xlsm")
            pck.SaveAs(fi)
        End Sub
        Private Sub CreateABattleShipsGame()
            'Now, lets do something a little bit more fun.
            'We are going to create a simple battleships game from scratch.

            Dim pck As ExcelPackage = New ExcelPackage()

            'Add a worksheet.
            Dim ws = pck.Workbook.Worksheets.Add("Battleship")

            ws.View.ShowGridLines = False
            ws.View.ShowHeaders = False

            ws.DefaultColWidth = 3
            ws.DefaultRowHeight = 15

            Dim gridSize = 10

            'Create the boards
            Dim board1 = ws.Cells(2, 2, 2 + gridSize - 1, 2 + gridSize - 1)
            Dim board2 = ws.Cells(2, 4 + gridSize - 1, 2 + gridSize - 1, 4 + (gridSize - 1) * 2)
            CreateBoard(board1)
            CreateBoard(board2)
            ws.Select("B2")
            ws.Protection.IsProtected = True
            ws.Protection.AllowSelectLockedCells = True

            'Create the VBA Project
            pck.Workbook.CreateVBAProject()
            'Password protect your code
            pck.Workbook.VbaProject.Protection.SetPassword("EPPlus")

            Dim codeDir = FileUtil.GetSubDirectory("08-Encryption Protection and VBA\02-VBA", "VBA-Code")

            'Add all the code from the textfiles in the Vba-Code sub-folder.
            pck.Workbook.CodeModule.Code = GetCodeModule(codeDir, "ThisWorkbook.txt")

            'Add the sheet code
            ws.CodeModule.Code = GetCodeModule(codeDir, "BattleshipSheet.txt")
            Dim m1 = pck.Workbook.VbaProject.Modules.AddModule("Code")
            Dim code = GetCodeModule(codeDir, "CodeModule.txt")

            'Insert your ships on the right board. you can changes these, but don't cheat ;)
            Dim ships = New String() {"N3:N7", "P2:S2", "V9:V11", "O10:Q10", "R11:S11"}

            'Note: For security reasons you should never mix external data and code(to avoid code injections!), especially not on a webserver. 
            'If you deside to do that anyway, be very careful with the validation of the data.
            'Be extra careful if you sign the code.
            'Read more here http://en.wikipedia.org/wiki/Code_injection

            code = String.Format(code, ships(0), ships(1), ships(2), ships(3), ships(4), board1.Address, board2.Address)  'Ships are injected into the constants in the module
            m1.Code = code

            'Ships are displayed with a black background
            Dim shipsaddress = String.Join(",", ships)
            ws.Cells(shipsaddress).Style.Fill.PatternType = ExcelFillStyle.Solid
            ws.Cells(shipsaddress).Style.Fill.BackgroundColor.SetColor(Color.Black)

            Dim m2 = pck.Workbook.VbaProject.Modules.AddModule("ComputerPlay")
            m2.Code = GetCodeModule(codeDir, "ComputerPlayModule.txt")
            Dim c1 = pck.Workbook.VbaProject.Modules.AddClass("Ship", False)
            c1.Code = GetCodeModule(codeDir, "ShipClass.txt")

            'Add the info text shape.
            Dim tb = ws.Drawings.AddShape("txtInfo", eShapeStyle.Rect)
            tb.SetPosition(1, 0, 27, 0)
            tb.Fill.Color = Color.LightSlateGray
            Dim rt1 = tb.RichText.Add("Battleships")
            rt1.Bold = True
            tb.RichText.Add(vbCrLf & "Double-click on the left board to make your move. Find and sink all ships to win!")

            'Set the headers.
            ws.SetValue("B1", "Computer Grid")
            ws.SetValue("M1", "Your Grid")
            ws.Rows(1).Style.Font.Size = 18

            AddChart(ws.Cells("B13"), "chtHitPercent", "Player")
            AddChart(ws.Cells("M13"), "chtComputerHitPercent", "Computer")

            ws.Names.Add("LogStart", ws.Cells("B24"))
            ws.Cells("B24:X224").Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
            ws.Cells("B25:X224").Style.Font.Name = "Consolas"
            ws.SetValue("B24", "Log")
            ws.Cells("B24").Style.Font.Bold = True
            ws.Cells("B24:X24").Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
            Dim cf = ws.Cells("B25:B224").ConditionalFormatting.AddContainsText()
            cf.Text = "hit"
            cf.Style.Font.Color.Color = Color.Red

            'If you have a valid certificate for code signing you can use this code to set it.
            ''' *** Try to find a cert valid for signing... ***/
            'X509Store store = new X509Store(StoreLocation.CurrentUser);
            'store.Open(OpenFlags.ReadOnly);   
            'foreach (var cert in store.Certificates)
            '{
            '    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
            '    {
            '        pck.Workbook.VbaProject.Signature.Certificate = cert;
            '        break;
            '    }
            '}

            Dim fi = FileUtil.GetCleanFileInfo("8.2-03-CreateABattleShipsGameVba.xlsm")
            pck.SaveAs(fi)
        End Sub

        Private Function GetCodeModule(ByVal codeDir As DirectoryInfo, ByVal fileName As String) As String
            Return File.ReadAllText(FileUtil.GetFileInfo(codeDir, fileName, False).FullName)
        End Function

        Private Sub AddChart(ByVal rng As ExcelRange, ByVal name As String, ByVal prefix As String)
            Dim chrt = rng.Worksheet.Drawings.AddPieChart(name, ePieChartType.Pie)
            chrt.SetPosition(rng.Start.Row - 1, 0, rng.Start.Column - 1, 0)
            chrt.To.Row = rng.Start.Row + 9
            chrt.To.Column = rng.Start.Column + 9
            chrt.Style = eChartStyle.Style18
            chrt.DataLabel.ShowPercent = True

            Dim serie = chrt.Series.Add(rng.Offset(2, 2, 1, 2), rng.Offset(1, 2, 1, 2))
            serie.Header = "Hits"

            chrt.Title.Text = "Hit ratio"

            Dim n1 = rng.Worksheet.Names.Add(prefix & "Misses", rng.Offset(2, 2))
            n1.Value = 0
            Dim n2 = rng.Worksheet.Names.Add(prefix & "Hits", rng.Offset(2, 3))
            n2.Value = 0
            rng.Offset(1, 2).Value = "Misses"
            rng.Offset(1, 3).Value = "Hits"
        End Sub

        Private Sub CreateBoard(ByVal rng As ExcelRange)
            'Create a gradiant background with one dark and one light blue color
            rng.Style.Fill.Gradient.Color1.SetColor(Color.FromArgb(&H80, &H80, &HFF))
            rng.Style.Fill.Gradient.Color2.SetColor(Color.FromArgb(&H20, &H20, &HFF))
            rng.Style.Fill.Gradient.Type = ExcelFillGradientType.None
            For col = 0 To rng.End.Column - rng.Start.Column
                For row = 0 To rng.End.Row - rng.Start.Row
                    If col Mod 4 = 0 Then
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 45
                    End If
                    If col Mod 4 = 1 Then
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 70
                    End If
                    If col Mod 4 = 2 Then
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 110
                    Else
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 135
                    End If
                Next
            Next
            'Set the inner cell border to thin, light gray
            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin
            rng.Style.Border.Top.Color.SetColor(Color.Gray)
            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin
            rng.Style.Border.Right.Color.SetColor(Color.Gray)
            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin
            rng.Style.Border.Left.Color.SetColor(Color.Gray)
            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            rng.Style.Border.Bottom.Color.SetColor(Color.Gray)

            'Solid black border around the board.
            rng.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black)
        End Sub
    End Module
End Namespace
