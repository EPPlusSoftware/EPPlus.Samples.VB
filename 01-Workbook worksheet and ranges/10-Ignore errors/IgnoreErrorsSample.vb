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
Imports OfficeOpenXml
Imports System

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module IgnoreErrorsSample
        Public Sub Run()
            Console.WriteLine("Running sample 1.10-Ignore Errors")
            Dim p = New ExcelPackage()

            Dim ws = p.Workbook.Worksheets.Add("IgnoreErrors")

            'Suppress Number stored as text
            ws.Cells("A1").Value = "1"
            ws.Cells("A2").Value = "2"
            Dim ie = ws.IgnoredErrors.Add(ws.Cells("A2"))
            ie.NumberStoredAsText = True                   ' Ignore errors on A2 only
            ws.Cells("A2").AddComment("Number stored as text error is ignored here", "EPPlus Sample")

            ws.Cells("C1").Value = "1"
            ws.Cells("C2").Value = "2"
            ws.Cells("C3").Value = "3"
            ws.Cells("C4").Value = "4"
            ws.Cells("C5").Value = "5"
            ie = ws.IgnoredErrors.Add(ws.Cells("C1:C5"))   ' Ignore errors on the range
            ie.NumberStoredAsText = True

            ws.Cells("D1:D5").Formula = "A1+C1"
            ws.Cells("D2").Formula = "A2+B2"               'This will generate a Inconsistant formula error
            ws.Cells("D4").Formula = "A1+B2"               'This will generate a Inconsistant formula error
            ws.Cells("D2,D4").AddComment("Inconsistant formula error is ignored here", "EPPlus Sample")

            ie = ws.IgnoredErrors.Add(ws.Cells("D2,D4"))
            ie.Formula = True                              ' Ignore the inconsistant formula error

            p.SaveAs(FileUtil.GetCleanFileInfo("1.10-IgnoreErrors.xlsx"))

            Console.WriteLine("Sample 1.10 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub
    End Module
End Namespace