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
' 01/01/2025         EPPlus Software AB           Initial release EPPlus 8
' ***********************************************************************************************
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Drawing.OleObject
Imports System.IO

Namespace EPPlusSamples._05_Drawings_charts_and_themes._06_OLE_Objects
    ''' <summary>
    ''' This sample shows how to embed or link files as OLE Objects using EPPLUS.
    ''' </summary>
    Public Module OLEObjectsSample
        Public Sub Run()
            Dim myPDF = FileUtil.GetFileInfo("05-Drawings charts and themes\06-OLE Objects", "MyPDF.pdf")
            Dim myWord = FileUtil.GetFileInfo("05-Drawings charts and themes\06-OLE Objects", "MyWord.docx")
            Dim myTxt = FileUtil.GetFileInfo("05-Drawings charts and themes\06-OLE Objects", "MyTextDocument.txt")
            Dim myIcon = FileUtil.GetFileInfo("05-Drawings charts and themes\06-OLE Objects", "SampleIcon.bmp")
            Dim newWorkbook = FileUtil.GetCleanFileInfo("5.6-OLE Objects.xlsx")


            ' Embedding a file.    
            'Create a workbook and a worksheet.
            Dim p = New ExcelPackage()
            Dim ws = p.Workbook.Worksheets.Add("Sheet 1")
            'Embed the file using AddOleObject method on the drawing.
            Dim EmbeddedWord = ws.Drawings.AddOleObject("MyWord", myWord)
            'Save the workbook
            p.SaveAs(newWorkbook)


            ' Link a file.    
            'Create a workbook and a worksheet.
            Dim p2 = New ExcelPackage(newWorkbook)
            Dim ws2 = p2.Workbook.Worksheets.Add("Sheet 2")
            'Link the file using AddOleObject method on the drawing.
            Dim LinkedPDF = ws2.Drawings.AddOleObject("MyPDF", myPDF, Sub(o) o.LinkToFile = True)

            'Save the workbook
            p2.SaveAs(newWorkbook)

            ' Link a file with ExcelOleObjectParameters.    
            ' Create a workbook and a worksheet.
            Dim p3 = New ExcelPackage(newWorkbook)
            Dim ws3 = p3.Workbook.Worksheets.Add("Sheet 3")
            'Link the file using AddOleObject method on the drawing with additional parameters.
            Dim LinkedPDF2 = ws3.Drawings.AddOleObject("MyPDF", myPDF, Sub(o)
                                                                           o.DisplayAsIcon = True
                                                                           o.LinkToFile = True
                                                                       End Sub)

            'Save the workbook
            p3.SaveAs(newWorkbook)


            ' Add custom icon.    
            'Create a workbook and a worksheet.
            Dim p4 = New ExcelPackage(newWorkbook)
            Dim ws4 = p4.Workbook.Worksheets.Add("Sheet 4")
            'Link the file using AddOleObject method on the drawing.
            Dim txt = ws4.Drawings.AddOleObject("MyText", myTxt, Sub(o)
                                                                     o.DisplayAsIcon = True
                                                                     o.LinkToFile = True
                                                                     o.Icon = New ExcelImage(myIcon)
                                                                 End Sub)
            'Save the workbook
            p4.SaveAs(newWorkbook)


            ' Copy OLE Object    
            'Create a workbook, get the worksheet and create a new worksheet.
            Dim p5 = New ExcelPackage(newWorkbook)
            Dim ws1 = p5.Workbook.Worksheets(0)
            Dim ws5 = p5.Workbook.Worksheets.Add("Sheet 5")
            'Get OLE Object.
            Dim newPDF = TryCast(ws1.Drawings(0), ExcelOleObject)
            'Copy OLE Object to a new worksheet.
            Dim copy = newPDF.Copy(ws5, 1, 4)
            'Save the workbook
            p5.SaveAs(newWorkbook)


            ' Delete OLE Object    
            'Create a workbook and get worksheet.
            Dim p6 = New ExcelPackage(newWorkbook)
            Dim ws6 = p6.Workbook.Worksheets.Add("Delete OLE object")
            'Get the OLE Object from worksheet 1.
            Dim myPdfDoc = TryCast(ws1.Drawings(0), ExcelOleObject)

            'Copy the OLE Object to a new worksheet.
            Dim copyToDelete = myPdfDoc.Copy(ws6, 1, 4)
            'Now remove the OLE Object from the worksheet.
            ws6.Drawings.Remove(copyToDelete)

            'Save the workbook
            p6.SaveAs(newWorkbook)


            ' Create OLE Object using a stream 
            'Create a workbook and create a new worksheet.
            Dim p7 = New ExcelPackage(newWorkbook)
            Dim ws7 = p7.Workbook.Worksheets.Add("Sheet 7")
            'Create the stream
            Using fileStream As FileStream = New FileStream(myPDF.FullName, FileMode.Open, FileAccess.Read)
                'Add OLE Object using stream and filename.
                Dim oleFromStream = ws7.Drawings.AddOleObject("MyPdfFromStream", fileStream, "MyPdfFromStream.pdf")
            End Using
            p7.SaveAs(newWorkbook)
        End Sub
    End Module
End Namespace
