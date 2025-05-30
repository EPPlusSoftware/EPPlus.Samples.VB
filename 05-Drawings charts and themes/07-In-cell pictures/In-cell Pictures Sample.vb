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
' 01/01/2025         EPPlus Software AB           Initial release EPPlus 8
' ***********************************************************************************************
Imports OfficeOpenXml

Namespace EPPlusSamples
    ''' <summary>
    ''' This sample shows how to work with in-cell picures directly and via formulas.
    ''' </summary>
    Public Module InCellPicturesSample
        Public Sub Run()
            Dim myPic = FileUtil.GetFileInfo("05-Drawings charts and themes\07-In-cell Pictures", "EPPlusLogo.jpg")
            Using package = New ExcelPackage()
                Dim ws = package.Workbook.Worksheets.Add("InCell Pictures")

                ws.Cells("A1").Picture.Set(myPic, "Alt text for image")
                ws.Cells("B1").Formula = "A1" ' Create a link to the picture.
                ws.Cells("C1").SetFormula("Image(""https://samples.epplussoftware.com/img/EPPlus-logo-full.png"")")                       'The image function fetches an image file from an url. Https is required.
                ws.Cells("D1").SetFormula("Image(""https://samples.epplussoftware.com/img/EPPlus-logo-full.png"",""Alt Text"",2)")        'Add the same image with an alt text and Sizing = Original Size. EPPlus will only download an image once and then cache the image file.
                ws.Calculate()                'To be able to access the images in cell B1:D1, we need to calculate the formulas.

                If ws.Cells("B1").Picture.Exists Then
                    Dim pic = ws.Cells("B1").Picture.Get()
                    Dim picTypeB1 = pic.PictureType            'We have a local image
                    Dim image = pic.GetImage()                 'You can read the bytes from the image. or use the pic.GetImageBytes directly if you know the bounds of the image.
                    Dim imageBytes = image.ImageBytes
                End If

                Dim newWorkbook = FileUtil.GetCleanFileInfo("5.7-InCellPictures.xlsx")
                package.SaveAs(newWorkbook)
            End Using
        End Sub
    End Module
End Namespace
