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
' 07/22/2020         EPPlus Software AB           EPPlus 5.2.1
' ***********************************************************************************************
Imports Newtonsoft.Json
Imports OfficeOpenXml
Imports OfficeOpenXml.LoadFunctions.Params
Imports OfficeOpenXml.Table
Imports System
Imports System.Collections.Generic
Imports System.Dynamic
Imports System.IO

Namespace EPPlusSamples.LoadingData
    Public Module LoadingDataWithDynamicObjects
        Public Sub Run()
            ' Create a list of dynamic objects
            Dim p1 As Object = New ExpandoObject()
            p1.Id = 1
            p1.FirstName = "Ivan"
            p1.LastName = "Horvat"
            p1.Age = 21
            Dim p2 As Object = New ExpandoObject()
            p2.Id = 2
            p2.FirstName = "John"
            p2.LastName = "Doe"
            p2.Age = 45
            Dim p3 As Object = New ExpandoObject()
            p3.Id = 3
            p3.FirstName = "Sven"
            p3.LastName = "Svensson"
            p3.Age = 68

            Dim items As List(Of ExpandoObject) = New List(Of ExpandoObject)() From {
    p1,
    p2,
    p3
}

            ' Create a workbook with a worksheet and load the data into a table
            Using package = New ExcelPackage(FileUtil.GetCleanFileInfo("2.1-LoadDynamicObjects.xlsx"))
                Dim sheet = package.Workbook.Worksheets.Add("Dynamic")
                sheet.Cells("A1").LoadFromDictionaries(items, Sub(c)
                                                                  ' Print headers using the property names
                                                                  c.PrintHeaders = True
                                                                  ' insert a space before each capital letter in the header
                                                                  c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace
                                                                  ' when TableStyle is not TableStyles.None the data will be loaded into a table with the 
                                                                  ' selected style.
                                                                  c.TableStyle = TableStyles.Medium1
                                                              End Sub)
                package.Save()
            End Using

            ' Load data from json (in this case a file)
            Dim jsonItems = JsonConvert.DeserializeObject(Of IEnumerable(Of ExpandoObject))(File.ReadAllText(FileUtil.GetFileInfo("02-Import and Export\01-Load data from collection", "testdata.json").FullName))
            Using package = New ExcelPackage(FileUtil.GetCleanFileInfo("2.1-LoadJsonFromFile.xlsx"))
                Dim sheet = package.Workbook.Worksheets.Add("Dynamic")
                sheet.Cells("A1").LoadFromDictionaries(jsonItems, Sub(c)
                                                                      ' Print headers using the property names
                                                                      c.PrintHeaders = True
                                                                      ' insert a space before each capital letter in the header
                                                                      c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace
                                                                      ' when TableStyle is not TableStyles.None the data will be loaded into a table with the 
                                                                      ' selected style.
                                                                      c.TableStyle = TableStyles.Medium1
                                                                  End Sub)
                sheet.Cells("D:D").Style.Numberformat.Format = "yyyy-mm-dd"
                sheet.Cells(1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column).AutoFitColumns()
                package.Save()
            End Using
        End Sub
    End Module
End Namespace
