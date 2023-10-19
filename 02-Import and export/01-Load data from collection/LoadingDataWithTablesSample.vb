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
Imports System.Linq
Imports System.IO
Imports OfficeOpenXml
Imports System.Data
Imports OfficeOpenXml.Table
Imports System.Reflection
Imports System.ComponentModel

Namespace EPPlusSamples.LoadingData
    ''' <summary>
    ''' This class shows how to load data in a few ways
    ''' </summary>
    Public Module LoadingDataWithTablesSample
        Public Sub Run()
            Dim pck = New ExcelPackage()

            'Create a datatable with the directories and files from the current directory...
            Dim dt = GetDataTable(FileUtil.GetDirectoryInfo("."))

            Dim wsDt = pck.Workbook.Worksheets.Add("FromDataTable")

            'Load the datatable and set the number formats...
            wsDt.Cells("A1").LoadFromDataTable(dt, True, TableStyles.Medium9)
            wsDt.Cells(2, 2, dt.Rows.Count + 1, 2).Style.Numberformat.Format = "#,##0"
            wsDt.Cells(2, 3, dt.Rows.Count + 1, 4).Style.Numberformat.Format = "mm-dd-yy"
            wsDt.Cells(wsDt.Dimension.Address).AutoFitColumns()


            'Select Name and Created-time...
            Dim collection = (From row In dt.Select() Select New With {
                .Name = row("Name"),
                .Created_time = CDate(row("Created"))
            })

            Dim wsEnum = pck.Workbook.Worksheets.Add("FromAnonymous")

            'Load the collection starting from cell A1...
            wsEnum.Cells("A1").LoadFromCollection(collection, True, TableStyles.Medium9)

            'Add some formating...
            wsEnum.Cells(2, 2, dt.Rows.Count - 1, 2).Style.Numberformat.Format = "mm-dd-yy"
            wsEnum.Cells(wsEnum.Dimension.Address).AutoFitColumns()

            'Load a list of FileDTO objects from the datatable...
            Dim wsList = pck.Workbook.Worksheets.Add("FromList")
            Dim list As List(Of FileDTO) = (From row In dt.Select() Select New FileDTO With {
                .Name = row(CStr("Name")).ToString(),
                .Size = If(row(CStr("Size")).GetType() Is GetType(Long), CLng(row("Size")), 0),
                .Created = CDate(row("Created")),
                .LastModified = CDate(row("Modified")),
                .IsDirectory = row("Size") Is DBNull.Value
                }).ToList()

            'Load files ordered by size...
            wsList.Cells("A1").LoadFromCollection((From file In list Order By file.Size Where file.IsDirectory = False Select file), True, TableStyles.Medium9)

            wsList.Cells(2, 2, dt.Rows.Count + 1, 2).Style.Numberformat.Format = "#,##0"
            wsList.Cells(2, 3, dt.Rows.Count + 1, 4).Style.Numberformat.Format = "mm-dd-yy"


            'Load directories ordered by Name...
            wsList.Cells("F1").LoadFromCollection((From file In list Order By file.Name Where file.IsDirectory = True Select New With {
                file.Name,
                file.Created,
                .Last_modified = file.LastModified
                }), True, TableStyles.Medium11) 'Use an underscore in the property name to get a space in the title.

            wsList.Cells(2, 7, dt.Rows.Count + 1, 8).Style.Numberformat.Format = "mm-dd-yy"

            'Load the list using a specified array of MemberInfo objects. Properties, fields and methods are supported.
            Dim rng = wsList.Cells("J1").LoadFromCollection(list, True, TableStyles.Medium10, BindingFlags.Instance Or BindingFlags.Public, New MemberInfo() {GetType(FileDTO).GetProperty("Name"), GetType(FileDTO).GetField("IsDirectory"), GetType(FileDTO).GetMethod("ToString")})

            wsList.Tables.GetFromRange(rng).Columns(2).Name = "Description"

            wsList.Cells(wsList.Dimension.Address).AutoFitColumns()

            '...and save
            Dim fi = FileUtil.GetCleanFileInfo("2.1-LoadingData.xlsx")
            pck.SaveAs(fi)
            pck.Dispose()
        End Sub
        Private Function GetDataTable(ByVal dir As DirectoryInfo) As DataTable
            Dim dt As DataTable = New DataTable("RootDir")
            dt.Columns.Add("Name", GetType(String))
            dt.Columns.Add("Size", GetType(Long))
            dt.Columns.Add("Created", GetType(Date))
            dt.Columns.Add("Modified", GetType(Date))
            For Each item In dir.GetDirectories()
                Dim row = dt.NewRow()
                row("Name") = item.Name
                row("Created") = item.CreationTime
                row("Modified") = item.LastWriteTime

                dt.Rows.Add(row)
            Next
            For Each item In dir.GetFiles()
                Dim row = dt.NewRow()
                row("Name") = item.Name
                row("Size") = item.Length
                row("Created") = item.CreationTime
                row("Modified") = item.LastWriteTime

                dt.Rows.Add(row)
            Next
            Return dt
        End Function
    End Module
    Public Class FileDTO
        Public Property Name As String
        Public Property Size As Long
        Public Property Created As Date
        <Description("Last Modified")>
        Public Property LastModified As Date

        <Description("Is a Directory")>
        Public IsDirectory As Boolean = False                  'This is a field variable

        Public Overrides Function ToString() As String
            If IsDirectory Then
                Return Name & vbTab & "<Dir>"
            Else
                Return Name & vbTab & Size.ToString("#,##0")
            End If
        End Function
    End Class
End Namespace
