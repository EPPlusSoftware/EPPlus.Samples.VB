Imports EPPlusSamples.FiltersAndValidations
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Table
Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SQLite
Imports System.Threading.Tasks

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public MustInherit Class ChartSampleBase
        Private Class RegionalSales
            Public Property Region As String
            Public Property SoldUnits As Integer
            Public Property TotalSales As Double
            Public Property Margin As Double
        End Class
        Public Shared Function GetCarDataTable() As DataTable
            Dim dt = New DataTable()
            dt.Columns.Add("Car", GetType(String))
            dt.Columns.Add("Acceleration Index", GetType(Integer))
            dt.Columns.Add("Size Index", GetType(Integer))
            dt.Columns.Add("Polution Index", GetType(Integer))
            dt.Columns.Add("Retro Index", GetType(Integer))
            dt.Rows.Add("Volvo 242", 1, 3, 4, 4)
            dt.Rows.Add("Lamborghini Countach", 5, 1, 5, 4)
            dt.Rows.Add("Tesla Model S", 5, 2, 1, 1)
            dt.Rows.Add("Hummer H1", 2, 5, 5, 2)

            Return dt
        End Function

        Protected Shared Async Function LoadFromDatabase(ByVal ws As ExcelWorksheet) As Task(Of ExcelRangeBase)
            Dim range As ExcelRangeBase
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Using sqlCmd = New SQLiteCommand("select orderdate as OrderDate, SUM(ordervalue) as OrderValue, SUM(tax) As Tax,SUM(freight) As Freight from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by OrderDate ORDER BY OrderDate desc limit 15", sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        range = Await ws.Cells("A1").LoadFromDataReaderAsync(sqlReader, True)
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = True
                        range.Offset(0, 0, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd"
                    End Using
                    'Set the numberformat
                End Using
            End Using
            Return range
        End Function
        Protected Shared Async Function LoadSalesFromDatabase(ByVal ws As ExcelWorksheet) As Task(Of ExcelRangeBase)
            Dim range As ExcelRangeBase
            Using sqlConn = New SQLiteConnection(ConnectionString)
                sqlConn.Open()
                Dim sql = GroupedOrdersSql
                Using sqlCmd = New SQLiteCommand(sql, sqlConn)
                    Using sqlReader = sqlCmd.ExecuteReader()
                        range = Await ws.Cells("A1").LoadFromDataReaderAsync(sqlReader, True)
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = True
                        range.Offset(0, 3, range.Rows, 3).Style.Numberformat.Format = "#,##0"
                    End Using
                End Using
            End Using

            Return range
        End Function

        Protected Shared Function CreateIceCreamData(ByVal ws As ExcelWorksheet) As ExcelRangeBase
            ws.SetValue("A1", "Icecream Sales-2019")
            ws.SetValue("A2", "Date")
            ws.SetValue("B2", "Sales")
            ws.SetValue("A3", New DateTime(2019, 1, 1))
            ws.SetValue("B3", 2500)
            ws.SetValue("A4", New DateTime(2019, 2, 1))
            ws.SetValue("B4", 3000)
            ws.SetValue("A5", New DateTime(2019, 3, 1))
            ws.SetValue("B5", 2700)
            ws.SetValue("A6", New DateTime(2019, 4, 1))
            ws.SetValue("B6", 4400)
            ws.SetValue("A7", New DateTime(2019, 5, 1))
            ws.SetValue("B7", 6900)
            ws.SetValue("A8", New DateTime(2019, 6, 1))
            ws.SetValue("B8", 11200)
            ws.SetValue("A9", New DateTime(2019, 7, 1))
            ws.SetValue("B9", 13200)
            ws.SetValue("A10", New DateTime(2019, 8, 1))
            ws.SetValue("B10", 12400)
            ws.SetValue("A11", New DateTime(2019, 9, 1))
            ws.SetValue("B11", 8700)
            ws.SetValue("A12", New DateTime(2019, 10, 1))
            ws.SetValue("B12", 4800)
            ws.SetValue("A13", New DateTime(2019, 11, 1))
            ws.SetValue("B13", 2000)
            ws.SetValue("A14", New DateTime(2019, 12, 1))
            ws.SetValue("B14", 2400)
            ws.Cells("A3:A14").Style.Numberformat.Format = "yyyy-MM"
            ws.Cells("B3:B14").Style.Numberformat.Format = "#,##0kr"
            Return ws.Cells("A2:B14")
        End Function
        Protected Shared Function LoadBubbleChartData(ByVal package As ExcelPackage) As ExcelWorksheet
            Dim data = New List(Of RegionalSales)() From {
        New RegionalSales() With {
        .Region = "North",
        .SoldUnits = 500,
        .TotalSales = 4800,
        .Margin = 0.200
    },
        New RegionalSales() With {
        .Region = "Central",
        .SoldUnits = 900,
        .TotalSales = 7330,
        .Margin = 0.333
    },
        New RegionalSales() With {
        .Region = "South",
        .SoldUnits = 400,
        .TotalSales = 3700,
        .Margin = 0.150
    },
        New RegionalSales() With {
        .Region = "East",
        .SoldUnits = 350,
        .TotalSales = 4400,
        .Margin = 0.102
    },
        New RegionalSales() With {
        .Region = "West",
        .SoldUnits = 700,
        .TotalSales = 6900,
        .Margin = 0.218
    },
        New RegionalSales() With {
        .Region = "Stockholm",
        .SoldUnits = 1200,
        .TotalSales = 8250,
        .Margin = 0.350
    }
}
            Dim wsData = package.Workbook.Worksheets.Add("ChartData")
            wsData.Cells("A1").LoadFromCollection(data, True, TableStyles.Medium15)
            wsData.Cells("B2:C7").Style.Numberformat.Format = "#,##0"
            wsData.Cells("D2:D7").Style.Numberformat.Format = "#,##0.00%"

            Dim shape = wsData.Drawings.AddShape("Shape1", eShapeStyle.Rect)
            shape.Text = "This worksheet contains the data for the bubble-chartsheet"
            shape.SetPosition(1, 0, 6, 0)
            shape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomLeft)
            Return wsData
        End Function
    End Class
End Namespace
