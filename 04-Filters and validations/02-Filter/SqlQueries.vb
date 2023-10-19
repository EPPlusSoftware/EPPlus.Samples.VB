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
Imports System.Text

Namespace EPPlusSamples.FiltersAndValidations
    Public Module SqlStatements
        Public ReadOnly Property OrdersSql As String = GetOrdersSql()
        Public ReadOnly Property OrdersWithTaxAndFreightSql As String = GetOrdersWithTaxAndFreightSql()
        Public ReadOnly Property GroupedOrdersSql As String = GetGroupedOrdersSql()
        Private Function GetGroupedOrdersSql() As String
            Dim sb = New StringBuilder()
            sb.Append("select co.Continent, co.Name as Country, ci.name as City, SUM(OrderValue) As Sales ")
            WriteOrdersWhereSql(sb)
            sb.Append("Where Currency='USD' group by co.continent, co.name, ci.name ORDER BY co.continent, co.name, ci.name")
            Return sb.ToString()
        End Function

        Private Function GetOrdersSql() As String
            Dim sb = New StringBuilder()
            sb.Append("select cu.Name as CompanyName, sp.Name, Email as 'E-Mail', co.Name as Country, OrderId As OrderId, OrderDate As OrderDate, OrderValue As OrderValue, Currency as Currency ")
            WriteOrdersWhereSql(sb)
            sb.Append("ORDER BY 1,2 desc")
            Return sb.ToString()

        End Function
        Private Function GetOrdersWithTaxAndFreightSql() As String
            Dim sb = New StringBuilder()
            sb.Append("select cu.Name as CompanyName, sp.Name, Email as 'E-Mail', co.Name as Country, OrderId, orderdate as 'OrderDate', ordervalue as 'OrderValue',tax as Tax, freight As Freight, currency As Currency ")
            WriteOrdersWhereSql(sb)
            sb.Append("ORDER BY 1,2 desc")
            Return sb.ToString()

        End Function
        Private Sub WriteOrdersWhereSql(ByVal sb As StringBuilder)
            sb.Append("from Customer cu inner join ")
            sb.Append("Orders od on cu.CustomerId=od.CustomerId inner join ")
            sb.Append("SalesPerson sp on od.salesPersonId = sp.salesPersonId inner join ")
            sb.Append("City ci on ci.cityId = sp.cityId inner join ")
            sb.Append("Country co on ci.countryId = co.countryId ")
        End Sub

    End Module
End Namespace
