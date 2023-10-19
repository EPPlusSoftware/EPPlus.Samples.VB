Imports OfficeOpenXml
Imports OfficeOpenXml.Attributes
Imports OfficeOpenXml.Table
Imports System
Imports System.Collections.Generic

Namespace EPPlusSamples.LoadingData
    <EpplusTable(TableStyle:=TableStyles.Dark1, PrintHeaders:=True, AutofitColumns:=True, AutoCalculate:=False, ShowTotal:=True, ShowFirstColumn:=True)>
    <EpplusFormulaTableColumn(Order:=6, NumberFormat:="€#,##0.00", Header:="Tax amount", FormulaR1C1:="RC[-2] * RC[-1]", TotalsRowFunction:=RowFunctions.Sum, TotalsRowNumberFormat:="€#,##0.00"), EpplusFormulaTableColumn(Order:=7, NumberFormat:="€#,##0.00", Header:="Net salary", Formula:="E2-G2", TotalsRowFunction:=RowFunctions.Sum, TotalsRowNumberFormat:="€#,##0.00")>
    Friend Class Actor
        <EpplusIgnore>
        Public Property Id As Integer

        <EpplusTableColumn(Order:=3)>
        Public Property LastName As String
        <EpplusTableColumn(Order:=1, Header:="First name")>
        Public Property FirstName As String
        <EpplusTableColumn(Order:=2)>
        Public Property MiddleName As String

        <EpplusTableColumn(Order:=0, NumberFormat:="yyyy-MM-dd", TotalsRowLabel:="Total")>
        Public Property Birthdate As Date

        <EpplusTableColumn(Order:=4, NumberFormat:="€#,##0.00", TotalsRowFunction:=RowFunctions.Sum, TotalsRowNumberFormat:="€#,##0.00")>
        Public Property Salary As Double

        <EpplusTableColumn(Order:=5, NumberFormat:="0%", TotalsRowFormula:="Table1[[#Totals],[Tax amount]]/Table1[[#Totals],[Salary]]", TotalsRowNumberFormat:="0 %")>
        Public Property Tax As Double
    End Class

    <EpplusTable(TableStyle:=TableStyles.Medium1, PrintHeaders:=True, AutofitColumns:=True, AutoCalculate:=True, ShowLastColumn:=True)>
    Friend Class Actor2
        Inherits Actor

    End Class

    ' classes used to demonstrate this functionality with a complex type property
    <EpplusTable(TableStyle:=TableStyles.Light14, PrintHeaders:=True, AutofitColumns:=True, AutoCalculate:=True, ShowLastColumn:=True)>
    Friend Class Actor3
        <EpplusIgnore>
        Public Property Id As Integer

        <EpplusNestedTableColumn(Order:=1)>
        Public Property Name As ActorName

        <EpplusTableColumn(Order:=0, NumberFormat:="yyyy-MM-dd", TotalsRowLabel:="Total")>
        Public Property Birthdate As Date

        <EpplusTableColumn(Order:=2, NumberFormat:="€#,##0.00", TotalsRowFunction:=RowFunctions.Sum, TotalsRowNumberFormat:="€#,##0.00")>
        Public Property Salary As Double

        <EpplusTableColumn(Order:=3, NumberFormat:="0%", TotalsRowFormula:="Table1[[#Totals],[Tax amount]]/Table1[[#Totals],[Salary]]", TotalsRowNumberFormat:="0 %")>
        Public Property Tax As Double
    End Class

    Friend Class ActorName
        <EpplusTableColumn(Order:=3)>
        Public Property LastName As String
        <EpplusTableColumn(Order:=1, Header:="First name")>
        Public Property FirstName As String
        <EpplusTableColumn(Order:=2)>
        Public Property MiddleName As String
    End Class

    Public Module LoadingDataFromCollectionWithAttributes
        Public Sub Run()
            ' sample data
            Dim actors = New List(Of Actor) From {
    New Actor With {
        .Salary = 256.24,
        .Tax = 0.21,
        .FirstName = "John",
        .MiddleName = "Bernhard",
        .LastName = "Doe",
        .Birthdate = New DateTime(1950, 3, 15)
    },
    New Actor With {
        .Salary = 278.55,
        .Tax = 0.23,
        .FirstName = "Sven",
        .MiddleName = "Bertil",
        .LastName = "Svensson",
        .Birthdate = New DateTime(1962, 6, 10)
    },
    New Actor With {
        .Salary = 315.34,
        .Tax = 0.28,
        .FirstName = "Lisa",
        .MiddleName = "Maria",
        .LastName = "Gonzales",
        .Birthdate = New DateTime(1971, 10, 2)
    }
}

            Dim subclassActors = New List(Of Actor2) From {
    New Actor2 With {
        .Salary = 256.24,
        .Tax = 0.21,
        .FirstName = "John",
        .MiddleName = "Bernhard",
        .LastName = "Doe",
        .Birthdate = New DateTime(1950, 3, 15)
    },
    New Actor2 With {
        .Salary = 278.55,
        .Tax = 0.23,
        .FirstName = "Sven",
        .MiddleName = "Bertil",
        .LastName = "Svensson",
        .Birthdate = New DateTime(1962, 6, 10)
    },
    New Actor2 With {
        .Salary = 315.34,
        .Tax = 0.28,
        .FirstName = "Lisa",
        .MiddleName = "Maria",
        .LastName = "Gonzales",
        .Birthdate = New DateTime(1971, 10, 2)
    }
}

            Dim complexTypeActors = New List(Of Actor3) From {
    New Actor3 With {
        .Salary = 256.24,
        .Tax = 0.21,
        .Name = New ActorName With {
            .FirstName = "John",
            .MiddleName = "Bernhard",
            .LastName = "Doe"
        },
        .Birthdate = New DateTime(1950, 3, 15)
    },
    New Actor3 With {
        .Salary = 278.55,
        .Tax = 0.23,
        .Name = New ActorName With {
            .FirstName = "Sven",
            .MiddleName = "Bertil",
            .LastName = "Svensson"
        },
        .Birthdate = New DateTime(1962, 6, 10)
    },
    New Actor3 With {
        .Salary = 315.34,
        .Tax = 0.28,
        .Name = New ActorName With {
            .FirstName = "Lisa",
            .MiddleName = "Maria",
            .LastName = "Gonzales"
        },
        .Birthdate = New DateTime(1971, 10, 2)
    }
}

            Using package = New ExcelPackage(FileUtil.GetCleanFileInfo("2.1-LoadFromCollectionAttributes.xlsx"))
                ' using the Actor class above
                Dim sheet = package.Workbook.Worksheets.Add("Actors")
                sheet.Cells("A1").LoadFromCollection(actors)

                ' using a subclass where we have overridden the EpplusTableAttribute (different TableStyle and highlight last column instead of the first).
                Dim subclassSheet = package.Workbook.Worksheets.Add("Using subclass with attributes")
                subclassSheet.Cells("A1").LoadFromCollection(subclassActors)

                ' using a subclass where we have overridden the EpplusTableAttribute (different TableStyle and highlight last column instead of the first).
                Dim complexTypePropertySheet = package.Workbook.Worksheets.Add("Complex type property")
                complexTypePropertySheet.Cells("A1").LoadFromCollection(complexTypeActors)

                package.Save()
            End Using
        End Sub
    End Module
End Namespace
