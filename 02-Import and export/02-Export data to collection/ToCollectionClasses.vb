Imports OfficeOpenXml.Attributes
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel

Namespace EPPlusSamples
    Public Class ToCollectionSamplePerson
        Public Sub New()

        End Sub

        Public Sub New(ByVal firstName As String, ByVal lastName As String, ByVal height As Integer, ByVal birthDate As Date)
            Me.FirstName = firstName
            Me.LastName = lastName
            Me.Height = height
            Me.BirthDate = birthDate
        End Sub

        Public Property FirstName As String

        Public Property LastName As String

        Public Property Height As Integer

        Public Property BirthDate As Date
    End Class

    Public Class ToCollectionSamplePersonAttr
        Public Sub New()

        End Sub

        Public Sub New(ByVal firstName As String, ByVal lastName As String, ByVal height As Integer, ByVal birthDate As Date)
            Me.FirstName = firstName
            Me.LastName = lastName
            Me.Height = height
            Me.BirthDate = birthDate
        End Sub

        <DisplayName("The persons first name")>
        Public Property FirstName As String

        <Description("The persons last name")>
        Public Property LastName As String

        <EpplusTableColumn(Header:="Height of the person")>
        Public Property Height As Integer

        Public Property BirthDate As Date
    End Class

    Public Module ToCollectionSampleData
        Public ReadOnly Property Persons As IEnumerable(Of ToCollectionSamplePerson)
            Get
                Return New List(Of ToCollectionSamplePerson) From {
                    New ToCollectionSamplePerson("John", "Doe", 176, New DateTime(1978, 3, 15)),
                    New ToCollectionSamplePerson("Sven", "Svensson", 183, New DateTime(1995, 11, 3)),
                    New ToCollectionSamplePerson("Jane", "Doe", 168, New DateTime(1989, 2, 26))
                }
            End Get
        End Property

        Public ReadOnly Property PersonsWithAttributes As IEnumerable(Of ToCollectionSamplePersonAttr)
            Get
                Return New List(Of ToCollectionSamplePersonAttr) From {
                    New ToCollectionSamplePersonAttr("John", "Doe", 176, New DateTime(1978, 3, 15)),
                    New ToCollectionSamplePersonAttr("Sven", "Svensson", 183, New DateTime(1995, 11, 3)),
                    New ToCollectionSamplePersonAttr("Jane", "Doe", 168, New DateTime(1989, 2, 26))
                }
            End Get
        End Property
    End Module
End Namespace
