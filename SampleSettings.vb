Imports System
Imports System.IO

Namespace EPPlusSamples
    Public Module SampleSettings
        Public ReadOnly Property ConnectionString As String = "Data Source=SampleDb\EPPlusSampleDb.db;Version=3;"
        'Set the output directory to the SampleApp folder where the app is running from. 
        Public ReadOnly Property OutputDir As DirectoryInfo = New DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp")

    End Module
End Namespace
