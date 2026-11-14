Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.IO
Imports System.Xml.Linq

Friend NotInheritable Class TestResources
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Loads a repository fixture copied into the test output directory.
    ''' </summary>
    Friend Shared Function ReadFixture(name As String) As String
        If name Is Nothing Then Throw New ArgumentNullException(NameOf(name))
        Return File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", name))
    End Function

    ''' <summary>
    ''' Loads an XML fixture copied into the test output directory.
    ''' </summary>
    Friend Shared Function ReadXml(name As String) As XDocument
        If name Is Nothing Then Throw New ArgumentNullException(NameOf(name))
        Return XDocument.Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", name))
    End Function
End Class
