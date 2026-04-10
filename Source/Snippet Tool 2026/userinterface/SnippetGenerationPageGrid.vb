' ***********************************************************************
' Author   : Elektro
' Modified : 09-April-2026
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Imports Microsoft.VisualStudio.Shell

#End Region

#Region " Snippet Generation PageGrid "

Namespace MyPackage.UserInterface

    ''' <summary>
    ''' Represents the options page displayed under <b>Tools → Options → Snippet Tool → Snippet Generation</b>,
    ''' exposing settings that control the generated <c>.snippet</c> file output.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/bb166195.aspx"/>
    ''' </remarks>
    <DesignerCategory("Code")>
    <ClassInterface(ClassInterfaceType.AutoDual)>
    <CLSCompliant(False), ComVisible(True)>
    Public NotInheritable Class SnippetGenerationPageGrid : Inherits DialogPage

#Region " Properties "

        ''' <summary>
        ''' Gets or sets the author name embedded in the generated <c>.snippet</c> file header.
        ''' </summary>
        ''' 
        ''' <value>
        ''' A <see cref="String"/> containing the author name. 
        ''' The default value is <see cref="System.String.Empty"/>.
        ''' </value>
        <Category("Fields")>
        <DisplayName("Author")>
        <Description("The author name to embedd in the header of generated .snippet files.")>
        Public Property Author As String = String.Empty

#End Region

    End Class

End Namespace

#End Region