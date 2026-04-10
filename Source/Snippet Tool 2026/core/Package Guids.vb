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

#End Region

#Region " Guids "

Namespace MyPackage

    ''' <summary>
    ''' Exposes the <see cref="Guid"/> identifiers for the package and all its command sets.
    ''' </summary>
    Friend Module Guids

#Region " Package "

        ''' <summary>
        ''' The unique identifier of this package, represented as a <see cref="String"/>.
        ''' </summary>
        Friend Const PackageGuid As String =
            "D8CCC633-16D2-458B-94BD-E92AD7E311D6"

#End Region

#Region " Command Groups "

        ''' <summary>
        ''' The unique identifier of the command set that groups reference-related commands
        ''' such as <c>&lt;see cref="..."/&gt;</c>, <c>&lt;paramref/&gt;</c> and <c>&lt;see langword="..."/&gt;</c>.
        ''' </summary>
        Friend ReadOnly CmdGroupCodeReferences As New Guid(
            "768A73BC-8F9D-4638-986F-000000000001")

        ''' <summary>
        ''' The unique identifier of the command set that groups hyperlink commands
        ''' such as <c>&lt;see href="..."/&gt;</c> and <c>&lt;seealso href="..."/&gt;</c>.
        ''' </summary>
        Friend ReadOnly CmdGroupHyperlinks As New Guid(
            "768A73BC-8F9D-4638-986F-000000000002")

        ''' <summary>
        ''' The unique identifier of the command set that groups code-formatting commands
        ''' such as <c>&lt;c&gt;</c>, <c>&lt;code&gt;</c> and <c>&lt;example&gt;</c>.
        ''' </summary>
        Friend ReadOnly CmdGroupCodeBlockFormatting As New Guid(
            "768A73BC-8F9D-4638-986F-000000000003")

        ''' <summary>
        ''' The unique identifier of the command set that groups miscellaneous formatting commands
        ''' such as <c>&lt;b&gt;</c>, <c>&lt;i&gt;</c>, <c>&lt;u&gt;</c>, <c>&lt;para&gt;</c> and <c>&lt;remarks&gt;</c>.
        ''' </summary>
        Friend ReadOnly CmdGroupCommonFormatting As New Guid(
            "768A73BC-8F9D-4638-986F-000000000004")

        ''' <summary>
        ''' The unique identifier of the command set that groups snippet-related commands
        ''' such as the <c>.snippet</c> XML structure generation operation.
        ''' </summary>
        Friend ReadOnly CmdGroupSnippetGeneration As New Guid(
            "768A73BC-8F9D-4638-986F-000000000005")

        ''' <summary>
        ''' The unique identifier of the command set that groups editor-level operations
        ''' such as collapsing, expanding and deleting XML documentation comment blocks.
        ''' </summary>
        Friend ReadOnly CmdGroupEditorOperations As New Guid(
            "768A73BC-8F9D-4638-986F-000000000006")

#End Region

    End Module

End Namespace

#End Region