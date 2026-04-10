' ***********************************************************************
' Author   : Elektro
' Modified : 09-April-2026
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Command IDs "

Namespace MyPackage.PackageSettings

    ''' <summary>
    ''' Exposes the numeric command identifiers for all commands defined in this package.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' These values must match the command identifiers declared in the <c>.vsct</c> file.
    ''' </remarks>
    Friend Module CommandIds

#Region " Section: Code References "

        ''' <summary>
        ''' Command identifier for the <c>&lt;see cref="..."/&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const CodeRef As UInteger = &H100UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;paramref name="..."/&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const ParamRef As UInteger = CommandIds.CodeRef + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;see langword="..."/&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const LangRef As UInteger = CommandIds.ParamRef + 1UI

#End Region

#Region " Section: Hyperlinks "

        ''' <summary>
        ''' Command identifier for the <c>&lt;see href="..."/&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const HrefLink As UInteger = CommandIds.LangRef + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;seealso href="..."/&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const SeeAlsoLink As UInteger = CommandIds.HrefLink + 1UI

#End Region

#Region " Section: Code Block Formatting "

        ''' <summary>
        ''' Command identifier for the <c>&lt;c&gt;...&lt;/c&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const InlineCode As UInteger = CommandIds.Paragraph + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;code&gt;...&lt;/code&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const MultilineCode As UInteger = CommandIds.InlineCode + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;example&gt;</c> block wrapping operation.
        ''' </summary>
        Friend Const CodeExample As UInteger = CommandIds.MultilineCode + 1UI

#End Region

#Region " Section: Common Formatting "

        ''' <summary>
        ''' Command identifier for the <c>&lt;b&gt;...&lt;/b&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const Bold As UInteger = CommandIds.SeeAlsoLink + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;i&gt;...&lt;/i&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const Italic As UInteger = CommandIds.Bold + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;u&gt;...&lt;/u&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const Underline As UInteger = CommandIds.Italic + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;para&gt;...&lt;/para&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const Paragraph As UInteger = CommandIds.Underline + 1UI

        ''' <summary>
        ''' Command identifier for the <c>&lt;remarks&gt;...&lt;/remarks&gt;</c> wrapping operation.
        ''' </summary>
        Friend Const Remarks As UInteger = CommandIds.CodeExample + 1UI

        ''' <summary>
        ''' Command identifier for the XML documentation separator line insertion operation.
        ''' </summary>
        Friend Const InsertSeparatorLine As UInteger = CommandIds.GenerateSnippet + 1UI

#End Region

#Region " Section: Snippet Generation "

        ''' <summary>
        ''' Command identifier for the <c>.snippet</c> XML structure generation operation.
        ''' </summary>
        Friend Const GenerateSnippet As UInteger = CommandIds.Remarks + 1UI

#End Region

#Region " Section: Editor Operations "

        ''' <summary>
        ''' Command identifier for the operation that collapses all XML documentation comment blocks
        ''' in the current editor.
        ''' </summary>
        Friend Const CollapseXmlComments As UInteger = CommandIds.InsertSeparatorLine + 1UI

        ''' <summary>
        ''' Command identifier for the operation that expands all XML documentation comment blocks
        ''' in the current editor.
        ''' </summary>
        Friend Const ExpandXmlComments As UInteger = CommandIds.CollapseXmlComments + 1UI

        ''' <summary>
        ''' Command identifier for the operation that permanently removes all XML documentation
        ''' comment blocks from the current editor.
        ''' </summary>
        Friend Const DeleteXmlComments As UInteger = CommandIds.ExpandXmlComments + 1UI

#End Region

    End Module

End Namespace

#End Region