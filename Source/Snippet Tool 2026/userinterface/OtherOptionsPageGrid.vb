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

#Region " Other Options PageGrid "

Namespace MyPackage.UserInterface

    ''' <summary>
    ''' Represents the options page displayed under <b>Tools → Options → Snippet Tool → Other Options</b>,
    ''' exposing settings that control the appearance of inserted XML documentation separator lines.
    ''' </summary>
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/bb166195.aspx"/>
    ''' </remarks>
    <DesignerCategory("Code")>
    <ClassInterface(ClassInterfaceType.AutoDual)>
    <CLSCompliant(False), ComVisible(True)>
    Public NotInheritable Class OtherOptionsPageGrid : Inherits DialogPage

#Region " Properties "

#Region " Section: Code References "

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Code Reference</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code References")>
        <DisplayName("Wrap as Code Reference")>
        <Description("When enabled, the 'Wrap as Code Reference' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsCodeRef As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Parameter Reference</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code References")>
        <DisplayName("Wrap as Parameter Reference")>
        <Description("When enabled, the 'Wrap as Parameter Reference' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsParamRef As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Language Word Reference</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code References")>
        <DisplayName("Wrap as Language Word Reference")>
        <Description("When enabled, the 'Wrap as Language Word Reference' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsLangRef As Boolean = True

#End Region

#Region " Section: HyperLinks "

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Href Link</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: HyperLinks")>
        <DisplayName("Wrap as Href Link")>
        <Description("When enabled, the 'Wrap as Href Link' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsHrefLink As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as SeeAlso Link</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: HyperLinks")>
        <DisplayName("Wrap as SeeAlso Link")>
        <Description("When enabled, the 'Wrap as SeeAlso Link' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsSeeAlsoLink As Boolean = True

#End Region

#Region " Section: Code Block Formatting"

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Inline Code</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code Block Formatting")>
        <DisplayName("Wrap as Inline Code")>
        <Description("When enabled, the 'Wrap as Inline Code' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsInlineCode As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Multiline Code</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code Block Formatting")>
        <DisplayName("Wrap as Multiline Code")>
        <Description("When enabled, the 'Wrap as Multiline Code' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsMultilineCode As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Code Example</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Code Block Formatting")>
        <DisplayName("Wrap as Code Example")>
        <Description("When enabled, the 'Wrap as Code Example' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsCodeExample As Boolean = True

#End Region

#Region " Section: Common Formatting "

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Bold</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Wrap as Bold")>
        <Description("When enabled, the 'Wrap as Bold' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsBold As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Italic</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Wrap as Italic")>
        <Description("When enabled, the 'Wrap as Italic' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsItalic As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Underline</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Wrap as Underline")>
        <Description("When enabled, the 'Wrap as Underline' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsUnderline As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Paragraph</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Wrap as Paragraph")>
        <Description("When enabled, the 'Wrap as Paragraph' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsParagraph As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Wrap as Remarks</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Wrap as Remarks")>
        <Description("When enabled, the 'Wrap as Remarks' command appears in the Snippet Tool context menu.")>
        Public Property EnableWrapAsRemarks As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Insert Separator Line</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Common Formatting")>
        <DisplayName("Insert Separator Line")>
        <Description("When enabled, the 'Insert Separator Line' command appears in the Snippet Tool context menu.")>
        Public Property EnableInsertSeparatorLine As Boolean = True

#End Region

#Region " Section: Snippet Generation "

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Generate Snippet File</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Snippet Generation")>
        <DisplayName("Generate Snippet File")>
        <Description("When enabled, the 'Generate Snippet File' command appears in the Snippet Tool context menu.")>
        Public Property EnableGenerateSnippet As Boolean = True

#End Region

#Region " Section: Editor Operations "

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Collapse XML Comment Blocks</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Editor Operations")>
        <DisplayName("Collapse XML Comment Blocks")>
        <Description("When enabled, the 'Collapse XML Comment Blocks' command appears in the Snippet Tool context menu.")>
        Public Property EnableCollapseXmlComments As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Expand XML Comment Blocks</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Editor Operations")>
        <DisplayName("Expand XML Comment Blocks")>
        <Description("When enabled, the 'Expand XML Comment Blocks' command appears in the Snippet Tool context menu.")>
        Public Property EnableExpandXmlComments As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the <c>Delete XML Comment Blocks</c> command
        ''' is visible in the context menu.
        ''' </summary>
        <Category("Enable Commands: Editor Operations")>
        <DisplayName("Delete XML Comment Blocks")>
        <Description("When enabled, the 'Delete XML Comment Blocks' command appears in the Snippet Tool context menu.")>
        Public Property EnableDeleteXmlComments As Boolean = True

#End Region

#Region " Separator Line - Additional Options "

        ''' <summary>
        ''' Gets or sets the character used to fill the XML documentation separator line.
        ''' </summary>
        ''' 
        ''' <value>
        ''' A <see cref="Char"/> representing the fill character. 
        ''' The default value is <c>'-'</c>.
        ''' </value>
        <Category("Separator Line")>
        <DisplayName("Fill Character")>
        <Description("The character used to fill the XML documentation separator line.")>
        Public Property Character As Char = "-"c

        ''' <summary>
        ''' Gets or sets the total length of the XML documentation separator line, in characters.
        ''' </summary>
        ''' 
        ''' <value>
        ''' An <see cref="Integer"/> representing the separator line length. 
        ''' The default value is <c>100</c>.
        ''' </value>
        <Category("Separator Line")>
        <DisplayName("Line Length")>
        <Description("The total length of the XML documentation separator line, in characters.")>
        Public Property Length As Integer = 100

#End Region

#End Region

    End Class

End Namespace

#End Region