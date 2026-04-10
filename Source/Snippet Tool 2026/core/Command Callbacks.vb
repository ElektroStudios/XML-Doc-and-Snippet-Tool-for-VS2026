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

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell

Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.Helpers
Imports Snippet_Tool_2026.MyPackage.UserInterface

#End Region

#Region " Main "

<ProvideMenuResource("Menus.ctmenu", 1)>
<ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)>
<ProvideOptionPage(GetType(SnippetGenerationPageGrid), "Snippet Tool", "Snippet Generation", 0, 0, True)>
<ProvideOptionPage(GetType(OtherOptionsPageGrid), "Snippet Tool", "Other Options", 0, 0, True)>
Partial Public NotInheritable Class Snippet_Tool_2026Package : Inherits AsyncPackage

#Region " Command CallBacks "

#Region " Section: Code References "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsCodeRef"/> command,
    ''' wrapping the selected text as a <c>&lt;see cref="..."/&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsCodeRef_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                      XmlDocOperation.WrapAsCodeRef), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsParamRef"/> command,
    ''' wrapping the selected text as a <c>&lt;paramref name="..."/&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsParamRef_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                      XmlDocOperation.WrapAsParamRef), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsLangRef"/> command,
    ''' wrapping the selected text as a <c>&lt;see langword="..."/&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsLangRef_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsLangRef), {0})
    End Sub

#End Region

#Region " Section: Hyperlinks "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsHrefLink"/> command,
    ''' wrapping the selected text as a <c>&lt;see href="..."/&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsHrefLink_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsHrefLink), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsSeeAlsoLink"/> command,
    ''' wrapping the selected text as a <c>&lt;seealso href="..."/&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsSeeAlsoLink_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsSeeAlsoLink), {0})
    End Sub

#End Region

#Region " Section: Code Block Formatting "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsInlineCode"/> command,
    ''' wrapping the selected text as a <c>&lt;c&gt;...&lt;/c&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsInlineCode_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsInlineCode), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsMultilineCode"/> command,
    ''' wrapping the selected text as a <c>&lt;code&gt;...&lt;/code&gt;</c> block.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsMultilineCode_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsMultilineCode), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsCodeExample"/> command,
    ''' wrapping the selected text inside a full <c>&lt;example&gt;</c> block.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsCodeExample_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsCodeExample), {0})
    End Sub

#End Region

#Region " Section: Common Formatting "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsBold"/> command,
    ''' wrapping the selected text as a <c>&lt;b&gt;...&lt;/b&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsBold_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsBold), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsItalic"/> command,
    ''' wrapping the selected text as an <c>&lt;i&gt;...&lt;/i&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsItalic_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsItalic), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsUnderline"/> command,
    ''' wrapping the selected text as a <c>&lt;u&gt;...&lt;/u&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsUnderline_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsUnderline), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsParagraph"/> command,
    ''' wrapping the selected text as a <c>&lt;para&gt;...&lt;/para&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsParagraph_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsParagraph), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdWrapAsRemarks"/> command,
    ''' wrapping the selected text as a <c>&lt;remarks&gt;...&lt;/remarks&gt;</c> tag.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdWrapAsRemarks_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.WrapAsRemarks), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdInsertSeparatorLine"/> command,
    ''' inserting an XML documentation separator line at the cursor position.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdInsertSeparatorLine_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.InsertSeparatorLine), {0})
    End Sub

#End Region

#Region " Section: Snippet Generation "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdGenerateSnippet"/> command,
    ''' formatting the selected text as a Visual Studio <c>.snippet</c> XML structure.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdGenerateSnippet_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.GenerateSnippet), {0})
    End Sub

#End Region

#Region " Section: Editor Operations "

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdCollapseXmlComments"/> command,
    ''' collapsing all expanded XML documentation comment blocks in the current editor.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdCollapseXmlComments_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.CollapseXmlComments), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdExpandXmlComments"/> command,
    ''' expanding all collapsed XML documentation comment blocks in the current editor.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdExpandXmlComments_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                             XmlDocOperation.ExpandXmlComments), {0})
    End Sub

    ''' <summary>
    ''' Handles the execution of the <see cref="CmdDeleteXmlComments"/> command,
    ''' permanently removing all XML documentation comment blocks from the current editor.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The <see cref="OleMenuCommand"/> that raised the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' An <see cref="EventArgs"/> instance containing no relevant event data.
    ''' </param>
    Private Sub CmdDeleteXmlComments_Callback(sender As Object, e As EventArgs)

        ErrorHandler.ThrowOnFailure(XmlDocOperationExecutor.PerformXmlDocOperation(Snippet_Tool_2026Package.Dte, TextEditorHelper.GetCurrentViewHost,
                                                                                   XmlDocOperation.DeleteXmlComments), {0})
    End Sub

#End Region

#End Region

End Class

#End Region