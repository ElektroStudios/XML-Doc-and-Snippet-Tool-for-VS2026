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

Imports Snippet_Tool_2026.MyPackage.Helpers
Imports Snippet_Tool_2026.MyPackage.UserInterface

#End Region

<ProvideMenuResource("Menus.ctmenu", 1)>
<ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)>
<ProvideOptionPage(GetType(SnippetGenerationPageGrid), "Snippet Tool", "Snippet Generation", 0, 0, True)>
<ProvideOptionPage(GetType(OtherOptionsPageGrid), "Snippet Tool", "Other Options", 0, 0, True)>
Partial Public NotInheritable Class Snippet_Tool_2026Package : Inherits AsyncPackage

#Region " Command Event-Handlers "

#Region " Section: Code References "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsCodeRef"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsCodeRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsCodeRef.BeforeQueryStatus

        Logger.Debug("CmdWrapAsCodeRef_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsCodeRef
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock()

        Logger.Debug($"CmdWrapAsCodeRef_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsParamRef"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsParamRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsParamRef.BeforeQueryStatus

        Logger.Debug("CmdWrapAsParamRef_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsParamRef
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock()

        Logger.Debug($"CmdWrapAsParamRef_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsLangRef"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsLangRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsLangRef.BeforeQueryStatus

        Logger.Debug("CmdWrapAsLangRef_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsLangRef
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock()

        Logger.Debug($"CmdWrapAsLangRef_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

#End Region

#Region " Section: Hyperlinks "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsHrefLink"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsHrefLink_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsHrefLink.BeforeQueryStatus

        Logger.Debug("CmdWrapAsHrefLink_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsHrefLink
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock AndAlso
                      TextEditorHelper.SelectedTextIsSingleLine()

        Logger.Debug($"CmdWrapAsHrefLink_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsSeeAlsoLink"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsSeeAlsoLink_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsSeeAlsoLink.BeforeQueryStatus

        Logger.Debug("CmdWrapAsSeeAlsoLink_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsSeeAlsoLink
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock AndAlso
                      TextEditorHelper.SelectedTextIsSingleLine()

        Logger.Debug($"CmdWrapAsSeeAlsoLink_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

#End Region

#Region " Section: Code Block Formatting "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsInlineCode"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsInlineCode_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsInlineCode.BeforeQueryStatus

        Logger.Debug("CmdWrapAsInlineCode_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsInlineCode
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock AndAlso
                      TextEditorHelper.SelectedTextIsSingleLine()

        Logger.Debug($"CmdWrapAsInlineCode_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsMultilineCode"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsMultilineCode_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsMultilineCode.BeforeQueryStatus

        Logger.Debug("CmdWrapAsMultilineCode_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsMultilineCode
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock AndAlso
                      TextEditorHelper.SelectedTextIsMultiLine()

        Logger.Debug($"CmdWrapAsMultilineCode_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsCodeExample"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsCodeExample_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsCodeExample.BeforeQueryStatus

        Logger.Debug("CmdWrapAsCodeExample_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsCodeExample
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsCodeExample_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

#End Region

#Region " Section: Common Formatting "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsBold"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsBold_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsBold.BeforeQueryStatus

        Logger.Debug("CmdWrapAsBold_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsBold
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock() 'AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsBold_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsItalic"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsItalic_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsItalic.BeforeQueryStatus

        Logger.Debug("CmdWrapAsItalic_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsItalic
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock() 'AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsItalic_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsUnderline"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsUnderline_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsUnderline.BeforeQueryStatus

        Logger.Debug("CmdWrapAsUnderline_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsUnderline
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock() 'AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsUnderline_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsParagraph"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsParagraph_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsParagraph.BeforeQueryStatus

        Logger.Debug("CmdWrapAsParagraph_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsParagraph
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock() 'AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsParagraph_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdWrapAsRemarks"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdWrapAsRemarks_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdWrapAsRemarks.BeforeQueryStatus

        Logger.Debug("CmdWrapAsRemarks_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableWrapAsRemarks
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsCaretOnXmlBlock AndAlso
                      TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdWrapAsRemarks_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdInsertSeparatorLine"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdInsertSeparatorLine_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdInsertSeparatorLine.BeforeQueryStatus

        ' This command is always enabled; no status update is required.

        Logger.Debug("CmdInsertSeparatorLine_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableInsertSeparatorLine
        cmd.Enabled = cmd.Visible AndAlso True

        Logger.Debug($"CmdInsertSeparatorLine_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

#End Region

#Region " Section: Snippet Generation "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdGenerateSnippet"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdGenerateSnippet_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdGenerateSnippet.BeforeQueryStatus

        Logger.Debug("CmdGenerateSnippet_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableGenerateSnippet
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.IsTextSelected()

        Logger.Debug($"CmdGenerateSnippet_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

#End Region

#Region " Section: Editor Operations "

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdCollapseXmlComments"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdCollapseXmlComments_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdCollapseXmlComments.BeforeQueryStatus

        Logger.Debug("CmdCollapseXmlComments_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableCollapseXmlComments
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.ActiveDocumentContainsXmlCommentCharsPrefix()

        Logger.Debug($"CmdCollapseXmlComments_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdExpandXmlComments"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdExpandXmlComments_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdExpandXmlComments.BeforeQueryStatus

        Logger.Debug("CmdExpandXmlComments_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableExpandXmlComments
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.ActiveDocumentContainsXmlCommentCharsPrefix()

        Logger.Debug($"CmdExpandXmlComments_BeforeQueryStatus — cmd.Enabled: {cmd.Enabled}")
    End Sub

    ''' <summary>
    ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event 
    ''' of the <see cref="Main.CmdDeleteXmlComments"/> command.
    ''' </summary>
    ''' 
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    Private Sub CmdDeleteXmlComments_BeforeQueryStatus(sender As Object, e As EventArgs) _
    Handles CmdDeleteXmlComments.BeforeQueryStatus

        Logger.Debug("CmdDeleteXmlComments_BeforeQueryStatus — START")

        Dim cmd As OleMenuCommand = DirectCast(sender, OleMenuCommand)
        Dim page As OtherOptionsPageGrid = DirectCast(Me.GetDialogPage(GetType(OtherOptionsPageGrid)), OtherOptionsPageGrid)
        cmd.Visible = page.EnableDeleteXmlComments
        cmd.Enabled = cmd.Visible AndAlso TextEditorHelper.ActiveDocumentContainsXmlCommentCharsPrefix()

        Logger.Debug($"CmdDeleteXmlComments_BeforeQueryStatus — END — cmd.Enabled: {cmd.Enabled}")

    End Sub

#End Region

#End Region

End Class
