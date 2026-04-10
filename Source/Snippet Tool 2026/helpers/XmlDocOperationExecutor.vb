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
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading

Imports EnvDTE

Imports EnvDTE80

Imports Microsoft.VisualBasic

Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Formatting

Imports Snippet_Tool_2026.MyPackage.Core

#End Region

Namespace MyPackage.Helpers

    Public NotInheritable Class XmlDocOperationExecutor

#Region " Public Methods "

        ''' <summary>
        ''' Performs the specified XML documentation operation on the currently selected text
        ''' in the active code editor window.
        ''' </summary>
        ''' 
        ''' <param name="dte">
        ''' The <see cref="DTE2"/> automation model instance.
        ''' </param>
        ''' 
        ''' <param name="viewHost">
        ''' The <see cref="IWpfTextViewHost"/> of the active editor window.
        ''' </param>
        ''' 
        ''' <param name="operation">
        ''' The <see cref="XmlDocOperation"/> to perform on the selected text.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <c>0</c> if the operation completed successfully;
        ''' <c>-1</c> if the operation failed, the document language is unsupported,
        ''' or an unhandled exception occurred.
        ''' </returns>
        Public Shared Function PerformXmlDocOperation(dte As DTE2, viewHost As IWpfTextViewHost, operation As XmlDocOperation) As Integer

            Logger.Debug($"PerformXmlDocOperation — START — operation: {operation}")

            Try
                Dim textView As IWpfTextView = viewHost.TextView
                Dim selection As ITextSelection = textView.Selection
                Dim lang As String = TextEditorHelper.GetActiveDocumentLanguage(dte)

                Logger.Debug($"PerformXmlDocOperation — detected language: '{lang}'")

                Dim codeExampleLang As String = ""
                Dim xmlCommentChars As String = ""
                Select Case lang

                    Case CodeModelLanguageConstants.vsCMLanguageVB
                        xmlCommentChars = DocsConstants.XmlCommentCharsVB
                        codeExampleLang = "VB"
                        Logger.Debug($"PerformXmlDocOperation — VB document — xmlCommentChars: '{xmlCommentChars}', codeExampleLang: '{codeExampleLang}'")

                    Case CodeModelLanguageConstants.vsCMLanguageCSharp
                        xmlCommentChars = DocsConstants.XmlCommentCharsCS
                        codeExampleLang = "CSharp"
                        Logger.Debug($"PerformXmlDocOperation — CSharp document — xmlCommentChars: '{xmlCommentChars}', codeExampleLang: '{codeExampleLang}'")

                    Case ""
                        ' Any document is open or else document without specific language.
                        Logger.Warn("PerformXmlDocOperation — empty language string, no specific language detected, returning -1")
                        Return -1

                    Case Else ' VC++
                        Logger.Warn($"PerformXmlDocOperation — unsupported language '{lang}' (VC++?), returning -1")
                        Return -1

                End Select

                ' Construct the replacement string.
                Logger.Debug($"PerformXmlDocOperation — dispatching operation: {operation}")
                Select Case operation

#Region " Section: Code References "

                    Case XmlDocOperation.WrapAsCodeRef
                        XmlDocOperationExecutor.Perform_WrapAsCodeRef(selection)

                    Case XmlDocOperation.WrapAsParamRef
                        XmlDocOperationExecutor.Perform_WrapAsParamRef(selection)

                    Case XmlDocOperation.WrapAsLangRef
                        XmlDocOperationExecutor.Perform_WrapAsLangRef(selection)

#End Region

#Region " Section: Hyperlinks "

                    Case XmlDocOperation.WrapAsHrefLink
                        XmlDocOperationExecutor.Perform_WrapAsHrefLink(selection)

                    Case XmlDocOperation.WrapAsSeeAlsoLink
                        XmlDocOperationExecutor.Perform_WrapAsSeeAlsoLink(selection)

#End Region

#Region " Section: Code Block Formatting "

                    Case XmlDocOperation.WrapAsInlineCode
                        XmlDocOperationExecutor.Perform_WrapAsInlineCode(selection)

                    Case XmlDocOperation.WrapAsMultilineCode
                        XmlDocOperationExecutor.Perform_WrapAsMultilineCode(selection)

                    Case XmlDocOperation.WrapAsCodeExample
                        XmlDocOperationExecutor.Perform_WrapAsCodeExample(selection, xmlCommentChars, codeExampleLang)

#End Region

#Region " Section: Common Formatting "

                    Case XmlDocOperation.WrapAsBold
                        XmlDocOperationExecutor.Perform_WrapAsBold(selection)

                    Case XmlDocOperation.WrapAsItalic
                        XmlDocOperationExecutor.Perform_WrapAsItalic(selection)

                    Case XmlDocOperation.WrapAsUnderline
                        XmlDocOperationExecutor.Perform_WrapAsUnderline(selection)

                    Case XmlDocOperation.WrapAsParagraph
                        XmlDocOperationExecutor.Perform_WrapAsParagraph(selection, xmlCommentChars)

                    Case XmlDocOperation.WrapAsRemarks
                        XmlDocOperationExecutor.Perform_WrapAsRemarks(selection, xmlCommentChars)

                    Case XmlDocOperation.InsertSeparatorLine
                        XmlDocOperationExecutor.Perform_InsertSeparatorLine(selection, xmlCommentChars)

#End Region

#Region " Section: Snippet Generation "

                    Case XmlDocOperation.GenerateSnippet
                        XmlDocOperationExecutor.Perform_GenerateSnippet(selection, lang)

#End Region

#Region " Section: Editor Operations "

                    Case XmlDocOperation.CollapseXmlComments
                        XmlDocOperationExecutor.Perform_CollapseXmlComments()

                    Case XmlDocOperation.ExpandXmlComments
                        XmlDocOperationExecutor.Perform_ExpandXmlComments()

                    Case XmlDocOperation.DeleteXmlComments
                        XmlDocOperationExecutor.Perform_DeleteXmlComments()

#End Region

                    Case Else
                        Logger.Warn($"PerformXmlDocOperation — unrecognized operation '{operation}', returning -1")
                        Return -1

                End Select

                Logger.Info($"PerformXmlDocOperation — END — operation {operation} completed successfully")
                Return 0

            Catch ex As Exception
                Logger.Error($"PerformXmlDocOperation — unhandled exception for operation '{operation}'", ex)
                Interaction.MsgBox(ex.Message, MsgBoxStyle.Critical, "Snippet Tool")
                Return -1

            End Try

        End Function

#End Region

#Region " Private Methods "

#Region " Section: Code References "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsCodeRef"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsCodeRef(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsCodeRef — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsCodeRef — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsCodeRef — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsCodeRef — span length: {span.Length} chars, text: '{span.GetText()}'")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see cref=""{0}""/>", span.GetText()))
            Logger.Info($"Perform_WrapAsCodeRef — selection wrapped in <see cref> tag")

            Logger.Debug("Perform_WrapAsCodeRef — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsParamRef"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsParamRef(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsParamRef — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsParamRef — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsParamRef — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsParamRef — span length: {span.Length} chars, text: '{span.GetText()}'")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<paramref name=""{0}""/>", span.GetText()))
            Logger.Info($"Perform_WrapAsParamRef — selection wrapped in <paramref name> tag")

            Logger.Debug("Perform_WrapAsParamRef — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsLangRef"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsLangRef(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsLangRef — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsLangRef — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsLangRef — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsLangRef — span length: {span.Length} chars, text: '{span.GetText()}'")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see langword=""{0}""/>", span.GetText()))
            Logger.Info($"Perform_WrapAsLangRef — selection wrapped in <see langword> tag")

            Logger.Debug("Perform_WrapAsLangRef — END")

        End Sub

#End Region

#Region " Section: Hyperlinks "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsHrefLink"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsHrefLink(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsHrefLink — START")

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsHrefLink — span length: {span.Length} chars, text: '{span.GetText()}'")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see href=""{0}""/>", span.GetText()))
            Logger.Info("Perform_WrapAsHrefLink — selection wrapped in <see href> tag")

            Logger.Debug("Perform_WrapAsHrefLink — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsSeeAlsoLink"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsSeeAlsoLink(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsSeeAlsoLink — START")

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsSeeAlsoLink — span length: {span.Length} chars, text: '{span.GetText()}'")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<seealso href=""{0}""/>", span.GetText()))
            Logger.Info($"Perform_WrapAsSeeAlsoLink — selection wrapped in <seealso> tag")

            Logger.Debug("Perform_WrapAsSeeAlsoLink — END")

        End Sub

#End Region

#Region " Section: Code Block Formatting "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsInlineCode"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsInlineCode(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsInlineCode — START")

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsInlineCode — span length: {span.Length} chars")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<c>{0}</c>", span.GetText()))
            Logger.Info("Perform_WrapAsInlineCode — selection wrapped in <c> tags")

            Logger.Debug("Perform_WrapAsInlineCode — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsMultilineCode"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsMultilineCode(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsMultilineCode — START")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Make multiline code")
                Logger.Debug("Perform_WrapAsMultilineCode — undo context opened")

                ' Create a SnapshotSpan for all text to be replaced.
                Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
                Logger.Debug($"Perform_WrapAsMultilineCode — span length: {span.Length} chars")

                ' Perform the replacement.
                span.Snapshot.TextBuffer.Replace(span, String.Format("<code>{0}</code>", span.GetText()))
                Logger.Info("Perform_WrapAsMultilineCode — selection wrapped in <code> tags")

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Debug("Perform_WrapAsMultilineCode — END — undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_WrapAsMultilineCode — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsCodeExample"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="xmlCommentChars">
        ''' The XML comment chars to use.
        ''' </param>
        Private Shared Sub Perform_WrapAsCodeExample(selection As ITextSelection, xmlCommentChars As String, codeLang As String)

            Logger.Debug($"Perform_WrapAsCodeExample — START — xmlCommentChars: '{xmlCommentChars}', codeLang: '{codeLang}'")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Make code example")
                Logger.Debug("Perform_WrapAsCodeExample — undo context opened")

                ' Get the start and end points of the selection.
                Dim start As VirtualSnapshotPoint = selection.Start
                Dim [end] As VirtualSnapshotPoint = selection.End
                Logger.Debug($"Perform_WrapAsCodeExample — selection Start: {start.Position}, End: {[end].Position}")

                ' Get the lines that contain the start and end points.
                Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                ' Get the start and end points of the lines.
                Dim startLinePoint As SnapshotPoint = startLine.Start
                Dim endLinePoint As SnapshotPoint = endLine.End

                ' Create a SnapshotSpan for all text to be replaced.
                Dim span As New SnapshotSpan(startLinePoint, endLinePoint)
                Logger.Debug($"Perform_WrapAsCodeExample — span length: {span.Length} chars")

                ' Compute margin.
                Dim lines As IEnumerable(Of String) =
            span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None) '.SkipWhile(Function(line As String) String.IsNullOrWhiteSpace(line))

                Logger.Debug($"Perform_WrapAsCodeExample — selection spans {lines.Count} line(s)")

                Dim margin As Integer =
                    lines.Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                          Select(Function(line)
                                     Dim count As Integer = 0
                                     While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                                     End While
                                     Return Interlocked.Decrement(count)
                                 End Function).Min

                Logger.Debug($"Perform_WrapAsCodeExample — computed margin: {margin}")

                Dim sb As New StringBuilder

                With sb
                    .AppendLine(String.Format("{0} <example> This is a code example.", xmlCommentChars))
                    .AppendLine(String.Format("{0} <code language=""{1}"">", xmlCommentChars, codeLang))

                    For Each line As String In lines

                        If String.IsNullOrWhiteSpace(line) Then
                            sb.AppendLine(New String(" "c, margin) & String.Format("{0} ", xmlCommentChars))
                        Else
                            sb.AppendLine(String.Format("{0}{1}{2}{3}", New String(" "c, margin), xmlCommentChars, If(margin = 0, " ", ""), line.Remove(0, margin)))
                        End If

                    Next line

                    .AppendLine(String.Format("{0} </code>", xmlCommentChars))
                    .AppendLine(String.Format("{0} </example>", xmlCommentChars))
                End With

                Logger.Debug($"Perform_WrapAsCodeExample — built code example block, total length: {sb.Length} chars")

                ' Perform the replacement.
                span.Snapshot.TextBuffer.Replace(span, sb.ToString)
                Logger.Info("Perform_WrapAsCodeExample — selection wrapped in code example block")

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Debug("Perform_WrapAsCodeExample — END — undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_WrapAsCodeExample — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

#End Region

#Region " Section: Common Formatting "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsBold"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsBold(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsBold — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsBold — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsBold — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsBold — span length: {span.Length} chars")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<b>{0}</b>", span.GetText()))
            Logger.Info("Perform_WrapAsBold — selection wrapped in <b> tags")

            Logger.Debug("Perform_WrapAsBold — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsItalic"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsItalic(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsItalic — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsItalic — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsItalic — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsItalic — span length: {span.Length} chars")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<i>{0}</i>", span.GetText()))
            Logger.Info("Perform_WrapAsItalic — selection wrapped in <i> tags")

            Logger.Debug("Perform_WrapAsItalic — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsUnderline"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsUnderline(selection As ITextSelection)

            Logger.Debug("Perform_WrapAsUnderline — START")

            If Not TextEditorHelper.IsTextSelected() Then
                Logger.Debug("Perform_WrapAsUnderline — no selection, selecting current word via Edit.SelectCurrentWord")
                Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.SelectCurrentWord")
                Logger.Debug("Perform_WrapAsUnderline — current word selected")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
            Logger.Debug($"Perform_WrapAsUnderline — span length: {span.Length} chars")

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<u>{0}</u>", span.GetText()))
            Logger.Info("Perform_WrapAsUnderline — selection wrapped in <u> tags")

            Logger.Debug("Perform_WrapAsUnderline — END")

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsParagraph"/> operation on the selected text.
        ''' <para></para>
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_WrapAsParagraph(selection As ITextSelection, xmlCommentChars As String)

            Logger.Debug($"Perform_WrapAsParagraph — START — xmlCommentChars: '{xmlCommentChars}'")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Make paragraph")
                Logger.Debug("Perform_WrapAsParagraph — undo context opened")

                If TextEditorHelper.IsTextSelected() Then
                    Logger.Debug("Perform_WrapAsParagraph — text selected, wrapping selection in <para> tags")

                    ' Create a SnapshotSpan for all text to be replaced.
                    Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan
                    Logger.Debug($"Perform_WrapAsParagraph — span length: {span.Length} chars")

                    ' Perform the replacement.
                    span.Snapshot.TextBuffer.Replace(span, String.Format("<para>{0}</para>", span.GetText()))
                    Logger.Info("Perform_WrapAsParagraph — selection wrapped in <para> tags")

                Else
                    Logger.Debug("Perform_WrapAsParagraph — no selection, inserting empty <para> tag after caret line")

                    Dim caret As ITextCaret = TextEditorHelper.GetCurrentViewHost.TextView.Caret
                    Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                    Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                    Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                    Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                    Logger.Debug($"Perform_WrapAsParagraph — caret line margin: {margin}, line length: {line.Length}")

                    snapshot.TextBuffer.Replace(span, line & ControlChars.NewLine & New String(" "c, margin) &
                                        String.Format("{0} {1}", xmlCommentChars, "<para></para>"))

                    Logger.Info("Perform_WrapAsParagraph — empty <para> tag inserted after caret line")

                End If

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Debug("Perform_WrapAsParagraph — END — undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_WrapAsParagraph — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.WrapAsRemarks"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="xmlCommentChars">
        ''' The XML comment chars to use.
        ''' </param>
        Private Shared Sub Perform_WrapAsRemarks(selection As ITextSelection, xmlCommentChars As String)

            Logger.Debug($"Perform_WrapAsRemarks — START — xmlCommentChars: '{xmlCommentChars}'")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Make remarks")
                Logger.Debug("Perform_WrapAsRemarks — undo context opened")

                Select Case TextEditorHelper.IsTextSelected()

                    Case False
                        Logger.Debug("Perform_WrapAsRemarks — no selection, inserting remarks tag after caret line")

                        Dim caret As ITextCaret = TextEditorHelper.GetCurrentViewHost.TextView.Caret
                        Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                        Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                        Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                        Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                        Logger.Debug($"Perform_WrapAsRemarks — caret line margin: {margin}, line length: {line.Length}")

                        snapshot.TextBuffer.Replace(span, line &
                                                  ControlChars.NewLine &
                                                  New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, "<remarks></remarks>"))

                        Logger.Info("Perform_WrapAsRemarks — empty remarks tag inserted after caret line")

                    Case Else
                        Logger.Debug("Perform_WrapAsRemarks — text selected, wrapping selection in remarks block")

                        ' Create a SnapshotSpan for all text to be replaced.
                        Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

                        Dim lines As IEnumerable(Of String) =
                    span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

                        Logger.Debug($"Perform_WrapAsRemarks — selection spans {lines.Count} line(s)")

                        ' Compute margin.
                        Dim margin As Integer = lines.
                               Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                               Select(Function(line)
                                          Dim count As Integer = 0
                                          While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                                          End While
                                          Return Interlocked.Decrement(count)
                                      End Function).Min

                        Logger.Debug($"Perform_WrapAsRemarks — computed margin: {margin}")

                        Dim sb As New StringBuilder
                        With sb
                            .AppendLine(String.Format("{0} <remarks>", xmlCommentChars))

                            For Each line As String In lines

                                If String.IsNullOrWhiteSpace(line) Then
                                    sb.AppendLine(New String(" "c, margin) & String.Format("{0} ", xmlCommentChars))
                                Else
                                    sb.AppendLine(String.Format("{0}{1}{2}{3}", New String(" "c, margin), xmlCommentChars, If(margin = 0, " ", ""), line.Remove(0, margin)))
                                End If

                            Next line

                            .AppendLine(String.Format("{0} </remarks>", xmlCommentChars))
                        End With

                        Logger.Debug($"Perform_WrapAsRemarks — built remarks block, total length: {sb.Length} chars")

                        ' Perform the replacement.
                        span.Snapshot.TextBuffer.Replace(span, sb.ToString)
                        Logger.Info("Perform_WrapAsRemarks — selection wrapped in remarks block")

                End Select

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Debug("Perform_WrapAsRemarks — END — undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_WrapAsRemarks — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.InsertSeparatorLine"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        Private Shared Sub Perform_InsertSeparatorLine(selection As ITextSelection, xmlCommentChars As String)

            Logger.Debug($"Perform_InsertSeparatorLine — START — xmlCommentChars: '{xmlCommentChars}'")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Make separator line")
                Logger.Debug("Perform_InsertSeparatorLine — undo context opened")

                Dim c As Char =
            Convert.ToChar(Snippet_Tool_2026Package.Dte.Properties("Snippet Tool", "Other Options").Item("Character").Value)

                Dim length As Integer =
           CInt(Snippet_Tool_2026Package.Dte.Properties("Snippet Tool", "Other Options").Item("Length").Value)

                Logger.Debug($"Perform_InsertSeparatorLine — separator char: '{c}', length: {length}")

                Select Case TextEditorHelper.IsTextSelected()

                    Case False
                        Logger.Debug("Perform_InsertSeparatorLine — no selection, inserting separator after caret line")

                        Dim caret As ITextCaret = TextEditorHelper.GetCurrentViewHost.TextView.Caret
                        Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                        Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                        Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                        Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                        Logger.Debug($"Perform_InsertSeparatorLine — caret line margin: {margin}, line length: {line.Length}")

                        snapshot.TextBuffer.Replace(span, line &
                                                  ControlChars.NewLine &
                                                  New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)))

                        Logger.Info("Perform_InsertSeparatorLine — separator line inserted after caret line")

                    Case Else
                        Logger.Debug("Perform_InsertSeparatorLine — text selected, wrapping selection with separator lines")

                        ' Get the start and end points of the selection.
                        Dim start As VirtualSnapshotPoint = selection.Start
                        Dim [end] As VirtualSnapshotPoint = selection.End
                        Logger.Debug($"Perform_InsertSeparatorLine — selection Start: {start.Position}, End: {[end].Position}")

                        ' Get the lines that contain the start and end points.
                        Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                        Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                        ' Get the start and end points of the lines.
                        Dim startLinePoint As SnapshotPoint = startLine.Start
                        Dim endLinePoint As SnapshotPoint = endLine.End

                        ' Create a SnapshotSpan for all text to be replaced.
                        Dim span As New SnapshotSpan(startLinePoint, endLinePoint)

                        Dim lines As IEnumerable(Of String) = span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

                        Logger.Debug($"Perform_InsertSeparatorLine — selection spans {lines.Count} line(s)")

                        ' Compute margin.
                        Dim margin As Integer = lines.
                               Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                               Select(Function(line)
                                          Dim count As Integer = 0
                                          While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                                          End While
                                          Return Interlocked.Decrement(count)
                                      End Function).Min

                        Logger.Debug($"Perform_InsertSeparatorLine — computed margin: {margin}")

                        ' Perform the replacement.
                        span.Snapshot.TextBuffer.Replace(span, New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)) &
                                                       ControlChars.NewLine &
                                                       span.GetText &
                                                       ControlChars.NewLine &
                                                       New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)))

                        Logger.Info("Perform_InsertSeparatorLine — selection wrapped with separator lines")

                End Select

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Debug("Perform_InsertSeparatorLine — END — undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_InsertSeparatorLine — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

#End Region

#Region " Section: Snippet Generation "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.GenerateSnippet"/> operation on the selected text.
        ''' </summary>
        ''' 
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="lang">
        ''' The current active document language.
        ''' </param>
        Private Shared Sub Perform_GenerateSnippet(selection As ITextSelection, lang As String)

            Logger.Debug($"Perform_GenerateSnippet — START — lang: '{lang}'")

            ' Get the start and end points of the selection.
            Dim start As VirtualSnapshotPoint = selection.Start
            Dim [end] As VirtualSnapshotPoint = selection.End
            Logger.Debug($"Perform_GenerateSnippet — selection Start: {start.Position}, End: {[end].Position}")

            ' Get the lines that contain the start and end points.
            Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
            Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

            ' Get the start and end points of the lines.
            Dim startLinePoint As SnapshotPoint = startLine.Start
            Dim endLinePoint As SnapshotPoint = endLine.End

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As New SnapshotSpan(startLinePoint, endLinePoint)
            Logger.Debug($"Perform_GenerateSnippet — span length: {span.Length} chars")

            Dim author As String = Snippet_Tool_2026Package.Dte.Properties("Snippet Tool", "Snippet Generation").Item("Author").Value.ToString
            Logger.Debug($"Perform_GenerateSnippet — author: '{author}'")

            Dim sb As New StringBuilder
            Select Case lang

                Case CodeModelLanguageConstants.vsCMLanguageVB
                    Logger.Debug("Perform_GenerateSnippet — applying VB snippet template")
                    sb.AppendFormat(DocsConstants.SnippetTemplateFormatVB, author, span.GetText()).Replace("$cdataend$", ">")

                Case CodeModelLanguageConstants.vsCMLanguageCSharp
                    Logger.Debug("Perform_GenerateSnippet — applying CS snippet template")
                    sb.AppendFormat(DocsConstants.SnippetTemplateFormatCS, author, span.GetText()).Replace("$cdataend$", ">")

                Case Else
                    Logger.Warn($"Perform_GenerateSnippet — unsupported language '{lang}', snippet template not applied")

            End Select

            If sb.Length = 0 Then
                Logger.Warn("Perform_GenerateSnippet — snippet content is empty, aborting file write")
                Return
            End If

            Dim tempFileName As String = String.Format("{0}.snippet", Path.GetTempFileName)
            Logger.Debug($"Perform_GenerateSnippet — writing snippet to temp file: '{tempFileName}'")

            File.WriteAllText(tempFileName, sb.ToString, Encoding.Default)
            Logger.Info($"Perform_GenerateSnippet — snippet file written ({sb.Length} chars), opening with default handler")

            Diagnostics.Process.Start(tempFileName)
            Logger.Debug("Perform_GenerateSnippet — END — process started for temp file")

        End Sub

#End Region

#Region " Section: Editor Operations "

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.CollapseXmlComments"/> operation.
        ''' </summary>
        Private Shared Sub Perform_CollapseXmlComments()

            Logger.Debug("Perform_CollapseXmlComments — START")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Collapse XML comments")
                Logger.Debug("Perform_CollapseXmlComments — undo context opened")

                Dim elements As CodeElements = Snippet_Tool_2026Package.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                Logger.Debug($"Perform_CollapseXmlComments — processing {elements.Count} top-level code element(s)")

                For Each ce As CodeElement2 In elements
                    Logger.Debug($"Perform_CollapseXmlComments — collapsing element '{ce.Name}', Kind: {ce.Kind}")
                    TextEditorHelper.CollapseSubmembers(ce, False)
                    Logger.Debug($"Perform_CollapseXmlComments — collapse done for '{ce.Name}'")
                Next

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Info("Perform_CollapseXmlComments — END — all comments collapsed, undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_CollapseXmlComments — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.ExpandXmlComments"/> operation.
        ''' </summary>
        Private Shared Sub Perform_ExpandXmlComments()

            Logger.Debug("Perform_ExpandXmlComments — START")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Expand XML comments")
                Logger.Debug("Perform_ExpandXmlComments — undo context opened")

                Dim elements As CodeElements = Snippet_Tool_2026Package.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                Logger.Debug($"Perform_ExpandXmlComments — processing {elements.Count} top-level code element(s)")

                For Each ce As CodeElement2 In elements
                    Logger.Debug($"Perform_ExpandXmlComments — collapsing then toggling element '{ce.Name}', Kind: {ce.Kind}")
                    TextEditorHelper.CollapseSubmembers(ce, False)
                    Logger.Debug($"Perform_ExpandXmlComments — collapse pass done for '{ce.Name}'")
                    TextEditorHelper.CollapseSubmembers(ce, True)
                    Logger.Debug($"Perform_ExpandXmlComments — toggle pass done for '{ce.Name}'")
                Next

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Info("Perform_ExpandXmlComments — END — all comments expanded, undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_ExpandXmlComments — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' <summary>
        ''' Performs the <see cref="XmlDocOperation.DeleteXmlComments"/> operation.
        ''' </summary>
        Private Shared Sub Perform_DeleteXmlComments()

            Logger.Debug("Perform_DeleteXmlComments — START")

            Try
                Snippet_Tool_2026Package.Dte.UndoContext.Open("Delete XML comments")
                Logger.Debug("Perform_DeleteXmlComments — undo context opened")

                Dim elements As CodeElements = Snippet_Tool_2026Package.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                Logger.Debug($"Perform_DeleteXmlComments — processing {elements.Count} top-level code element(s)")

                For Each ce As CodeElement2 In elements
                    Logger.Debug($"Perform_DeleteXmlComments — deleting comment for element '{ce.Name}', Kind: {ce.Kind}")
                    TextEditorHelper.DeleteComment(ce)
                Next

                Snippet_Tool_2026Package.Dte.UndoContext.Close()
                Logger.Info("Perform_DeleteXmlComments — END — all comments deleted, undo context closed")

            Catch ex As Exception
                Logger.Error("Perform_DeleteXmlComments — exception, closing undo context", ex)
                Snippet_Tool_2026Package.Dte.UndoContext.Close()

            End Try

        End Sub

#End Region

#End Region

    End Class

End Namespace
