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
Imports System.Linq

Imports EnvDTE

Imports EnvDTE80

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Editor
Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Formatting
Imports Microsoft.VisualStudio.TextManager.Interop

Imports Snippet_Tool_2026.MyPackage.Core

#End Region

#Region " Text Editor Helper "

Namespace MyPackage.Helpers

    Public NotInheritable Class TextEditorHelper

#Region " Fields "

        ''' <summary>
        ''' Keeps track of the most recently observed <see cref="ITextSnapshot"/> instance, used to detect document changes
        ''' and avoid redundant calls to <see cref="ITextSnapshot.GetText"/>.
        ''' </summary>
        Private Shared _lastSnapshot As ITextSnapshot = Nothing

        ''' <summary>
        ''' Keeps track of the full text content of <see cref="_lastSnapshot"/>, cached to avoid repeated allocations
        ''' on every <see cref="TextContains"/> call.
        ''' </summary>
        Private Shared _lastSnapshotText As String = String.Empty

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Determines whether there is selected text in the active code editor window.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if there is selected text;
        ''' <see langword="False"/> if the selection is empty or no editor window is available.
        ''' </returns>
        Public Shared Function IsTextSelected() As Boolean

            Logger.Debug("IsTextSelected — START")

            Dim viewHost As IWpfTextViewHost = GetCurrentViewHost()
            If viewHost Is Nothing Then
                Logger.Warn("IsTextSelected — viewHost is Nothing, returning False")
                Return False
            End If

            Dim result As Boolean = Not viewHost.TextView.Selection.IsEmpty
            Logger.Debug($"IsTextSelected — END — result: {result}")
            Return result

        End Function

        ''' <summary>
        ''' Determines whether the current text selection spans a single line in the active code editor window.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if there is a non-empty selection contained within a single line;
        ''' <see langword="False"/> if the selection is empty, spans multiple lines, or no editor window is available.
        ''' </returns>
        Public Shared Function SelectedTextIsSingleLine() As Boolean

            Logger.Debug("SelectedTextIsSingleLine — START")

            Dim viewHost As IWpfTextViewHost = GetCurrentViewHost()
            If viewHost Is Nothing Then
                Logger.Warn("SelectedTextIsSingleLine — viewHost is Nothing, returning False")
                Return False
            End If

            Dim selection As ITextSelection = viewHost.TextView.Selection
            If selection.IsEmpty Then
                Logger.Debug("SelectedTextIsSingleLine — selection is empty, returning False")
                Return False
            End If

            Dim startLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(selection.Start.Position)
            Dim endLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(selection.End.Position.Subtract(1))
            Logger.Debug($"SelectedTextIsSingleLine — startLine.Start: {startLine.Start}, endLine.Start: {endLine.Start}")

            Dim result As Boolean = startLine.Equals(endLine)
            Logger.Debug($"SelectedTextIsSingleLine — END — result: {result}")
            Return result

        End Function

        ''' <summary>
        ''' Determines whether the current text selection spans multiple lines in the active code editor window.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if there is a non-empty selection that spans more than one line;
        ''' <see langword="False"/> if the selection is empty, contained within a single line, or no editor window is available.
        ''' </returns>
        Public Shared Function SelectedTextIsMultiLine() As Boolean

            Logger.Debug("SelectedTextIsMultiLine — START")

            Dim viewHost As IWpfTextViewHost = GetCurrentViewHost()
            If viewHost Is Nothing Then
                Logger.Warn("SelectedTextIsMultiLine — viewHost is Nothing, returning False")
                Return False
            End If

            Dim selection As ITextSelection = viewHost.TextView.Selection
            If selection.IsEmpty Then
                Logger.Debug("SelectedTextIsMultiLine — selection is empty, returning False")
                Return False
            End If

            ' Start and end points of the selection.
            Dim selectionStart As VirtualSnapshotPoint = selection.Start
            Dim selectionEnd As VirtualSnapshotPoint = selection.End
            Logger.Debug($"SelectedTextIsMultiLine — selection Start: {selectionStart.Position}, End: {selectionEnd.Position}")

            ' Get the lines that contains the start and end points.
            Dim startLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(selectionStart.Position)
            Dim endLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(selectionEnd.Position.Subtract(1))
            Logger.Debug($"SelectedTextIsMultiLine — startLine.Start: {startLine.Start}, endLine.Start: {endLine.Start}")

            Dim isMultiLineSelection As Boolean = Not startLine.Equals(endLine)
            Logger.Debug($"SelectedTextIsMultiLine — END — isMultiLine: {isMultiLineSelection}")
            Return isMultiLineSelection

        End Function

        ''' <summary>
        ''' Determines whether the caret is positioned within an XML documentation comment block
        ''' in the active code editor window.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the current line starts with <c>'''</c> for Visual Basic .NET
        ''' or <c>///</c> for C#;
        ''' <see langword="False"/> if the caret is not on a documentation comment line or no editor window is available.
        ''' </returns>
        Public Shared Function IsCaretOnXmlBlock() As Boolean

            Logger.Debug("IsCaretOnXmlBlock — START")

            Dim viewHost As IWpfTextViewHost = TextEditorHelper.GetCurrentViewHost()
            If viewHost Is Nothing Then
                Logger.Warn("IsCaretOnXmlBlock — viewHost is Nothing, returning False")
                Return False
            End If

            Dim text As String

            If viewHost.TextView.Selection.IsEmpty Then
                ' No selection — check the line where the caret is.
                Logger.Debug("IsCaretOnXmlBlock — no selection, checking caret line")
                Dim caretLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(
                                             viewHost.TextView.Caret.Position.BufferPosition)
                text = caretLine.Extent.GetText()?.TrimStart()
            Else
                ' Selection exists — check from start line to end line.
                Logger.Debug("IsCaretOnXmlBlock — selection active, checking selection span")
                Dim startLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(viewHost.TextView.Selection.Start.Position)
                Dim endLine As ITextViewLine = viewHost.TextView.GetTextViewLineContainingBufferPosition(viewHost.TextView.Selection.End.Position.Subtract(1))

                ' Create a SnapshotSpan for all text to be replaced.
                Dim span As New SnapshotSpan(startLine.Start, endLine.End)
                text = span.GetText()?.TrimStart()
            End If

            Logger.Debug($"IsCaretOnXmlBlock — text (first 80 chars): '{If(text?.Length > 80, text.Substring(0, 80) & "...", text)}'")

            Dim result As Boolean = Not String.IsNullOrWhiteSpace(text) AndAlso
                            (
                             text.StartsWith(DocsConstants.XmlCommentCharsVB) OrElse
                             text.StartsWith(DocsConstants.XmlCommentCharsCS)
                            )

            Logger.Debug($"IsCaretOnXmlBlock — END — result: {result}")
            Return result

        End Function

        ''' <summary>
        ''' Retrieves the <see cref="IWpfTextViewHost"/> for the currently active text editor window.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' The <see cref="IWpfTextViewHost"/> of the active editor, or <see langword="Nothing"/>
        ''' if no active text view is available or the view does not support WPF hosting.
        ''' </returns>
        Public Shared Function GetCurrentViewHost() As IWpfTextViewHost

            Logger.Debug("GetCurrentViewHost — START")

            Dim txtMgr As IVsTextManager = DirectCast(Snippet_Tool_2026Package.GetGlobalService(GetType(SVsTextManager)), IVsTextManager)
            If txtMgr Is Nothing Then
                Logger.Warn("GetCurrentViewHost — IVsTextManager is Nothing, returning Nothing")
                Return Nothing
            End If

            Logger.Debug("GetCurrentViewHost — IVsTextManager obtained, requesting active view...")

            Dim vTextView As IVsTextView = Nothing
            Const mustHaveFocus As Integer = 1
            Dim hr As Integer = txtMgr.GetActiveView(mustHaveFocus, Nothing, vTextView)

            If ErrorHandler.Failed(hr) OrElse vTextView Is Nothing Then
                Logger.Warn($"GetCurrentViewHost — GetActiveView failed or returned Nothing — HRESULT: 0x{hr:X8}, returning Nothing")
                Return Nothing
            End If

            Logger.Debug("GetCurrentViewHost — active IVsTextView obtained, casting to IVsUserData...")

            Dim userData As IVsUserData = TryCast(vTextView, IVsUserData)
            If userData Is Nothing Then
                Logger.Warn("GetCurrentViewHost — IVsUserData cast failed, returning Nothing")
                Return Nothing
            End If

            Dim viewHost As IWpfTextViewHost
            Dim holder As Object = Nothing
            Dim guidViewHost As Guid = DefGuidList.guidIWpfTextViewHost
            userData.GetData(guidViewHost, holder)
            viewHost = TryCast(holder, IWpfTextViewHost)

            If viewHost Is Nothing Then
                Logger.Warn("GetCurrentViewHost — IWpfTextViewHost cast failed, returning Nothing")
            Else
                Logger.Debug("GetCurrentViewHost — END — IWpfTextViewHost obtained successfully")
            End If

            Return viewHost

        End Function

        ''' <summary>
        ''' Returns the language identifier of the currently active document.
        ''' </summary>
        ''' 
        ''' <param name="dte">
        ''' The <see cref="EnvDTE80.DTE2"/> automation model instance.
        ''' </param>
        ''' 
        ''' <returns>
        ''' A <see cref="String"/> containing the language identifier of the active document;
        ''' or <see langword="Nothing"/> if no document is active or the document
        ''' does not have an associated language.
        ''' </returns>
        Public Shared Function GetActiveDocumentLanguage(dte As EnvDTE80.DTE2) As String

            Logger.Debug("GetActiveDocumentLanguage — START")

            Dim lang As String
            Try
                lang = dte?.ActiveDocument?.ProjectItem?.FileCodeModel?.Language
                Logger.Debug($"GetActiveDocumentLanguage — raw language from FileCodeModel: '{lang}'")
            Catch ex As Exception
                Logger.Error("GetActiveDocumentLanguage — exception reading language from FileCodeModel", ex)
                Return Nothing
            End Try

            If lang Is Nothing Then
                Logger.Warn("GetActiveDocumentLanguage — language is Nothing, returning Nothing")
                Return Nothing
            End If

            Select Case lang

                Case CodeModelLanguageConstants.vsCMLanguageVB
                    Logger.Debug("GetActiveDocumentLanguage — END — returning VB")
                    Return CodeModelLanguageConstants.vsCMLanguageVB

                Case CodeModelLanguageConstants.vsCMLanguageCSharp
                    Logger.Debug("GetActiveDocumentLanguage — END — returning CSharp")
                    Return CodeModelLanguageConstants.vsCMLanguageCSharp

                Case Else ' Not supported language.
                    Logger.Warn($"GetActiveDocumentLanguage — unsupported language '{lang}', returning Nothing")
                    Return Nothing

            End Select

        End Function

        ''' <summary>
        ''' Resolves the XML documentation comment prefix characters appropriate
        ''' for the language of the currently active document.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <c>"'''"</c> for Visual Basic .NET, <c>"///"</c> for C#,
        ''' or <see langword="Nothing"/> if the active document or its language is unavailable.
        ''' </returns>
        Private Shared Function GetActiveDocumentXmlCommentCharsPrefix() As String

            Logger.Debug("GetActiveDocumentXmlCommentCharsPrefix — START")

            Dim lang As String
            Try
                lang = Snippet_Tool_2026Package.Dte?.ActiveDocument?.ProjectItem?.FileCodeModel?.Language
                Logger.Debug($"GetActiveDocumentXmlCommentCharsPrefix — detected language: '{lang}'")
            Catch ex As Exception
                Logger.Error("GetActiveDocumentXmlCommentCharsPrefix — exception reading language from FileCodeModel", ex)
                Return Nothing
            End Try

            If lang Is Nothing Then
                Logger.Warn("GetActiveDocumentXmlCommentCharsPrefix — language is Nothing, returning Nothing")
                Return Nothing
            End If

            Select Case lang

                Case CodeModelLanguageConstants.vsCMLanguageVB
                    Logger.Debug($"GetActiveDocumentXmlCommentCharsPrefix — returning VB comment chars: '{DocsConstants.XmlCommentCharsVB}'")
                    Return DocsConstants.XmlCommentCharsVB

                Case CodeModelLanguageConstants.vsCMLanguageCSharp
                    Logger.Debug($"GetActiveDocumentXmlCommentCharsPrefix — returning CS comment chars: '{DocsConstants.XmlCommentCharsCS}'")
                    Return DocsConstants.XmlCommentCharsCS

                Case Else ' Not supported language.
                    Logger.Warn($"GetActiveDocumentXmlCommentCharsPrefix — unsupported language '{lang}', returning Nothing")
                    Return Nothing

            End Select

        End Function

        ''' <summary>
        ''' Determines whether the active document contains XML documentation comment characters
        ''' appropriate for its language.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the active document contains <c>'''</c> characters for Visual Basic
        ''' or <c>///</c> characters for C#.
        ''' <para></para>
        ''' <see langword="False"/> if the active document is
        ''' unavailable, unsupported, or contains no XML documentation comment characters.
        ''' </returns>
        Public Shared Function ActiveDocumentContainsXmlCommentCharsPrefix() As Boolean

            Logger.Debug("ActiveDocumentContainsXmlCommentCharsPrefix — START")

            Dim findText As String

            Dim lang As String = TextEditorHelper.GetActiveDocumentLanguage(Snippet_Tool_2026Package.Dte)
            If lang Is Nothing Then
                Logger.Warn("ActiveDocumentContainsXmlCommentCharsPrefix — language is Nothing, returning Nothing")
                Return Nothing
            End If

            Logger.Debug($"ActiveDocumentContainsXmlCommentCharsPrefix — detected language: '{lang}'")

            Select Case lang

                Case CodeModelLanguageConstants.vsCMLanguageVB
                    findText = DocsConstants.XmlCommentCharsVB
                    Logger.Debug($"ActiveDocumentContainsXmlCommentCharsPrefix — using VB comment chars: '{findText}'")

                Case CodeModelLanguageConstants.vsCMLanguageCSharp
                    findText = DocsConstants.XmlCommentCharsCS
                    Logger.Debug($"ActiveDocumentContainsXmlCommentCharsPrefix — using CS comment chars: '{findText}'")

                Case Else
                    Logger.Warn($"ActiveDocumentContainsXmlCommentCharsPrefix — unsupported language '{lang}', returning False")
                    Return False

            End Select

            Dim result As Boolean = TextEditorHelper.TextContains(findText)
            Logger.Debug($"ActiveDocumentContainsXmlCommentCharsPrefix — END — result: {result}")
            Return result

        End Function

        ''' <summary>
        ''' Returns an iterator over all direct child members of the given <see cref="CodeElement2"/>,
        ''' traversing both type members and namespace members.
        ''' </summary>
        ''' 
        ''' <param name="ce">
        ''' The <see cref="CodeElement2"/> whose child members will be enumerated.
        ''' </param>
        ''' 
        ''' <returns>
        ''' An <see cref="IEnumerable(Of CodeElement2)"/> containing the direct child members,
        ''' or an empty sequence if the element has no traversable members.
        ''' </returns>
        Private Shared Iterator Function GetSubmembers(ce As CodeElement2) As IEnumerable(Of CodeElement2)

            If ce.IsCodeType Then
                For Each member As CodeElement2 In CType(ce, CodeType).Members
                    Yield member
                Next

            ElseIf ce.Kind = vsCMElement.vsCMElementNamespace Then
                For Each member As CodeElement2 In CType(ce, CodeNamespace).Members
                    Yield member
                Next

            End If
        End Function

        ''' <summary>
        ''' Permanently removes the XML documentation comment block associated with the given
        ''' <see cref="CodeElement2"/> and recursively processes all its child members.
        ''' </summary>
        ''' 
        ''' <param name="ce">
        ''' The <see cref="CodeElement2"/> whose XML documentation comment will be deleted.
        ''' </param>
        ''' 
        ''' <remarks>
        ''' This operation is performed inside a Visual Studio undo transaction, so it can be
        ''' reversed with a single <b>Ctrl+Z</b> regardless of how many comment blocks are deleted.
        ''' </remarks>
        Public Shared Sub DeleteComment(ce As CodeElement2)

            Logger.Debug($"DeleteComment — START — Element: '{ce.Name}', Kind: {ce.Kind}")

            Dim comChars As String = TextEditorHelper.GetActiveDocumentXmlCommentCharsPrefix()
            If comChars Is Nothing Then
                Logger.Warn("DeleteComment — comChars is Nothing (unsupported language?), aborting")
                Return
            End If

            Logger.Debug($"DeleteComment — using comment chars: '{comChars}'")

            Dim undo As UndoContext = Snippet_Tool_2026Package.Dte.UndoContext
            Dim undoOpened As Boolean = False

            Try
                If Not undo.IsOpen Then
                    undo.Open("Delete XML Comments")
                    undoOpened = True
                    Logger.Debug("DeleteComment — undo context opened")
                Else
                    Logger.Debug("DeleteComment — undo context was already open, skipping Open()")
                End If

                Logger.Debug("DeleteComment — calling DeleteCommentCore...")
                TextEditorHelper.DeleteCommentCore(ce, comChars)
                Logger.Info($"DeleteComment — DeleteCommentCore completed for element '{ce.Name}'")

            Catch ex As Exception
                Logger.Error($"DeleteComment — exception during delete for element '{ce.Name}'", ex)
                If undoOpened AndAlso undo.IsOpen Then
                    undo.SetAborted()
                    undoOpened = False
                    Logger.Warn("DeleteComment — undo context aborted due to exception")
                End If
                Throw

            Finally
                If undoOpened AndAlso undo.IsOpen Then
                    undo.Close()
                    Logger.Debug("DeleteComment — undo context closed")
                End If

            End Try

            Logger.Debug($"DeleteComment — END — Element: '{ce.Name}'")

        End Sub

        ''' <summary>
        ''' Core recursive implementation of <see cref="DeleteComment"/>.
        ''' Removes the XML documentation comment block of the given element
        ''' and recurses into all child members.
        ''' </summary>
        ''' 
        ''' <param name="ce">
        ''' The <see cref="CodeElement2"/> to process.
        ''' </param>
        ''' 
        ''' <param name="comChars">
        ''' The XML comment prefix characters appropriate for the active document language.
        ''' </param>
        Private Shared Sub DeleteCommentCore(ce As CodeElement2, comChars As String)

            Logger.Debug($"DeleteCommentCore — START — Element: '{ce.Name}', Kind: {ce.Kind}, ComChars: '{comChars}'")

            Dim memberStart As EditPoint = ce.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint()
            Dim commentStart As EditPoint = GetCommentStart(memberStart.CreateEditPoint(), comChars)

            If commentStart IsNot Nothing Then
                Dim commentEnd As EditPoint = GetCommentEnd(commentStart.CreateEditPoint(), comChars)
                Logger.Debug($"DeleteCommentCore — comment block found, deleting from line {commentStart.Line} to line {commentEnd.Line}")
                commentStart.ReplaceText(commentEnd, String.Empty, 0)
                Logger.Info($"DeleteCommentCore — comment deleted from element '{ce.Name}'")
            Else
                Logger.Debug($"DeleteCommentCore — no comment block found for element '{ce.Name}'")
            End If

            Dim submembers As List(Of CodeElement2) = TextEditorHelper.GetSubmembers(ce).ToList()
            Logger.Debug($"DeleteCommentCore — processing {submembers.Count} submember(s) of '{ce.Name}'")

            For Each member As CodeElement2 In submembers
                TextEditorHelper.DeleteCommentCore(member, comChars)
            Next

            Logger.Debug($"DeleteCommentCore — END — Element: '{ce.Name}'")

        End Sub

        ''' <summary>
        ''' Collapses or toggles the XML documentation comment outline region associated with
        ''' the given <see cref="CodeElement2"/> and recursively processes all its child members.
        ''' </summary>
        ''' 
        ''' <param name="ce">
        ''' The <see cref="CodeElement2"/> whose XML documentation comment outline will be processed.
        ''' </param>
        ''' 
        ''' <param name="toggle">
        ''' <see langword="True"/> to toggle the current outlining expansion state;
        ''' <see langword="False"/> to always collapse the outline region.
        ''' </param>
        ''' 
        ''' <remarks>
        ''' This operation is performed inside a Visual Studio undo transaction, so it can be
        ''' reversed with a single <b>Ctrl+Z</b> regardless of how many regions are affected.
        ''' </remarks>
        Public Shared Sub CollapseSubmembers(ce As CodeElement2, toggle As Boolean)

            Logger.Debug($"CollapseSubmembers — START — Element: '{ce.Name}', Toggle: {toggle}")

            Dim comChars As String = TextEditorHelper.GetActiveDocumentXmlCommentCharsPrefix()
            If comChars Is Nothing Then
                Logger.Warn("CollapseSubmembers — comChars is Nothing (unsupported language?), aborting")
                Return
            End If

            Logger.Debug($"CollapseSubmembers — using comment chars: '{comChars}'")

            Dim undo As UndoContext = Snippet_Tool_2026Package.Dte.UndoContext
            Dim undoOpened As Boolean = False

            Try
                If Not undo.IsOpen Then
                    undo.Open("Collapse XML Comments")
                    undoOpened = True
                    Logger.Debug("CollapseSubmembers — undo context opened")
                Else
                    Logger.Debug("CollapseSubmembers — undo context was already open, skipping Open()")
                End If

                Logger.Debug($"CollapseSubmembers — calling CollapseSubmembersCore...")
                TextEditorHelper.CollapseSubmembersCore(ce, comChars, toggle)
                Logger.Info($"CollapseSubmembers — CollapseSubmembersCore completed for element '{ce.Name}'")

            Catch ex As Exception
                Logger.Error($"CollapseSubmembers — exception during collapse for element '{ce.Name}'", ex)
                If undoOpened AndAlso undo.IsOpen Then
                    undo.SetAborted()
                    undoOpened = False
                    Logger.Warn("CollapseSubmembers — undo context aborted due to exception")
                End If
                Throw

            Finally
                If undoOpened AndAlso undo.IsOpen Then
                    undo.Close()
                    Logger.Debug("CollapseSubmembers — undo context closed")
                End If

            End Try

            Logger.Debug($"CollapseSubmembers — END — Element: '{ce.Name}'")

        End Sub

        ''' <summary>
        ''' Core recursive implementation of <see cref="CollapseSubmembers"/>.
        ''' Collapses or toggles the outline region of the given element
        ''' and recurses into all child members.
        ''' </summary>
        ''' 
        ''' <param name="ce">
        ''' The <see cref="CodeElement2"/> to process.
        ''' </param>
        ''' 
        ''' <param name="comChars">
        ''' The XML comment prefix characters appropriate for the active document language.
        ''' </param>
        ''' 
        ''' <param name="toggle">
        ''' <see langword="True"/> to toggle the current outlining expansion state;
        ''' <see langword="False"/> to always collapse the outline region.
        ''' </param>
        Private Shared Sub CollapseSubmembersCore(ce As CodeElement2, comChars As String, toggle As Boolean)

            Logger.Debug($"CollapseSubmembersCore — START — Element: '{ce.Name}', Kind: {ce.Kind}, ComChars: '{comChars}', Toggle: {toggle}")

            Dim memberStart As EditPoint = ce.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint()
            Dim commentStart As EditPoint = GetCommentStart(memberStart.CreateEditPoint(), comChars)

            If commentStart IsNot Nothing Then
                Dim commentEnd As EditPoint = GetCommentEnd(commentStart.CreateEditPoint(), comChars)
                Logger.Debug($"CollapseSubmembersCore — comment block found, lines {commentStart.Line} to {commentEnd.Line}")

                If toggle Then
                    Logger.Debug($"CollapseSubmembersCore — toggle mode: moving caret to line {commentStart.Line} and executing ToggleOutliningExpansion")
                    CType(Snippet_Tool_2026Package.Dte.ActiveDocument.Selection, TextSelection).MoveToPoint(commentStart)
                    Snippet_Tool_2026Package.Dte.ExecuteCommand("Edit.ToggleOutliningExpansion")
                    Logger.Info($"CollapseSubmembersCore — toggled outlining for element '{ce.Name}'")
                Else
                    Logger.Debug($"CollapseSubmembersCore — outline mode: calling OutlineSection from line {commentStart.Line} to {commentEnd.Line}")
                    commentStart.OutlineSection(commentEnd)
                    Logger.Info($"CollapseSubmembersCore — outlined section for element '{ce.Name}'")
                End If
            Else
                Logger.Debug($"CollapseSubmembersCore — no comment block found for element '{ce.Name}'")
            End If

            Dim submembers As List(Of CodeElement2) = TextEditorHelper.GetSubmembers(ce).ToList()
            Logger.Debug($"CollapseSubmembersCore — processing {submembers.Count} submember(s) of '{ce.Name}'")

            For Each member As CodeElement2 In submembers
                TextEditorHelper.CollapseSubmembersCore(member, comChars, toggle)
            Next

            Logger.Debug($"CollapseSubmembersCore — END — Element: '{ce.Name}'")

        End Sub

        ''' <summary>
        ''' Locates the <see cref="EditPoint"/> at the beginning of the XML documentation comment block
        ''' that precedes the given member start point, by scanning upward through consecutive comment lines.
        ''' </summary>
        ''' 
        ''' <param name="ep">
        ''' An <see cref="EditPoint"/> positioned at the start of the documented member.
        ''' </param>
        ''' 
        ''' <param name="commentChars">
        ''' The XML documentation comment prefix to search for:
        ''' <c>"'''"</c> for Visual Basic .NET or <c>"///"</c> for C#.
        ''' </param>
        ''' 
        ''' <returns>
        ''' An <see cref="EditPoint"/> positioned at the first character of the opening comment line,
        ''' or <see langword="Nothing"/> if no comment block is found or an error occurs.
        ''' </returns>
        Private Shared Function GetCommentStart(ep As EditPoint, commentChars As String) As EditPoint

            Try
                Dim line As String
                Dim lastCommentLine As Integer = 0

                ep.StartOfLine()
                ep.CharLeft()

                While Not ep.AtStartOfDocument
                    line = ep.GetLines(ep.Line, ep.Line + 1).Trim()
                    If line.Length = 0 OrElse line.StartsWith(commentChars, StringComparison.Ordinal) Then
                        If line.Length > 0 Then
                            lastCommentLine = ep.Line
                        End If
                        ep.StartOfLine()
                        ep.CharLeft()
                    Else
                        Exit While
                    End If
                End While

                If lastCommentLine = 0 Then
                    Return Nothing
                End If

                ep.MoveToLineAndOffset(lastCommentLine, 1)

                While Not ep.GetText(commentChars.Length).Equals(commentChars, StringComparison.Ordinal)
                    ep.CharRight()
                End While

                Return ep.CreateEditPoint()

            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Locates the <see cref="EditPoint"/> at the end of the XML documentation comment block
        ''' that precedes the given member, by scanning forward through consecutive comment lines.
        ''' </summary>
        ''' 
        ''' <param name="ep">
        ''' An <see cref="EditPoint"/> positioned at the start of the comment block.
        ''' </param>
        ''' 
        ''' <param name="commentChars">
        ''' The XML documentation comment prefix to search for:
        ''' <c>"'''"</c> for Visual Basic .NET or <c>"///"</c> for C#.
        ''' </param>
        ''' 
        ''' <returns>
        ''' An <see cref="EditPoint"/> positioned at the end of the last comment line,
        ''' or <see langword="Nothing"/> if an error occurs.
        ''' </returns>
        Private Shared Function GetCommentEnd(ep As EditPoint, commentChars As String) As EditPoint

            Try
                Dim line As String
                Dim lastCommentPoint As EditPoint = ep.CreateEditPoint()

                ep.EndOfLine()
                ep.CharRight()

                While Not ep.AtEndOfDocument
                    line = ep.GetLines(ep.Line, ep.Line + 1).Trim()
                    If line.StartsWith(commentChars, StringComparison.Ordinal) Then
                        lastCommentPoint = ep.CreateEditPoint()
                        ep.EndOfLine()
                        ep.CharRight()
                    Else
                        Exit While
                    End If
                End While

                lastCommentPoint.EndOfLine()
                Return lastCommentPoint

            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Determines whether the active document contains the specified text,
        ''' using a snapshot cache to avoid redundant full-document reads.
        ''' </summary>
        ''' 
        ''' <param name="findText">
        ''' The text to search for within the active document.
        ''' </param>
        ''' 
        ''' <param name="ignoreCase">
        ''' <see langword="True"/> to perform a case-insensitive search;
        ''' <see langword="False"/> to perform a case-sensitive search.
        ''' The default value is <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the active document contains <paramref name="findText"/>;
        ''' <see langword="False"/> if the text is not found or no editor window is available.
        ''' </returns>
        Public Shared Function TextContains(findText As String, Optional ignoreCase As Boolean = False) As Boolean

            Logger.Debug($"TextContains — findText: '{findText}', ignoreCase: {ignoreCase}")

            Dim viewHost As IWpfTextViewHost = GetCurrentViewHost()
            If viewHost Is Nothing Then
                Logger.Warn("TextContains — viewHost is Nothing, returning False")
                Return False
            End If

            Dim snapshot As ITextSnapshot = viewHost.TextView.TextSnapshot

            ' Only update cached text if the snapshot instance has changed.
            If snapshot IsNot TextEditorHelper._lastSnapshot Then
                Logger.Debug("TextContains — snapshot changed, updating cache")
                TextEditorHelper._lastSnapshot = snapshot
                TextEditorHelper._lastSnapshotText = snapshot.GetText()
                Logger.Debug($"TextContains — cache updated, new text length: {TextEditorHelper._lastSnapshotText.Length}")
            Else
                Logger.Debug("TextContains — using cached snapshot text")
            End If

            ' Search in cached text.
            Dim result As Boolean = If(ignoreCase,
                               TextEditorHelper._lastSnapshotText.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0,
                               TextEditorHelper._lastSnapshotText.IndexOf(findText, StringComparison.Ordinal) >= 0)

            Logger.Debug($"TextContains — result: {result}")
            Return result

        End Function

        ' UNUSED. Kept for potential future use if line-specific searches are needed.
        ' ---------------------------------------------------------------------------
        'Public Shared Function TextContainsInCurrentLine(findText As String,
        '                                                 Optional ignoreCase As Boolean = False) As Boolean
        '
        '    Logger.Debug($"TextContainsInCurrentLine — START — findText: '{findText}', ignoreCase: {ignoreCase}")
        '
        '    Dim viewHost As IWpfTextViewHost = GetCurrentViewHost()
        '    If viewHost Is Nothing Then
        '        Logger.Warn("TextContainsInCurrentLine — viewHost is Nothing, returning False")
        '        Return False
        '    End If
        '
        '    Dim caretPosition As SnapshotPoint =
        '        viewHost.TextView.Caret.Position.BufferPosition
        '
        '    Logger.Debug($"TextContainsInCurrentLine — caret buffer position: {caretPosition.Position}")
        '
        '    Dim line As ITextSnapshotLine =
        '        caretPosition.GetContainingLine()
        '
        '    Dim lineText As String =
        '        line.GetText()
        '
        '    Logger.Debug($"TextContainsInCurrentLine — current line {line.LineNumber}, text: '{lineText}'")
        '
        '    Dim comparison As StringComparison =
        '        If(ignoreCase,
        '           StringComparison.OrdinalIgnoreCase,
        '           StringComparison.Ordinal)
        '
        '    Dim result As Boolean = lineText.IndexOf(findText, comparison) >= 0
        '    Logger.Debug($"TextContainsInCurrentLine — END — result: {result}")
        '    Return result
        'End Function

#End Region

    End Class

End Namespace

#End Region
