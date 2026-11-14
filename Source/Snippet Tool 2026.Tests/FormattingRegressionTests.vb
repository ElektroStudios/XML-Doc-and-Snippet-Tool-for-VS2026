Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq
Imports Xunit
Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.Helpers

''' <summary>Exercises regressions in the formatting code used by the editor commands.</summary>
Public Class FormattingRegressionTests

    ''' <summary>Provides supported line endings and both documentation languages.</summary>
    Public Shared Iterator Function BlockCases() As IEnumerable(Of TheoryDataRow(Of String, String, String))
        For Each ending As String In New String() {vbCrLf, vbLf, vbCr}
            Yield New TheoryDataRow(Of String, String, String)("'''", "VB", ending)
            Yield New TheoryDataRow(Of String, String, String)("///", "CSharp", ending)
        Next
    End Function

    ''' <summary>Provides comment prefixes and line endings for language-independent remarks formatting.</summary>
    Public Shared Iterator Function DocumentationBlockCases() As IEnumerable(Of TheoryDataRow(Of String, String))
        For Each ending As String In New String() {vbCrLf, vbLf, vbCr}
            Yield New TheoryDataRow(Of String, String)("'''", ending)
            Yield New TheoryDataRow(Of String, String)("///", ending)
        Next
    End Function

    ''' <summary>Removes comment prefixes so a generated block can be parsed as XML.</summary>
    Private Shared Function ParseBlock(block As String, prefix As String) As XElement
        Dim lines As New List(Of String)()
        For Each line As String In XmlDocumentationFormatting.SplitLines(block)
            If line.Length = 0 Then Continue For
            Dim trimmed As String = line.TrimStart()
            Assert.True(trimmed.StartsWith(prefix & " ", StringComparison.Ordinal), line)
            lines.Add(trimmed.Substring(prefix.Length + 1))
        Next
        Return XElement.Parse(String.Join(vbLf, lines), LoadOptions.PreserveWhitespace)
    End Function

    ''' <summary>Checks exact nesting, escaping, relative indentation, and blank lines.</summary>
    <Theory, MemberData(NameOf(BlockCases))>
    Public Sub CodeExamplePreservesRelativeIndentation(prefix As String, language As String, ending As String)
        Dim source As String = "  If x < 2 Then" & ending & "    Print(""A&B"")" & ending & " " & ending & "  End If"
        Dim output As String = XmlDocumentationFormatting.FormatCodeExample(source, prefix, language)
        Dim root As XElement = ParseBlock(output, prefix)
        Assert.Equal("example", root.Name.LocalName)
        Dim code As XElement = root.Element("code")
        Assert.NotNull(code)
        Assert.Equal(language, CStr(code.Attribute("language")))
        Assert.Equal(vbLf & "If x < 2 Then" & vbLf & "  Print(""A&B"")" & vbLf & vbLf & "End If" & vbLf, code.Value)
        Assert.False(code.HasElements)
        Assert.True(output.StartsWith("  " & prefix & " <example>", StringComparison.Ordinal))
        Assert.True(output.EndsWith("  " & prefix & " </example>" & ending, StringComparison.Ordinal))
        Assert.Equal(ending, XmlDocumentationFormatting.GetNewLine(output))
    End Sub

    ''' <summary>Prevents all-whitespace selections from throwing or producing invalid XML.</summary>
    <Theory, MemberData(NameOf(BlockCases))>
    Public Sub EmptyCodeExampleIsValid(prefix As String, language As String, ending As String)
        Dim output As String = XmlDocumentationFormatting.FormatCodeExample(" " & ending & vbTab, prefix, language)
        Assert.NotNull(ParseBlock(output, prefix).Element("code"))
    End Sub

    ''' <summary>Retains nested documentation spanning several lines without duplicating prefixes.</summary>
    <Theory, MemberData(NameOf(DocumentationBlockCases))>
    Public Sub RemarksRetainsMultilineMarkup(prefix As String, ending As String)
        Dim source As String = "  " & prefix & " <para>" & ending & "  " & prefix & " A &amp; B <see cref=""T""/>" & ending & "  " & prefix & " </para>"
        Dim output As String = XmlDocumentationFormatting.FormatRemarks(source, prefix)
        Dim root As XElement = ParseBlock(output, prefix)
        Assert.Equal("remarks", root.Name.LocalName)
        Dim paragraph As XElement = root.Element("para")
        Assert.NotNull(paragraph)
        Assert.Equal("T", CStr(paragraph.Element("see").Attribute("cref")))
        Assert.Contains("A & B", paragraph.Value, StringComparison.Ordinal)
        Assert.False(root.Value.Contains(prefix))
        Assert.Equal(ending, XmlDocumentationFormatting.GetNewLine(output))
    End Sub

    ''' <summary>Escapes raw prose that is not a well-formed XML fragment.</summary>
    <Theory, MemberData(NameOf(DocumentationBlockCases))>
    Public Sub RemarksEscapesPlainText(prefix As String, ending As String)
        Dim root As XElement = ParseBlock(XmlDocumentationFormatting.FormatRemarks("A&B" & ending & "x < y", prefix), prefix)
        Assert.False(root.HasElements)
        Assert.Equal(vbLf & "A&B" & vbLf & "x < y" & vbLf, root.Value)
    End Sub

    ''' <summary>Provides the formatting elements that must retain existing documentation markup.</summary>
    Public Shared Iterator Function FragmentCases() As IEnumerable(Of TheoryDataRow(Of String, String))
        For Each tag As String In New String() {"b", "i", "u", "para"}
            Yield New TheoryDataRow(Of String, String)(tag, "A &amp; B <see cref=""T""/>")
            Yield New TheoryDataRow(Of String, String)(tag, "<c>x &lt; y</c>")
            Yield New TheoryDataRow(Of String, String)(tag, "<list>" & vbLf & "<item>Text</item>" & vbLf & "</list>")
        Next
    End Function

    ''' <summary>Ensures that formatting a reference does not turn it into visible XML text.</summary>
    <Theory, MemberData(NameOf(FragmentCases))>
    Public Sub FormattingRetainsExistingXml(tag As String, fragment As String)
        Dim output As String = XmlDocumentationFormatting.WrapDocumentationFragment(tag, fragment)
        Assert.Equal("<" & tag & ">" & fragment & "</" & tag & ">", output)
        Assert.True(XElement.Parse(output).HasElements)
    End Sub

    ''' <summary>Preserves plain text and prevents external entity expansion.</summary>
    <Theory>
    <InlineData("A&B")>
    <InlineData("x < y")>
    <InlineData("<!DOCTYPE a [<!ENTITY e SYSTEM 'file:///nonexistent'>]><a>&e;</a>")>
    <InlineData("<?xml version='1.0'?><x/>")>
    Public Sub FormattingEscapesInvalidFragments(value As String)
        Dim element As XElement = XElement.Parse(XmlDocumentationFormatting.WrapDocumentationFragment("b", value), LoadOptions.PreserveWhitespace)
        Assert.Equal(value, element.Value)
        Assert.False(element.HasElements)
    End Sub

    ''' <summary>Ensures a code operation still treats code-like XML as literal source.</summary>
    <Fact>
    Public Sub LiteralCodeDoesNotBecomeNestedXml()
        Dim element As XElement = XElement.Parse(XmlDocumentationFormatting.WrapElement("code", "<see cref=""T""/>"))
        Assert.False(element.HasElements)
        Assert.Equal("<see cref=""T""/>", element.Value)
    End Sub

    ''' <summary>Verifies exact line splitting, including mixed line endings and terminal empty lines.</summary>
    <Fact>
    Public Sub SplitLinesRetainsAllBoundaries()
        Assert.Equal(Of String)(DirectCast(New String() {"a", "b", "c", "", ""}, IEnumerable(Of String)), XmlDocumentationFormatting.SplitLines("a" & vbCrLf & "b" & vbLf & "c" & vbCr & vbLf & vbLf))
    End Sub

    ''' <summary>Checks empty input and each newline preference independently of the host OS.</summary>
    <Theory>
    <InlineData("", "\r\n")>
    <InlineData("x", "\r\n")>
    <InlineData("x\ny\r\n", "\n")>
    <InlineData("x\ry\n", "\r")>
    <InlineData("x\r\ny", "\r\n")>
    Public Sub FirstLineEndingDeterminesOutput(source As String, expected As String)
        Assert.Equal(Unescape(expected), XmlDocumentationFormatting.GetNewLine(Unescape(source)))
    End Sub

    ''' <summary>Converts readable data row line endings into actual control characters.</summary>
    Private Shared Function Unescape(value As String) As String
        Return value.Replace("\r", vbCr).Replace("\n", vbLf)
    End Function

    ''' <summary>Prevents position -1 and selecting the previous line when the selection is empty.</summary>
    <Theory>
    <InlineData(0, 0, 0)>
    <InlineData(5, 5, 5)>
    <InlineData(0, 1, 0)>
    <InlineData(0, 10, 9)>
    <InlineData(4, 7, 6)>
    <InlineData(Integer.MaxValue, Integer.MaxValue, Integer.MaxValue)>
    Public Sub SelectionEndHandlesEmptyAndNonemptyRanges(startPosition As Integer, endPosition As Integer, expected As Integer)
        Assert.Equal(expected, XmlDocumentationFormatting.GetSelectionEndPosition(startPosition, endPosition))
    End Sub

    ''' <summary>Rejects inconsistent or negative selection ranges.</summary>
    <Theory>
    <InlineData(-1, 0)>
    <InlineData(1, 0)>
    <InlineData(0, -1)>
    Public Sub SelectionEndRejectsInvalidRanges(startPosition As Integer, endPosition As Integer)
        Assert.Throws(Of ArgumentOutOfRangeException)(Sub() XmlDocumentationFormatting.GetSelectionEndPosition(startPosition, endPosition))
    End Sub

    ''' <summary>Ensures the file bytes match the XML declaration and preserve Unicode source.</summary>
    <Theory>
    <InlineData("VB")>
    <InlineData("CSharp")>
    Public Sub SnippetFileIsUtf8WithExactSource(language As String)
        Dim template As String = If(language = "VB", DocsConstants.SnippetTemplateFormatVB, DocsConstants.SnippetTemplateFormatCS)
        Dim source As String = "á漢😀 $cdataend$ ]]>" & vbCrLf & "  End"
        Dim content As String = XmlDocumentationFormatting.FormatSnippet(template, "José & Co", source)
        Dim path As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".snippet")
        Try
            XmlDocumentationFormatting.WriteSnippet(path, content)
            Dim bytes As Byte() = File.ReadAllBytes(path)
            Assert.Equal(content, New UTF8Encoding(False, True).GetString(bytes))
            Assert.False(bytes.Take(3).SequenceEqual(New Byte() {&HEF, &HBB, &HBF}))
            Dim document As XDocument = XDocument.Load(path, LoadOptions.PreserveWhitespace)
            Dim code As XElement = document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Code")
            Assert.Equal(vbLf & vbLf & "        " & source & vbLf & vbLf & "        ", code.Value)
            Assert.Equal("José & Co", document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Author").Value)
        Finally
            If File.Exists(path) Then File.Delete(path)
        End Try
    End Sub

    ''' <summary>Rejects namespace-qualified names because this helper does not declare namespaces.</summary>
    <Theory>
    <InlineData("a:b")>
    <InlineData("1tag")>
    Public Sub XmlNamesMustBeUnqualified(name As String)
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapElement(name, "x"))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapAttribute("see", name, "x"))
    End Sub

    ''' <summary>Checks null arguments and invalid prefixes before any formatting or file I/O.</summary>
    <Fact>
    Public Sub NewHelpersRejectInvalidArguments()
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.SplitLines(Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.GetNewLine(Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapDocumentationFragment("b", Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapDocumentationFragment(Nothing, "x"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatCodeExample(Nothing, "///", "VB"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatCodeExample("x", Nothing, "VB"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatCodeExample("x", "///", Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatRemarks(Nothing, "///"))
        Assert.Throws(Of ArgumentException)(Sub() XmlDocumentationFormatting.FormatRemarks("x", "//"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WriteSnippet(Nothing, "x"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WriteSnippet("x", Nothing))
    End Sub
End Class
