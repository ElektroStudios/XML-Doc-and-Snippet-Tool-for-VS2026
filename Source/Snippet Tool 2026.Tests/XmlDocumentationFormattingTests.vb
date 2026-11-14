Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml
Imports System.Xml.Linq
Imports Xunit
Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.Helpers

Public Class XmlDocumentationFormattingTests

    Private Shared ReadOnly ValidTextValues As String() = {
        String.Empty, "A", "word", "has spaces", "a&b", "<T>", "x > y", "a""b",
        "a'b", "<&>""'", "café", "漢字", "Emoji 😀", vbTab,
        "a" & vbCrLf & "b", "a" & vbLf & "b", "a" & vbCr & "b", "a$end$z", "]]>",
        "\u0026", "quotes ""here"" & <this>", "a  b", "  padded  ", "\", "a=b"
    }

    ''' <summary>
    ''' Supplies real documentation attribute names and text values, including XML-sensitive characters.
    ''' </summary>
    Public Shared Iterator Function AttributeCases() As IEnumerable(Of TheoryDataRow(Of String, String, String))
        Dim attributes As String() = {"see|cref", "paramref|name", "see|langword", "see|href", "seealso|href"}
        For Each definition As String In attributes
            Dim parts As String() = definition.Split("|"c)
            For Each value As String In ValidTextValues
                Yield New TheoryDataRow(Of String, String, String)(parts(0), parts(1), value)
            Next
        Next
    End Function

    ''' <summary>
    ''' Ensures that generated attribute wrappers parse as XML and preserve the original value.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(AttributeCases))>
    Public Sub AttributeWrapperProducesValidXml(tagName As String, attributeName As String, value As String)
        Dim output As String = XmlDocumentationFormatting.WrapAttribute(tagName, attributeName, value)
        Dim element As XElement = XElement.Parse(output, LoadOptions.PreserveWhitespace)
        Assert.Equal(tagName, element.Name.LocalName)
        Assert.Equal(value, CStr(element.Attribute(attributeName)))
        Assert.False(element.HasElements)
    End Sub

    ''' <summary>
    ''' Supplies actual XML documentation element names with varied source text.
    ''' </summary>
    Public Shared Iterator Function ElementCases() As IEnumerable(Of TheoryDataRow(Of String, String))
        Dim tags As String() = {"c", "code", "b", "i", "u", "para", "remarks"}
        For Each tag As String In tags
            For Each value As String In ValidTextValues
                Yield New TheoryDataRow(Of String, String)(tag, value)
            Next
        Next
    End Function

    ''' <summary>
    ''' Ensures that element wrappers produce valid XML and do not lose or reinterpret source text.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(ElementCases))>
    Public Sub ElementWrapperPreservesXmlSensitiveText(tagName As String, value As String)
        Dim output As String = XmlDocumentationFormatting.WrapElement(tagName, value)
        Dim element As XElement = XElement.Parse(output, LoadOptions.PreserveWhitespace)
        Assert.Equal(tagName, element.Name.LocalName)
        Assert.Equal(value, element.Value)
    End Sub

    Private Shared ReadOnly SnippetValues As String() = {
        String.Empty, "X", "Console.WriteLine(42)", "If x < 3 Then", "var x = a && b;",
        "a&b", "<code/>", "a""b", "a'b", "漢字", "😀", "a=b", "{0}",
        "{{braces}}", "$end$", "$selected$", "$cdataend$", "]]>", "a]]>b",
        "]]>]]>", "<![CDATA[x]]>", "x" & vbLf & "y", "x" & vbCrLf & "y",
        """"""", "' <example>", "// </Code>", "a]]>b$cdataend$c", vbTab & "x"
    }

    ''' <summary>
    ''' Supplies both production snippet templates and source payloads.
    ''' </summary>
    Public Shared Iterator Function SnippetCases() As IEnumerable(Of TheoryDataRow(Of String, String, String))
        For Each source As String In SnippetValues
            Yield New TheoryDataRow(Of String, String, String)("VB", DocsConstants.SnippetTemplateFormatVB, source)
            Yield New TheoryDataRow(Of String, String, String)("CSharp", DocsConstants.SnippetTemplateFormatCS, source)
        Next
    End Function

    ''' <summary>
    ''' Validates real snippet XML after formatting special character sequences.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(SnippetCases))>
    Public Sub SnippetContainsOriginalSourceInValidCdata(language As String, template As String, source As String)
        Dim result As String = XmlDocumentationFormatting.FormatSnippet(template, "Example Author", source)
        Dim document As XDocument = XDocument.Parse(result, LoadOptions.PreserveWhitespace)
        Dim code As XElement = document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Code")
        Assert.Equal(language, CStr(code.Attribute("Language")))
        If source.Length > 0 Then
            Assert.Contains(source, code.Value, StringComparison.Ordinal)
        End If
        Assert.Equal("Example Author", document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Author").Value)
    End Sub

    ''' <summary>
    ''' Supplies author names requiring XML escaping, including text identical to template markers.
    ''' </summary>
    Public Shared Iterator Function AuthorCases() As IEnumerable(Of TheoryDataRow(Of String, String))
        Dim authors As String() = {String.Empty, "A", "A&B", "<Admin>", "a>b", """Quote""", "O'Neil",
            "<&>""'", "áéíóú", "中文", "$cdataend$", "{0}", "a" & vbLf & "b", "  Name  ", "a" & vbCr & "b", "a" & vbCrLf & "b", "a" & vbTab & "b"}
        For Each author As String In authors
            Yield New TheoryDataRow(Of String, String)(DocsConstants.SnippetTemplateFormatVB, author)
            Yield New TheoryDataRow(Of String, String)(DocsConstants.SnippetTemplateFormatCS, author)
        Next
    End Function

    ''' <summary>
    ''' Prevents a malformed snippet when the configured author has XML punctuation.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(AuthorCases))>
    Public Sub SnippetEscapesAuthorWithoutChangingItsMeaning(template As String, author As String)
        Dim result As String = XmlDocumentationFormatting.FormatSnippet(template, author, "Dim result As Integer = 1")
        Dim document As XDocument = XDocument.Parse(result, LoadOptions.PreserveWhitespace)
        Assert.Equal(author, document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Author").Value)
    End Sub

    ''' <summary>
    ''' Returns indentation cases including one-character lines and all-whitespace input.
    ''' </summary>
    Public Shared Iterator Function IndentationCases() As IEnumerable(Of TheoryDataRow(Of String, Integer))
        Dim lines As String() = {String.Empty, "a", "ab", " abc", "  abc", vbTab & "abc", "     x" & vbLf & "      y", "   ", vbTab, "x" & vbLf & "y"}
        Dim expected As Integer() = {0, 0, 0, 1, 2, 1, 5, 0, 0, 0}
        For index As Integer = 0 To lines.Length - 1
            Yield New TheoryDataRow(Of String, Integer)(lines(index), expected(index))
        Next
        Yield New TheoryDataRow(Of String, Integer)("  a" & vbCrLf & " b", 1)
        Yield New TheoryDataRow(Of String, Integer)("  a" & vbLf & vbLf & "  b", 2)
        Yield New TheoryDataRow(Of String, Integer)(vbTab & "a" & vbCr & vbTab & vbTab & "b", 1)
        Yield New TheoryDataRow(Of String, Integer)("  x" & vbLf & "y", 0)
        Yield New TheoryDataRow(Of String, Integer)(vbLf & vbCrLf, 0)
        Yield New TheoryDataRow(Of String, Integer)("   " & vbLf & vbTab & " " & vbLf, 0)
        Yield New TheoryDataRow(Of String, Integer)("a" & vbLf & "  b", 0)
        Yield New TheoryDataRow(Of String, Integer)("    a" & vbLf & "     b", 4)
        Yield New TheoryDataRow(Of String, Integer)("   a" & vbLf & " b" & vbLf & "      c", 1)
        Yield New TheoryDataRow(Of String, Integer)(" " & vbLf & "   x" & vbLf & " ", 3)
        For indentation As Integer = 0 To 16
            Yield New TheoryDataRow(Of String, Integer)(New String(" "c, indentation) & "x", indentation)
        Next
    End Function

    ''' <summary>
    ''' Verifies that indentation is calculated safely for short and whitespace-only lines.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(IndentationCases))>
    Public Sub MinimumIndentationHandlesBoundaryCases(source As String, expected As Integer)
        Assert.Equal(expected, XmlDocumentationFormatting.GetMinimumIndentation(source))
    End Sub

    ''' <summary>
    ''' Rejects invalid XML characters rather than producing corrupt documentation.
    ''' </summary>
    <Fact>
    Public Sub XmlFormattingRejectsControlCharacters()
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapAttribute("see", "href", "x" & ChrW(1)))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapElement("c", "x" & ChrW(1)))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.FormatSnippet(DocsConstants.SnippetTemplateFormatVB, "a", "x" & ChrW(1)))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.FormatSnippet(DocsConstants.SnippetTemplateFormatVB, "a" & ChrW(1), "x"))
    End Sub

    ''' <summary>
    ''' Rejects invalid tag names rather than constructing malformed XML.
    ''' </summary>
    <Fact>
    Public Sub XmlFormattingRejectsInvalidElementOrAttributeNames()
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapAttribute("invalid tag", "href", "x"))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapAttribute("see", "invalid attr", "x"))
        Assert.Throws(Of XmlException)(Sub() XmlDocumentationFormatting.WrapElement("not valid", "x"))
    End Sub

    ''' <summary>
    ''' Rejects missing arguments before XML formatting is attempted.
    ''' </summary>
    <Fact>
    Public Sub XmlFormattingRejectsNullArguments()
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapAttribute(Nothing, "href", "x"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapAttribute("see", Nothing, "x"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapAttribute("see", "href", Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapElement(Nothing, "x"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.WrapElement("c", Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatSnippet(Nothing, "a", "b"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatSnippet("x", Nothing, "b"))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.FormatSnippet("x", "a", Nothing))
        Assert.Throws(Of ArgumentNullException)(Sub() XmlDocumentationFormatting.GetMinimumIndentation(Nothing))
    End Sub
End Class
