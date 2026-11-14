Option Strict On
Option Explicit On
Option Infer Off

Imports Microsoft.VisualBasic

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports System.IO
Imports System.Security
Imports System.Xml

Namespace MyPackage.Helpers

    ''' <summary>
    ''' Provides deterministic XML documentation and snippet formatting without editor dependencies.
    ''' </summary>
    Friend NotInheritable Class XmlDocumentationFormatting

        ''' <summary>Prevents instantiation of this static formatting helper.</summary>
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Wraps text in a self-closing XML tag with an escaped attribute value.
        ''' </summary>
        Friend Shared Function WrapAttribute(tagName As String, attributeName As String, value As String) As String

            If tagName Is Nothing Then
                Throw New ArgumentNullException(NameOf(tagName))
            End If

            If attributeName Is Nothing Then
                Throw New ArgumentNullException(NameOf(attributeName))
            End If

            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If

            XmlConvert.VerifyNCName(tagName)
            XmlConvert.VerifyNCName(attributeName)
            XmlConvert.VerifyXmlChars(value)
            Return String.Format(CultureInfo.InvariantCulture, "<{0} {1}=""{2}""/>", tagName, attributeName, EscapeXmlAttribute(value))
        End Function

        ''' <summary>
        ''' Wraps text in an XML element while escaping XML-sensitive characters.
        ''' </summary>
        Friend Shared Function WrapElement(tagName As String, value As String) As String

            If tagName Is Nothing Then
                Throw New ArgumentNullException(NameOf(tagName))
            End If

            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If

            XmlConvert.VerifyNCName(tagName)
            XmlConvert.VerifyXmlChars(value)
            Return String.Format(CultureInfo.InvariantCulture, "<{0}>{1}</{0}>", tagName, EscapeXmlText(value))
        End Function

        ''' <summary>
        ''' Returns the minimum leading whitespace among nonempty lines, or zero for empty input.
        ''' </summary>
        Friend Shared Function GetMinimumIndentation(value As String) As Integer
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))

            Dim minimum As Integer = Integer.MaxValue
            Using reader As New StringReader(value)
                Dim currentLine As String = reader.ReadLine()
                While currentLine IsNot Nothing
                    If Not String.IsNullOrWhiteSpace(currentLine) Then
                        Dim indentation As Integer = 0
                        While indentation < currentLine.Length AndAlso Char.IsWhiteSpace(currentLine(indentation))
                            indentation += 1
                        End While

                        minimum = Math.Min(minimum, indentation)
                        If minimum = 0 Then Return 0
                    End If

                    currentLine = reader.ReadLine()
                End While
            End Using

            Return If(minimum = Integer.MaxValue, 0, minimum)
        End Function

        ''' <summary>
        ''' Creates an XML-valid Visual Studio snippet, preserving source text containing CDATA delimiters.
        ''' </summary>
        Friend Shared Function FormatSnippet(template As String, author As String, source As String) As String

            If template Is Nothing Then
                Throw New ArgumentNullException(NameOf(template))
            End If

            If author Is Nothing Then
                Throw New ArgumentNullException(NameOf(author))
            End If

            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            XmlConvert.VerifyXmlChars(author)
            XmlConvert.VerifyXmlChars(source)

            ' Decode only the marker belonging to the template, not matching text supplied by the user.
            Dim normalizedTemplate As String = template.Replace("$cdataend$", ">")
            Dim safeSource As String = source.Replace("]]>", "]]]]><![CDATA[>").Replace(vbCr, "]]>&#xD;<![CDATA[")
            Return String.Format(CultureInfo.InvariantCulture, normalizedTemplate, EscapeXmlText(author), safeSource)
        End Function

        ''' <summary>
        ''' Escapes XML text without losing carriage returns during XML parsing.
        ''' </summary>
        Private Shared Function EscapeXmlText(value As String) As String

            Return SecurityElement.Escape(value).Replace(vbCr, "&#xD;")
        End Function

        ''' <summary>
        ''' Escapes XML attributes, including whitespace normalized by XML parsers.
        ''' </summary>
        Private Shared Function EscapeXmlAttribute(value As String) As String

            Return EscapeXmlText(value).Replace(vbLf, "&#xA;").Replace(vbTab, "&#x9;")
        End Function

        ''' <summary>
        ''' Wraps an existing documentation fragment, preserving valid markup and entities.
        ''' Plain text that is not a well-formed fragment is escaped as text.
        ''' </summary>
        Friend Shared Function WrapDocumentationFragment(tagName As String, value As String) As String

            If tagName Is Nothing Then
                Throw New ArgumentNullException(NameOf(tagName))
            End If

            XmlConvert.VerifyNCName(tagName)
            Return "<" & tagName & ">" & NormalizeDocumentationFragment(value) & "</" & tagName & ">"
        End Function

        ''' <summary>Preserves valid XML fragments without resolving DTDs or external entities.</summary>
        Private Shared Function NormalizeDocumentationFragment(value As String, Optional preserveCarriageReturns As Boolean = True) As String

            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If

            XmlConvert.VerifyXmlChars(value)
            Dim settings As New XmlReaderSettings With {
                .ConformanceLevel = ConformanceLevel.Fragment,
                .DtdProcessing = DtdProcessing.Prohibit,
                .XmlResolver = Nothing
            }
            Try
                Using input As New StringReader(value), reader As XmlReader = XmlReader.Create(input, settings)
                    While reader.Read()
                        If reader.NodeType = XmlNodeType.XmlDeclaration Then
                            Return If(preserveCarriageReturns, EscapeXmlText(value), SecurityElement.Escape(value))
                        End If
                    End While
                End Using
                Return value
            Catch ex As XmlException
                Return If(preserveCarriageReturns, EscapeXmlText(value), SecurityElement.Escape(value))
            End Try
        End Function

        ''' <summary>Splits CRLF, LF, and CR text, retaining empty and trailing lines.</summary>
        Friend Shared Function SplitLines(value As String) As String()

            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If

            Return value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.None)
        End Function

        ''' <summary>Returns the first line ending in the text, defaulting to Windows CRLF.</summary>
        Friend Shared Function GetNewLine(value As String) As String

            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If

            Dim index As Integer = value.IndexOfAny(New Char() {ChrW(13), ChrW(10)})
            If index < 0 Then
                Return vbCrLf
            End If

            If value(index) = ChrW(13) Then
                Return If(index + 1 < value.Length AndAlso value(index + 1) = ChrW(10), vbCrLf, vbCr)
            End If
            Return vbLf
        End Function

        ''' <summary>Returns the last selected character, or the caret for an empty selection.</summary>
        Friend Shared Function GetSelectionEndPosition(startPosition As Integer, endPosition As Integer) As Integer

            If startPosition < 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(startPosition))
            End If

            If endPosition < startPosition Then
                Throw New ArgumentOutOfRangeException(NameOf(endPosition))
            End If

            Return If(endPosition = startPosition, startPosition, endPosition - 1)
        End Function

        ''' <summary>Formats selected source as an escaped, indented XML documentation example.</summary>
        Friend Shared Function FormatCodeExample(source As String, prefix As String, language As String) As String

            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            ValidatePrefix(prefix)

            If language Is Nothing Then
                Throw New ArgumentNullException(NameOf(language))
            End If

            XmlConvert.VerifyXmlChars(source)
            XmlConvert.VerifyXmlChars(language)
            Dim margin As Integer = GetMinimumIndentation(source)
            Dim indentation As New String(" "c, margin)
            Dim newLine As String = GetNewLine(source)
            Dim result As New StringBuilder()
            result.Append(indentation).Append(prefix).Append(" <example> This is a code example.").Append(newLine)
            result.Append(indentation).Append(prefix).Append(" <code language=""").Append(EscapeXmlAttribute(language)).Append(""">").Append(newLine)
            For Each line As String In SplitLines(source)
                Dim content As String = If(String.IsNullOrWhiteSpace(line), String.Empty, line.Substring(margin))
                result.Append(indentation).Append(prefix).Append(" "c).Append(EscapeXmlText(content)).Append(newLine)
            Next
            result.Append(indentation).Append(prefix).Append(" </code>").Append(newLine)
            result.Append(indentation).Append(prefix).Append(" </example>").Append(newLine)
            Return result.ToString()
        End Function

        ''' <summary>Wraps selected documentation in remarks while retaining valid nested markup.</summary>
        Friend Shared Function FormatRemarks(source As String, prefix As String) As String

            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If

            ValidatePrefix(prefix)

            Dim margin As Integer = GetMinimumIndentation(source)
            Dim indentation As New String(" "c, margin)
            Dim newLine As String = GetNewLine(source)
            Dim contents As New List(Of String)()
            For Each line As String In SplitLines(source)
                Dim content As String = If(String.IsNullOrWhiteSpace(line), String.Empty, line.Substring(margin))
                Dim trimmed As String = content.TrimStart()
                If trimmed.StartsWith(prefix, StringComparison.Ordinal) Then
                    content = trimmed.Substring(prefix.Length)
                    If content.StartsWith(" "c, StringComparison.Ordinal) Then content = content.Substring(1)
                End If
                contents.Add(content)
            Next
            Dim fragment As String = NormalizeDocumentationFragment(String.Join(newLine, contents), preserveCarriageReturns:=False)
            Dim result As New StringBuilder()
            result.Append(prefix).Append(" <remarks>").Append(newLine)
            For Each line As String In SplitLines(fragment)
                result.Append(indentation).Append(prefix).Append(" "c).Append(line).Append(newLine)
            Next
            result.Append(indentation).Append(prefix).Append(" </remarks>").Append(newLine)
            Return result.ToString()
        End Function

        ''' <summary>Checks that generated lines use a supported XML documentation prefix.</summary>
        Private Shared Sub ValidatePrefix(prefix As String)

            If prefix Is Nothing Then
                Throw New ArgumentNullException(NameOf(prefix))
            End If

            If prefix <> "'''" AndAlso prefix <> "///" Then
                Throw New ArgumentException("Unsupported documentation comment prefix.", NameOf(prefix))
            End If
        End Sub

        ''' <summary>Writes a snippet using the UTF-8 encoding declared in its XML header.</summary>
        Friend Shared Sub WriteSnippet(path As String, content As String)

            If path Is Nothing Then
                Throw New ArgumentNullException(NameOf(path))
            End If

            If content Is Nothing Then
                Throw New ArgumentNullException(NameOf(content))
            End If
            File.WriteAllText(path, content, New UTF8Encoding(False, True))
        End Sub

    End Class

End Namespace
