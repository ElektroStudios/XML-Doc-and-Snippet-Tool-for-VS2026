Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports Snippet_Tool_2026.MyPackage
Imports System.Xml.Linq
Imports Xunit
Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.PackageSettings

''' <summary>Validates static command wiring and compiled command metadata.</summary>
Public Class CommandContractTests

    ''' <summary>
    ''' Maps every command across the VSCT file, registration, callbacks, and operation enum.
    ''' </summary>
    Private Shared Iterator Function GetCommandDefinitions() As IEnumerable(Of String())
        Dim entries As String() = {
            "cmdidCRef|CodeRef|WrapAsCodeRef|WrapAsCodeRef|0100",
            "cmdidParamRef|ParamRef|WrapAsParamRef|WrapAsParamRef|0101",
            "cmdidLangRef|LangRef|WrapAsLangRef|WrapAsLangRef|0102",
            "cmdidLink|HrefLink|WrapAsHrefLink|WrapAsHrefLink|0103",
            "cmdidLinkAlter|SeeAlsoLink|WrapAsSeeAlsoLink|WrapAsSeeAlsoLink|0104",
            "cmdidBold|Bold|WrapAsBold|WrapAsBold|0105",
            "cmdidItalic|Italic|WrapAsItalic|WrapAsItalic|0106",
            "cmdidUnderline|Underline|WrapAsUnderline|WrapAsUnderline|0107",
            "cmdidParagraph|Paragraph|WrapAsParagraph|WrapAsParagraph|0108",
            "cmdidSinglelineCode|InlineCode|WrapAsInlineCode|WrapAsInlineCode|0109",
            "cmdidMultilineCode|MultilineCode|WrapAsMultilineCode|WrapAsMultilineCode|010A",
            "cmdidCodeExample|CodeExample|WrapAsCodeExample|WrapAsCodeExample|010B",
            "cmdidRemarks|Remarks|WrapAsRemarks|WrapAsRemarks|010C",
            "cmdidSnippet|GenerateSnippet|GenerateSnippet|GenerateSnippet|010D",
            "cmdidSeparator|InsertSeparatorLine|InsertSeparatorLine|InsertSeparatorLine|010E",
            "cmdidCollapse|CollapseXmlComments|CollapseXmlComments|CollapseXmlComments|010F",
            "cmdidExpand|ExpandXmlComments|ExpandXmlComments|ExpandXmlComments|0110",
            "cmdidDelete|DeleteXmlComments|DeleteXmlComments|DeleteXmlComments|0111"
        }
        For Each entry As String In entries
            Dim parts As String() = entry.Split("|"c)
            Yield parts
        Next
    End Function

    ''' <summary>Supplies only the symbol, constant name, and expected numeric identifier.</summary>
    Public Shared Iterator Function CommandIdentifierCases() As IEnumerable(Of TheoryDataRow(Of String, String, String))
        For Each definition As String() In GetCommandDefinitions()
            Yield New TheoryDataRow(Of String, String, String)(definition(0), definition(1), definition(4))
        Next
    End Function

    ''' <summary>Supplies only the fields used to validate command registration and callbacks.</summary>
    Public Shared Iterator Function CommandRegistrationCases() As IEnumerable(Of TheoryDataRow(Of String, String, String, String))
        For Each definition As String() In GetCommandDefinitions()
            Yield New TheoryDataRow(Of String, String, String, String)(definition(0), definition(1), definition(2), definition(3))
        Next
    End Function

    ''' <summary>
    ''' Returns a command constant compiled from the production VB.NET source.
    ''' </summary>
    Private Shared Function GetProductionCommandId(fieldName As String) As Integer
        Select Case fieldName
            Case "CodeRef"
                Return CInt(CommandIds.CodeRef)
            Case "ParamRef"
                Return CInt(CommandIds.ParamRef)
            Case "LangRef"
                Return CInt(CommandIds.LangRef)
            Case "HrefLink"
                Return CInt(CommandIds.HrefLink)
            Case "SeeAlsoLink"
                Return CInt(CommandIds.SeeAlsoLink)
            Case "Bold"
                Return CInt(CommandIds.Bold)
            Case "Italic"
                Return CInt(CommandIds.Italic)
            Case "Underline"
                Return CInt(CommandIds.Underline)
            Case "Paragraph"
                Return CInt(CommandIds.Paragraph)
            Case "InlineCode"
                Return CInt(CommandIds.InlineCode)
            Case "MultilineCode"
                Return CInt(CommandIds.MultilineCode)
            Case "CodeExample"
                Return CInt(CommandIds.CodeExample)
            Case "Remarks"
                Return CInt(CommandIds.Remarks)
            Case "GenerateSnippet"
                Return CInt(CommandIds.GenerateSnippet)
            Case "InsertSeparatorLine"
                Return CInt(CommandIds.InsertSeparatorLine)
            Case "CollapseXmlComments"
                Return CInt(CommandIds.CollapseXmlComments)
            Case "ExpandXmlComments"
                Return CInt(CommandIds.ExpandXmlComments)
            Case "DeleteXmlComments"
                Return CInt(CommandIds.DeleteXmlComments)
            Case Else : Throw New ArgumentOutOfRangeException(NameOf(fieldName), fieldName, "Unknown command constant.")
        End Select
    End Function

    ''' <summary>
    ''' Validates the numeric VSCT identifier against the constant compiled from production source.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(CommandIdentifierCases))>
    Public Sub VsctIdentifiersMatchCompiledConstants(symbol As String, fieldName As String, hexadecimal As String)
        Dim document As XDocument = TestResources.ReadXml("Snippet Tool.vsct")
        Dim elements As XElement() = document.Descendants().Where(Function(e As XElement) e.Name.LocalName = "IDSymbol" AndAlso CStr(e.Attribute("name")) = symbol).ToArray()
        Assert.Single(elements)
        Dim value As Integer = Integer.Parse(CStr(elements(0).Attribute("value")).Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
        Dim expected As Integer = Integer.Parse(hexadecimal, NumberStyles.HexNumber, CultureInfo.InvariantCulture)
        Assert.Equal(expected, value)
        Dim actual As Integer = GetProductionCommandId(fieldName)
        Assert.Equal(expected, actual)
        Dim groupSymbol As XElement = elements(0).Parent
        Dim groupName As String = CStr(groupSymbol.Attribute("name"))
        Assert.True(groupName.StartsWith("guidCmdGroup", StringComparison.Ordinal))
        Dim groupField As FieldInfo = GetType(Guids).GetField(groupName.Substring(4), BindingFlags.Static Or BindingFlags.NonPublic)
        Assert.NotNull(groupField)
        Assert.Equal(New Guid(CStr(groupSymbol.Attribute("value"))), DirectCast(groupField.GetValue(Nothing), Guid))
        Dim button As XElement = document.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Button" AndAlso CStr(e.Attribute("id")) = symbol)
        Assert.Equal(groupName, CStr(button.Attribute("guid")))

    End Sub

    ''' <summary>
    ''' Validates that every registered command has a matching callback and query-status handler.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(CommandRegistrationCases))>
    Public Sub CommandHasExpectedRegistrationAndHandlers(symbol As String, fieldName As String, callbackName As String, operationName As String)
        Dim packageText As String = TestResources.ReadFixture("Snippet_Tool_2026Package.vb")
        Dim callbacks As String = TestResources.ReadFixture("Command Callbacks.vb")
        Dim handlers As String = TestResources.ReadFixture("Command EventHandlers.vb")
        Dim command As String = "Cmd" & callbackName
        Dim registration As String = packageText.Split(New Char() {ChrW(10)}).
            Single(Function(line As String) line.TrimStart().StartsWith("Me." & command & " = Snippet_Tool_2026Package.CreateAndRegister(", StringComparison.Ordinal))
        Assert.Contains("CreateAndRegister(", registration, StringComparison.Ordinal)
        Assert.Contains("AddressOf Me." & command & "_Callback,", registration, StringComparison.Ordinal)
        Assert.Contains("CommandIds." & fieldName & ")", registration, StringComparison.Ordinal)
        Dim vsct As XDocument = TestResources.ReadXml("Snippet Tool.vsct")
        Dim groupName As String = CStr(vsct.Descendants().Single(Function(e As XElement) e.Name.LocalName = "IDSymbol" AndAlso CStr(e.Attribute("name")) = symbol).Parent.Attribute("name"))
        Assert.Contains("Guids." & groupName.Substring(4) & ",", registration, StringComparison.Ordinal)
        Dim callbackStart As Integer = callbacks.IndexOf("Sub " & command & "_Callback(", StringComparison.Ordinal)
        Assert.True(callbackStart >= 0)
        Dim callbackEnd As Integer = callbacks.IndexOf("End Sub", callbackStart, StringComparison.Ordinal)
        Assert.True(callbackEnd > callbackStart)
        Dim callbackBody As String = callbacks.Substring(callbackStart, callbackEnd - callbackStart)
        Assert.Contains("XmlDocOperation." & operationName & ")", callbackBody, StringComparison.Ordinal)
        Assert.Contains("Sub " & command & "_BeforeQueryStatus(", handlers, StringComparison.Ordinal)
        Assert.Contains("Handles " & command & ".BeforeQueryStatus", handlers, StringComparison.Ordinal)
    End Sub

    ''' <summary>
    ''' Verifies that every command identifier in VSCT is unique.
    ''' </summary>
    <Fact>
    Public Sub VsctCommandIdentifiersAreUnique()
        Dim document As XDocument = TestResources.ReadXml("Snippet Tool.vsct")
        Dim commands As XElement() = document.Descendants().Where(Function(e As XElement)
            Return e.Name.LocalName = "IDSymbol" AndAlso CStr(e.Attribute("name")).StartsWith("cmdid", StringComparison.Ordinal)
        End Function).ToArray()
        Assert.Equal(18, commands.Length)
        Assert.Equal(commands.Length, commands.Select(Function(e As XElement) CStr(e.Attribute("name"))).Distinct(StringComparer.Ordinal).Count())
        Assert.Equal(commands.Length, commands.Select(Function(e As XElement) CStr(e.Attribute("value"))).Distinct(StringComparer.OrdinalIgnoreCase).Count())
    End Sub

    ''' <summary>
    ''' Ensures all serialized operation values and names remain stable.
    ''' </summary>
    <Fact>
    Public Sub OperationsHaveStableNumericValues()
        Dim expected As String() = {
            "WrapAsCodeRef", "WrapAsParamRef", "WrapAsLangRef", "WrapAsHrefLink", "WrapAsSeeAlsoLink",
            "WrapAsInlineCode", "WrapAsMultilineCode", "WrapAsCodeExample", "WrapAsBold", "WrapAsItalic",
            "WrapAsUnderline", "WrapAsParagraph", "WrapAsRemarks", "InsertSeparatorLine", "GenerateSnippet",
            "CollapseXmlComments", "ExpandXmlComments", "DeleteXmlComments"
        }
        Assert.Equal(expected.Length, [Enum].GetNames(GetType(XmlDocOperation)).Length)
        For i As Integer = 0 To expected.Length - 1
            Dim operation As XmlDocOperation = DirectCast([Enum].Parse(GetType(XmlDocOperation), expected(i)), XmlDocOperation)
            Assert.Equal(i, CInt(operation))
        Next
    End Sub

    ''' <summary>
    ''' Confirms that the two recognized documentation languages retain their comment prefixes.
    ''' </summary>
    <Fact>
    Public Sub LanguageCommentPrefixesAreStable()
        Assert.Equal("'''", CStr(GetType(DocsConstants).GetField(NameOf(DocsConstants.XmlCommentCharsVB), BindingFlags.NonPublic Or BindingFlags.Static).GetRawConstantValue()))
        Assert.Equal("///", CStr(GetType(DocsConstants).GetField(NameOf(DocsConstants.XmlCommentCharsCS), BindingFlags.NonPublic Or BindingFlags.Static).GetRawConstantValue()))
    End Sub

    ''' <summary>
    ''' Ensures that the package registration contains a recovery path when DTE arrives late.
    ''' </summary>
    <Fact>
    Public Sub DeferredDteInitializationRegistersMenuHandlers()
        Dim source As String = TestResources.ReadFixture("Snippet_Tool_2026Package.vb")
        Dim callbackStart As Integer = source.IndexOf("DteInitializer = New DteInitializer(", StringComparison.Ordinal)
        Assert.True(callbackStart >= 0)
        Dim callbackEnd As Integer = source.IndexOf("Logger.Debug(""DteInitializer registered successfully"")", callbackStart, StringComparison.Ordinal)
        Assert.True(callbackEnd > callbackStart)
        Dim fallback As String = source.Substring(callbackStart, callbackEnd - callbackStart)
        Assert.Contains("InitializeMenuHandlersAsync", fallback, StringComparison.Ordinal)
    End Sub
End Class
