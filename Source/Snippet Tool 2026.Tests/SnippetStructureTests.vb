Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Xml.Linq
Imports Xunit
Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.Helpers

Public Class SnippetStructureTests

    ''' <summary>
    ''' Supplies each supported snippet language.
    ''' </summary>
    Public Shared Iterator Function LanguageCases() As IEnumerable(Of TheoryDataRow(Of String, String))
        Yield New TheoryDataRow(Of String, String)("VB", DocsConstants.SnippetTemplateFormatVB)
        Yield New TheoryDataRow(Of String, String)("CSharp", DocsConstants.SnippetTemplateFormatCS)
    End Function

    ''' <summary>
    ''' Ensures the generated snippet contains required schema elements and correct metadata.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(LanguageCases))>
    Public Sub SnippetContainsRequiredSchemaNodes(language As String, template As String)
        Dim document As XDocument = XDocument.Parse(XmlDocumentationFormatting.FormatSnippet(template, "Tester", "x"))
        Dim ns As XNamespace = "http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet"
        Assert.Equal(ns + "CodeSnippets", document.Root.Name)
        Dim item As XElement = document.Root.Element(ns + "CodeSnippet")
        Assert.NotNull(item)
        Assert.Equal("1.0.0", CStr(item.Attribute("Format")))
        Assert.Equal("Title", item.Element(ns + "Header").Element(ns + "Title").Value)
        Assert.Equal("Description", item.Element(ns + "Header").Element(ns + "Description").Value)
        Assert.Equal("Tester", item.Element(ns + "Header").Element(ns + "Author").Value)
        Dim code As XElement = item.Element(ns + "Snippet").Element(ns + "Code")
        Assert.NotNull(code)
        Assert.Equal(language, CStr(code.Attribute("Language")))
    End Sub

    ''' <summary>
    ''' Confirms that the snippet template has the original author and source placeholders.
    ''' </summary>
    <Theory>
    <MemberData(NameOf(LanguageCases))>
    Public Sub SnippetTemplateHasRequiredFormatMarkers(language As String, template As String)
        Assert.Contains("<Author>{0}</Author>", template, StringComparison.Ordinal)
        Assert.Contains("{1}", template, StringComparison.Ordinal)
        Assert.Contains("$cdataend$", template, StringComparison.Ordinal)
        Assert.Contains("<Code Language=""" & language & """>", template, StringComparison.Ordinal)
    End Sub

    ''' <summary>
    ''' Ensures that the compiled package targets the correct .NET Framework and is a VSIX package.
    ''' </summary>
    <Fact>
    Public Sub ProductionProjectRetainsVsixConfiguration()
        Dim project As XDocument = TestResources.ReadXml("Snippet Tool 2026.vbproj")
        Dim names As String() = project.Descendants().Select(Function(e As XElement) e.Name.LocalName).ToArray()
        Assert.True(names.Contains("TargetFrameworkVersion"))
        Assert.Equal("v4.8", project.Descendants().First(Function(e As XElement) e.Name.LocalName = "TargetFrameworkVersion").Value)
        Assert.Equal("true", project.Descendants().First(Function(e As XElement) e.Name.LocalName = "IncludeAssemblyInVSIXContainer").Value)
    End Sub

    ''' <summary>
    ''' Ensures the VSIX package asset and Visual Studio installation target are present.
    ''' </summary>
    <Fact>
    Public Sub VsixManifestDeclaresPackageAndVisualStudioTarget()
        Dim manifest As XDocument = TestResources.ReadXml("source.extension.vsixmanifest")
        Dim asset As XElement = manifest.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Asset")
        Assert.Equal("Microsoft.VisualStudio.VsPackage", CStr(asset.Attribute("Type")))
        Dim installation As XElement = manifest.Descendants().Single(Function(e As XElement) e.Name.LocalName = "InstallationTarget")
        Assert.Equal("Microsoft.VisualStudio.Community", CStr(installation.Attribute("Id")))
        Assert.NotNull(installation.Element(installation.Name.Namespace + "ProductArchitecture"))
    End Sub

    ''' <summary>
    ''' Ensures the Visual Studio command table has references to the declared command symbols.
    ''' </summary>
    <Fact>
    Public Sub VsctIsWellFormedAndDeclaresCommandGroups()
        Dim vsct As XDocument = TestResources.ReadXml("Snippet Tool.vsct")
        Dim groups As XElement() = vsct.Descendants().Where(Function(e As XElement) e.Name.LocalName = "GuidSymbol" AndAlso
            CStr(e.Attribute("name")).StartsWith("guidCmdGroup", StringComparison.Ordinal)).ToArray()
        Assert.Equal(6, groups.Length)
        For Each group As XElement In groups
            Assert.True(group.Elements().Any(Function(e As XElement) e.Name.LocalName = "IDSymbol" AndAlso
                CStr(e.Attribute("name")).StartsWith("cmdid", StringComparison.Ordinal)))
        Next
    End Sub
    ''' <summary>Prevents mismatched VSIX and assembly versions.</summary>
    <Fact>
    Public Sub VersionMetadataIsConsistent()
        Dim manifest As XDocument = TestResources.ReadXml("source.extension.vsixmanifest")
        Dim version As String = CStr(manifest.Descendants().Single(Function(e As XElement) e.Name.LocalName = "Identity").Attribute("Version"))
        Dim assembly As String = TestResources.ReadFixture("AssemblyInfo.vb")
        Assert.Contains("AssemblyVersion(""" & version & ".0"")", assembly, StringComparison.Ordinal)
        Assert.Contains("AssemblyFileVersion(""" & version & ".0"")", assembly, StringComparison.Ordinal)
    End Sub
End Class
