' ***********************************************************************
' Author   : Elektro
' Modified : 09-April-2026
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Docs Constants "

Namespace MyPackage.Core

    Partial Public NotInheritable Class DocsConstants

        ''' <summary>
        ''' The character sequence that identifies an XML documentation comment line
        ''' in Visual Basic .NET source files.
        ''' </summary>
        Friend Const XmlCommentCharsVB As String = "'''"

        ''' <summary>
        ''' The character sequence that identifies an XML documentation comment line
        ''' in C# source files.
        ''' </summary>
        Friend Const XmlCommentCharsCS As String = "///"

        ''' <summary>
        ''' The Visual Studio <c>.snippet</c> XML template used to generate snippet files
        ''' for Visual Basic .NET, formatted with two placeholders:
        ''' <c>{0}</c> for the author name and <c>{1}</c> for the selected code block.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' The template targets the Visual Studio Code Snippet schema version 1.0.0.
        ''' <para>
        ''' <see href="https://learn.microsoft.com/en-us/visualstudio/ide/code-snippets-schema-reference"/>
        ''' </para>
        ''' </remarks>
        Public Shared ReadOnly SnippetTemplateFormatVB As String =
<a><![CDATA[<?xml version="1.0" encoding="utf-8"?>
        <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
        <CodeSnippet Format="1.0.0">

        <Header>
          <Title>Title</Title>
          <Description>Description</Description>
          <Author>{0}</Author>
        </Header>

        <Snippet>

        <References>
          <Reference>
            <Assembly>System.dll</Assembly>
          </Reference>
        </References>

        <Imports>
          <Import>
            <Namespace>System</Namespace>
          </Import>
        </Imports>

        <Declarations>
          <Literal Editable="false">
            <ID>CDATAEnd</ID>
            <ToolTip>Closing CDATA sequence used to embed XML-unsafe (illegal) characters in the snippet body if needed.</ToolTip>
            <Default>&gt;</Default>
          </Literal>
        </Declarations>

        <Code Language="VB"><![CDATA[

        {1}

        ]]$cdataend$</a>
          </Snippet>
         </CodeSnippet>
        </CodeSnippets>]]></a>.Value

        ''' <summary>
        ''' The Visual Studio <c>.snippet</c> XML template used to generate snippet files
        ''' for C#, formatted with two placeholders:
        ''' <c>{0}</c> for the author name and <c>{1}</c> for the selected code block.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' The template targets the Visual Studio Code Snippet schema version 1.0.0.
        ''' <para>
        ''' <see href="https://learn.microsoft.com/en-us/visualstudio/ide/code-snippets-schema-reference"/>
        ''' </para>
        ''' </remarks>
        Public Shared ReadOnly SnippetTemplateFormatCS As String =
<a><![CDATA[<?xml version="1.0" encoding="utf-8"?>
        <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
        <CodeSnippet Format="1.0.0">

        <Header>
          <Title>Title</Title>
          <Description>Description</Description>
          <Author>{0}</Author>
        </Header>

        <Snippet>

        <References>
          <Reference>
            <Assembly>System.dll</Assembly>
          </Reference>
        </References>

        <Declarations>
          <Literal Editable="false">
            <ID>CDATAEnd</ID>
            <ToolTip>Closing CDATA sequence used to embed XML-unsafe (illegal) characters in the snippet body if needed.</ToolTip>
            <Default>&gt;</Default>
          </Literal>
        </Declarations>

        <Code Language="CSharp"><![CDATA[

        {1}

        ]]$cdataend$</a>
          </Snippet>
         </CodeSnippet>
        </CodeSnippets>]]></a>.Value

    End Class

End Namespace

#End Region
