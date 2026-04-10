' ***********************************************************************
' Author   : Elektro
' Modified : 09-April-2026
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " XmlDocOperation "

Namespace MyPackage.Core

    ''' <summary>
    ''' Specifies an XML documentation operation to perform on the selected text in the current editor.
    ''' </summary>
    Public Enum XmlDocOperation As Integer

#Region " Section: Code References "

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;see cref="..."/&gt;</c> tag,
        ''' creating a cross-reference link to a type or member in the codebase.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;see cref="MyClass"/&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/cc837134.aspx"/>
        ''' </remarks>
        WrapAsCodeRef

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;paramref name="..."/&gt;</c> tag,
        ''' creating a reference to a parameter of the enclosing method or constructor.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;paramref name="myParam"/&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172662.aspx"/>
        ''' </remarks>
        WrapAsParamRef

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;see langword="..."/&gt;</c> tag,
        ''' creating a reference to a language keyword such as <see langword="Nothing"/>,
        ''' <see langword="True"/> or <see langword="False"/>.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;see langword="Nothing"/&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' This tag is not officially listed in MSDN documentation but is supported by Visual Studio
        ''' and renders the keyword with its standard formatting in IntelliSense tooltips.
        ''' </remarks>
        WrapAsLangRef

#End Region

#Region " Section: Hyperlinks "

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;see href="..."/&gt;</c> tag,
        ''' creating a clickable hyperlink to an external URL in the generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;see href="https://www.example.com"/&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172667.aspx"/>
        ''' </remarks>
        WrapAsHrefLink

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;seealso href="..."/&gt;</c> tag,
        ''' adding a secondary hyperlink reference that appears in the "See Also" section
        ''' of the generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;seealso href="https://www.example.com"/&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172668.aspx"/>
        ''' </remarks>
        WrapAsSeeAlsoLink

#End Region

#Region " Section: Code Block Formatting "

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;c&gt;...&lt;/c&gt;</c> tag,
        ''' formatting it as inline monospace code within a documentation sentence.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;c&gt;myVariable&lt;/c&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172654.aspx"/>
        ''' </remarks>
        WrapAsInlineCode

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;code&gt;...&lt;/code&gt;</c> tag,
        ''' formatting it as a multiline code block in the generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>
        ''' &lt;code&gt;
        ''' Dim x As Integer = 42
        ''' Console.WriteLine(x)
        ''' &lt;/code&gt;
        ''' </code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172655.aspx"/>
        ''' </remarks>
        WrapAsMultilineCode

        ''' <summary>
        ''' Wraps the selected text inside an <c>&lt;example&gt;</c> block containing a description
        ''' and a <c>&lt;code&gt;</c> section, illustrating a usage example in the generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>
        ''' &lt;example&gt;
        ''' This example shows how to use the method.
        ''' &lt;code&gt;
        '''     Dim result As String = MyMethod("hello")
        ''' &lt;/code&gt;
        ''' &lt;/example&gt;
        ''' </code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/9w4cf933.aspx"/>
        ''' </remarks>
        WrapAsCodeExample

#End Region

#Region " Section: Common Formatting "

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;b&gt;...&lt;/b&gt;</c> tag,
        ''' rendering it in bold typeface within IntelliSense tooltips and generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;b&gt;important term&lt;/b&gt;</code>
        ''' </para>
        ''' </summary>
        WrapAsBold

        ''' <summary>
        ''' Wraps the selected text inside an <c>&lt;i&gt;...&lt;/i&gt;</c> tag,
        ''' rendering it in italic typeface within IntelliSense tooltips and generated documentation.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;i&gt;emphasized term&lt;/i&gt;</code>
        ''' </para>
        ''' </summary>
        WrapAsItalic

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;u&gt;...&lt;/u&gt;</c> tag,
        ''' rendering it with an underline in documentation generators that support raw HTML rendering.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;u&gt;underlined term&lt;/u&gt;</code>
        ''' </para>
        ''' </summary>
        WrapAsUnderline

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;para&gt;...&lt;/para&gt;</c> tag,
        ''' grouping it as a distinct paragraph within a documentation comment section.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;para&gt;This is a paragraph of documentation text.&lt;/para&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172660.aspx"/>
        ''' </remarks>
        WrapAsParagraph

        ''' <summary>
        ''' Wraps the selected text inside a <c>&lt;remarks&gt;...&lt;/remarks&gt;</c> tag,
        ''' adding supplementary information that extends the main <c>&lt;summary&gt;</c> description.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;remarks&gt;This method is not thread-safe.&lt;/remarks&gt;</code>
        ''' </para>
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/ms172664.aspx"/>
        ''' </remarks>
        WrapAsRemarks

        ''' <summary>
        ''' Inserts a horizontal XML separator line at the cursor position in the current editor,
        ''' used to visually divide sections within a documentation comment block.
        ''' <para>
        ''' Example output:
        ''' <code>&lt;!-- ==================== --&gt;</code>
        ''' </para>
        ''' </summary>
        InsertSeparatorLine

#End Region

#Region " Section: Snippet Generation "

        ''' <summary>
        ''' Formats the currently selected text as a Visual Studio <c>.snippet</c> XML structure,
        ''' wrapping it inside the required snippet schema ready to be saved as a reusable code snippet.
        ''' <para>
        ''' Example output:
        ''' <code>
        ''' &lt;?xml version="1.0" encoding="utf-8"?&gt;
        ''' &lt;CodeSnippets&gt;
        '''   &lt;CodeSnippet Format="1.0.0"&gt;
        '''     &lt;Header&gt;
        '''       &lt;Title&gt;My Snippet&lt;/Title&gt;
        '''     &lt;/Header&gt;
        '''     &lt;Snippet&gt;
        '''       &lt;Code Language="VB"&gt;
        '''         &lt;![CDATA[' selected code here]]&gt;
        '''       &lt;/Code&gt;
        '''     &lt;/Snippet&gt;
        '''   &lt;/CodeSnippet&gt;
        ''' &lt;/CodeSnippets&gt;
        ''' </code>
        ''' </para>
        ''' </summary>
        GenerateSnippet

#End Region

#Region " Section: Editor Operations "

        ''' <summary>
        ''' Collapses all expanded XML documentation comment blocks in the current editor,
        ''' reducing visual noise and making the source code easier to navigate.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="http://www.helixoft.com/blog/collapse-all-xml-comments-in-vb-net-or-c.html"/>
        ''' </remarks>
        CollapseXmlComments

        ''' <summary>
        ''' Expands all collapsed XML documentation comment blocks in the current editor,
        ''' making the full documentation content visible alongside the source code.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="http://www.helixoft.com/blog/collapse-all-xml-comments-in-vb-net-or-c.html"/>
        ''' </remarks>
        ExpandXmlComments

        ''' <summary>
        ''' Permanently removes all XML documentation comment blocks from the current editor,
        ''' stripping every <c>'''</c>-prefixed line in Visual Basic source files,
        ''' and every <c>///</c>-prefixed line in C# source files.
        ''' </summary>
        DeleteXmlComments

#End Region

    End Enum

End Namespace

#End Region