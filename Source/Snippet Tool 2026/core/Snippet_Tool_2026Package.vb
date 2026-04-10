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
Imports System.ComponentModel.Design
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Threading

Imports EnvDTE80

Imports Microsoft.VisualBasic

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Threading

Imports Snippet_Tool_2026.MyPackage
Imports Snippet_Tool_2026.MyPackage.Core
Imports Snippet_Tool_2026.MyPackage.Helpers
Imports Snippet_Tool_2026.MyPackage.PackageSettings
Imports Snippet_Tool_2026.MyPackage.UserInterface

Imports stdole

Imports Task = System.Threading.Tasks.Task

#End Region

''' <summary>
''' This is the class that implements the package exposed by this assembly.
''' </summary>
''' 
''' <remarks>
''' <para>
''' The minimum requirement for a class to be considered a valid package for Visual Studio
''' Is to implement the IVsPackage interface And register itself with the shell.
''' This package uses the helper classes defined inside the Managed Package Framework (MPF)
''' to do it: it derives from the Package Class that provides the implementation Of the 
''' IVsPackage interface And uses the registration attributes defined in the framework to 
''' register itself And its components with the shell. These attributes tell the pkgdef creation
''' utility what data to put into .pkgdef file.
''' </para>
''' <para>
''' To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
''' </para>
''' </remarks>
<InstalledProductRegistration("#110", "#112", "1.6", IconResourceID:=400)>
<ProvideMenuResource("Menus.ctmenu", 1)>
<ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)>
<Guid(Guids.PackageGuid)>
<ProvideOptionPage(GetType(SnippetGenerationPageGrid), "Snippet Tool", "Snippet Generation", 0, 0, True)>
<ProvideOptionPage(GetType(OtherOptionsPageGrid), "Snippet Tool", "Other Options", 0, 0, True)>
<PackageRegistration(UseManagedResourcesOnly:=True, AllowsBackgroundLoading:=True)>
Public NotInheritable Class Snippet_Tool_2026Package : Inherits AsyncPackage

#Region " Fields "

#Region " DTE "

    ''' <summary>
    ''' The shared <see cref="EnvDTE80.DTE2"/> instance used across the package.
    ''' </summary>
    Friend Shared Dte As DTE2

    ''' <summary>
    ''' The <see cref="Helpers.DteInitializer"/> instance used as a fallback mechanism
    ''' to initialize <see cref="Snippet_Tool_2026Package.Dte"/> when the IDE is not yet fully loaded.
    ''' </summary>
    Friend DteInitializer As DteInitializer

#End Region

#Region " Commands "

#Region " Section: Code References "

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;see cref="..."/&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsCodeRef As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;paramref name="..."/&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsParamRef As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;see langword="..."/&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsLangRef As OleMenuCommand

#End Region

#Region " Section: Hyperlinks "

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;see href="..."/&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsHrefLink As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;seealso href="..."/&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsSeeAlsoLink As OleMenuCommand

#End Region

#Region " Section: Code Block Formatting "

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;c&gt;...&lt;/c&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsInlineCode As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;code&gt;...&lt;/code&gt;</c> block.
    ''' </summary>
    Private WithEvents CmdWrapAsMultilineCode As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text inside a full <c>&lt;example&gt;</c> block.
    ''' </summary>
    Private WithEvents CmdWrapAsCodeExample As OleMenuCommand

#End Region

#Region " Section: Common Formatting "

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;b&gt;...&lt;/b&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsBold As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as an <c>&lt;i&gt;...&lt;/i&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsItalic As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;u&gt;...&lt;/u&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsUnderline As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;para&gt;...&lt;/para&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsParagraph As OleMenuCommand

    ''' <summary>
    ''' The command that wraps the selected text as a <c>&lt;remarks&gt;...&lt;/remarks&gt;</c> tag.
    ''' </summary>
    Private WithEvents CmdWrapAsRemarks As OleMenuCommand

    ''' <summary>
    ''' The command that inserts an XML documentation separator line at the cursor position.
    ''' </summary>
    Private WithEvents CmdInsertSeparatorLine As OleMenuCommand

#End Region

#Region " Section: Snippet Generation "

    ''' <summary>
    ''' The command that formats the selected text as a Visual Studio <c>.snippet</c> XML structure.
    ''' </summary>
    Private WithEvents CmdGenerateSnippet As OleMenuCommand

#End Region

#Region " Section: Editor Operations "

    ''' <summary>
    ''' The command that collapses all expanded XML documentation comment blocks in the current editor.
    ''' </summary>
    Private WithEvents CmdCollapseXmlComments As OleMenuCommand

    ''' <summary>
    ''' The command that expands all collapsed XML documentation comment blocks in the current editor.
    ''' </summary>
    Private WithEvents CmdExpandXmlComments As OleMenuCommand

    ''' <summary>
    ''' The command that permanently removes all XML documentation comment blocks from the current editor.
    ''' </summary>
    Private WithEvents CmdDeleteXmlComments As OleMenuCommand

#End Region

#End Region

#End Region

#Region " Constructors "

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Main"/> class.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' No Visual Studio services should be consumed here, as the package is not yet
    ''' sited inside the Visual Studio environment at construction time.
    ''' </remarks>
    Public Sub New()
    End Sub

#End Region

#Region " Initialize Methods "

    ''' <summary>
    ''' Performs asynchronous initialization of the package.
    ''' Switches to the UI thread before accessing Visual Studio services.
    ''' </summary>
    ''' 
    ''' <param name="cancellationToken">
    ''' A <see cref="CancellationToken"/> that can be used to cancel the initialization.
    ''' </param>
    ''' <param name="progress">
    ''' A provider for progress updates during initialization.
    ''' </param>
    Protected Overrides Async Function InitializeAsync(cancellationToken As CancellationToken,
                                                       progress As IProgress(Of ServiceProgressData)) As Task

        Logger.Initialize("Snippet_Tool_2026.txt")
        Logger.Info("InitializeAsync — START")

        Try
            Logger.Debug("Calling MyBase.InitializeAsync...")
            Await MyBase.InitializeAsync(cancellationToken, progress)
            Logger.Info("MyBase.InitializeAsync completed")

            Logger.Debug("Switching to main thread...")
            Await Me.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken)
            Logger.Info("Now running on main thread")

            Logger.Debug("Calling InitializeDteAsync...")
            Await Me.InitializeDteAsync(cancellationToken)
            Logger.Info($"InitializeDteAsync completed — DTE IsNot Nothing: {Snippet_Tool_2026Package.Dte IsNot Nothing}")

            If Snippet_Tool_2026Package.Dte IsNot Nothing Then
                Logger.Info("DTE available — calling InitializeMenuHandlersAsync...")
                Await Me.InitializeMenuHandlersAsync()
                Logger.Info("InitializeMenuHandlersAsync completed")
            Else
                Logger.Warn("DTE is Nothing — InitializeMenuHandlersAsync skipped")
            End If

            Logger.Info("InitializeAsync — END")

        Catch ex As OperationCanceledException
            Logger.Warn($"InitializeAsync cancelled: {ex.Message}")
            Throw

        Catch ex As Exception
            Logger.Error("InitializeAsync failed with unexpected exception", ex)
            Throw

        End Try

    End Function

    ''' <summary>
    ''' Attempts to obtain the <see cref="EnvDTE80.DTE2"/> automation model instance asynchronously.
    ''' Falls back to a <see cref="DteInitializer"/> shell-event listener if the IDE
    ''' is not yet fully loaded at the time of the call.
    ''' </summary>
    ''' 
    ''' <param name="cancellationToken">
    ''' A <see cref="CancellationToken"/> that can be used to cancel the operation.
    ''' </param>
    ''' 
    ''' <remarks>
    ''' <see href="http://www.mztools.com/articles/2013/MZ2013029.aspx"/>
    ''' </remarks>
    Private Async Function InitializeDteAsync(cancellationToken As CancellationToken) As Task

        Logger.Debug("InitializeDteAsync — START")

        Logger.Debug("Switching to main thread...")
        Await Me.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken)
        Logger.Debug("Now on main thread")

        Logger.Debug("Requesting DTE2 service...")
        Snippet_Tool_2026Package.Dte = TryCast(Await Me.GetServiceAsync(GetType(SDTE)), DTE2)

        If Snippet_Tool_2026Package.Dte Is Nothing Then
            Logger.Warn("DTE2 is Nothing — IDE not fully initialized yet, registering shell-event fallback")

            Logger.Debug("Requesting IVsShell service...")
            Dim shellService As IVsShell =
                TryCast(Await Me.GetServiceAsync(GetType(SVsShell)), IVsShell)

            If shellService IsNot Nothing Then
                Logger.Info("IVsShell obtained — creating DteInitializer fallback")
                Me.DteInitializer = New DteInitializer(
                    shellService,
                    Sub()
                        Logger.Info("DteInitializer callback fired — retrying InitializeDteAsync")
                        Dim pendingInit As JoinableTask = Me.JoinableTaskFactory.RunAsync(
                            Function() Me.InitializeDteAsync(cancellationToken))
                    End Sub)
                Logger.Debug("DteInitializer registered successfully")
            Else
                Logger.Error("IVsShell service is Nothing — cannot register DteInitializer fallback")
            End If

        Else
            Logger.Info($"DTE2 obtained successfully — FullName: {Snippet_Tool_2026Package.Dte.FullName}")
            Me.DteInitializer = Nothing
            Logger.Debug("DteInitializer cleared (not needed)")
        End If

        Logger.Debug("InitializeDteAsync — END")

    End Function

    ''' <summary>
    ''' Registers all menu command handlers with the <see cref="OleMenuCommandService"/>.
    ''' </summary>
    Private Async Function InitializeMenuHandlersAsync() As Task

        Logger.Debug("InitializeMenuHandlersAsync — START")

        Logger.Debug("Switching to main thread...")
        Await Me.JoinableTaskFactory.SwitchToMainThreadAsync()
        Logger.Debug("Now on main thread")

        Logger.Debug("Requesting OleMenuCommandService...")
        Dim mcs As OleMenuCommandService =
            TryCast(Await Me.GetServiceAsync(GetType(IMenuCommandService)), OleMenuCommandService)

        If mcs Is Nothing Then
            Logger.Warn("OleMenuCommandService not available — retrying after 500ms delay...")
            Await Task.Delay(500)
            mcs = TryCast(Await Me.GetServiceAsync(GetType(IMenuCommandService)), OleMenuCommandService)

            If mcs Is Nothing Then
                Logger.Error("OleMenuCommandService is definitively unavailable — commands cannot be registered, aborting")
                Return
            End If

            Logger.Info("OleMenuCommandService obtained on retry")
        Else
            Logger.Info("OleMenuCommandService obtained on first attempt")
        End If

        ' Section: references 
        Logger.Debug("Registering References section commands...")
        Me.CmdWrapAsCodeRef = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsCodeRef_Callback, Guids.CmdGroupCodeReferences, CommandIds.CodeRef)
        Me.CmdWrapAsParamRef = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsParamRef_Callback, Guids.CmdGroupCodeReferences, CommandIds.ParamRef)
        Me.CmdWrapAsLangRef = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsLangRef_Callback, Guids.CmdGroupCodeReferences, CommandIds.LangRef)
        Me.CmdWrapAsHrefLink = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsHrefLink_Callback, Guids.CmdGroupHyperlinks, CommandIds.HrefLink)
        Me.CmdWrapAsSeeAlsoLink = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsSeeAlsoLink_Callback, Guids.CmdGroupHyperlinks, CommandIds.SeeAlsoLink)
        Logger.Info("References commands registered (CodeRef, ParamRef, LangRef, HrefLink, SeeAlsoLink)")

        ' Section: code formatting
        Logger.Debug("Registering Code Formatting section commands...")
        Me.CmdWrapAsBold = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsBold_Callback, Guids.CmdGroupCommonFormatting, CommandIds.Bold)
        Me.CmdWrapAsItalic = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsItalic_Callback, Guids.CmdGroupCommonFormatting, CommandIds.Italic)
        Me.CmdWrapAsUnderline = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsUnderline_Callback, Guids.CmdGroupCommonFormatting, CommandIds.Underline)
        Me.CmdWrapAsParagraph = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsParagraph_Callback, Guids.CmdGroupCommonFormatting, CommandIds.Paragraph)
        Me.CmdWrapAsInlineCode = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsInlineCode_Callback, Guids.CmdGroupCodeBlockFormatting, CommandIds.InlineCode)
        Me.CmdWrapAsMultilineCode = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsMultilineCode_Callback, Guids.CmdGroupCodeBlockFormatting, CommandIds.MultilineCode)
        Me.CmdWrapAsCodeExample = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsCodeExample_Callback, Guids.CmdGroupCodeBlockFormatting, CommandIds.CodeExample)
        Me.CmdWrapAsRemarks = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdWrapAsRemarks_Callback, Guids.CmdGroupCommonFormatting, CommandIds.Remarks)
        Me.CmdGenerateSnippet = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdGenerateSnippet_Callback, Guids.CmdGroupSnippetGeneration, CommandIds.GenerateSnippet)
        Me.CmdInsertSeparatorLine = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdInsertSeparatorLine_Callback, Guids.CmdGroupCommonFormatting, CommandIds.InsertSeparatorLine)
        Logger.Info("Code Formatting commands registered (Bold, Italic, Underline, Paragraph, InlineCode, MultilineCode, CodeExample, Remarks, GenerateSnippet, InsertSeparatorLine)")

        ' Section: editor operations
        Logger.Debug("Registering Editor Operations section commands...")
        Me.CmdCollapseXmlComments = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdCollapseXmlComments_Callback, Guids.CmdGroupEditorOperations, CommandIds.CollapseXmlComments)
        Me.CmdExpandXmlComments = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdExpandXmlComments_Callback, Guids.CmdGroupEditorOperations, CommandIds.ExpandXmlComments)
        Me.CmdDeleteXmlComments = Snippet_Tool_2026Package.CreateAndRegister(mcs, AddressOf Me.CmdDeleteXmlComments_Callback, Guids.CmdGroupEditorOperations, CommandIds.DeleteXmlComments)
        Logger.Info("Editor Operations commands registered (CollapseXmlComments, ExpandXmlComments, DeleteXmlComments)")

        Logger.Info($"InitializeMenuHandlersAsync — END — total commands registered: 18")

    End Function

    ''' <summary>
    ''' Creates a new <see cref="OleMenuCommand"/> and registers it with the given <see cref="OleMenuCommandService"/> in a single atomic operation.
    ''' </summary>
    ''' 
    ''' <param name="mcs">
    ''' The <see cref="OleMenuCommandService"/> with which the command will be registered.
    ''' </param>
    ''' 
    ''' <param name="callback">
    ''' The <see cref="EventHandler"/> invoked when the command is executed.
    ''' </param>
    ''' 
    ''' <param name="guid">
    ''' The <see cref="Guid"/> identifying the command group, as defined in the <c>.vsct</c> file.
    ''' </param>
    ''' 
    ''' <param name="id">
    ''' The numeric identifier of the command within the group, as defined in the <c>.vsct</c> file.
    ''' </param>
    ''' 
    ''' <returns>
    ''' The newly created and registered <see cref="OleMenuCommand"/> instance.
    ''' </returns>
    Private Shared Function CreateAndRegister(mcs As OleMenuCommandService, callback As EventHandler,
                                              guid As System.Guid, id As Integer) As OleMenuCommand

        Logger.Debug($"Registering command — GUID: {guid}, ID: {id}, Callback: {callback.Method.Name}")

        Dim cmd As New OleMenuCommand(callback, New CommandID(guid, id))

        Try
            mcs.AddCommand(cmd)
            Logger.Info($"Command registered — GUID: {guid}, ID: {id}, Callback: {callback.Method.Name}")
        Catch ex As Exception
            Logger.Error($"Failed to register command — GUID: {guid}, ID: {id}, Callback: {callback.Method.Name}", ex)
            Throw
        End Try

        Return cmd

    End Function

#End Region

#Region " IDisposable Implementation "

    ''' <summary>
    ''' Releases the resources used by this <see cref="Main"/> package instance.
    ''' </summary>
    ''' 
    ''' <param name="disposing">
    ''' <see langword="True"/> if the object is being disposed explicitly;
    ''' <see langword="False"/> if it is being finalized.
    ''' </param>
    Protected Overrides Sub Dispose(disposing As Boolean)

        Logger.Debug($"Dispose — START — disposing: {disposing}")

        If disposing Then
            Logger.Debug("Dispose — releasing managed resources")

            ' Section: references
            Me.CmdWrapAsCodeRef = Nothing
            Me.CmdWrapAsParamRef = Nothing
            Me.CmdWrapAsLangRef = Nothing
            Me.CmdWrapAsHrefLink = Nothing
            Me.CmdWrapAsSeeAlsoLink = Nothing
            Logger.Debug("Dispose — References commands released")

            ' Section: code formatting
            Me.CmdWrapAsBold = Nothing
            Me.CmdWrapAsItalic = Nothing
            Me.CmdWrapAsUnderline = Nothing
            Me.CmdWrapAsParagraph = Nothing
            Me.CmdWrapAsInlineCode = Nothing
            Me.CmdWrapAsMultilineCode = Nothing
            Me.CmdWrapAsCodeExample = Nothing
            Me.CmdWrapAsRemarks = Nothing
            Me.CmdGenerateSnippet = Nothing
            Me.CmdInsertSeparatorLine = Nothing
            Logger.Debug("Dispose — Code Formatting commands released")

            ' Section: editor operations
            Me.CmdCollapseXmlComments = Nothing
            Me.CmdExpandXmlComments = Nothing
            Me.CmdDeleteXmlComments = Nothing
            Logger.Debug("Dispose — Editor Operations commands released")

            ' DTE-related resources.
            Me.DteInitializer = Nothing
            Snippet_Tool_2026Package.Dte = Nothing
            Logger.Debug("Dispose — DTE resources released")

            Logger.Info("Dispose — all managed resources released")
        End If

        MyBase.Dispose(disposing)
        Logger.Debug("Dispose — END — MyBase.Dispose called")

    End Sub

#End Region

End Class
