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

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop

#End Region

#Region " DTE Initializer "

Namespace MyPackage.Helpers

    ''' <summary>
    ''' Fallback mechanism that defers DTE initialization until the Visual Studio IDE
    ''' is fully loaded, by listening for shell property change notifications.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' This class is only instantiated when <see cref="EnvDTE80.DTE2"/> is not yet available
    ''' at the time <c>InitializeDteAsync</c> is called. Once the IDE signals it is no longer
    ''' in a zombie state, the registered callback is invoked to retry DTE acquisition.
    ''' </remarks>
    Friend NotInheritable Class DteInitializer : Implements IVsShellPropertyEvents

#Region " Private Fields "

        ''' <summary>
        ''' The <see cref="IVsShell"/> instance used to subscribe and unsubscribe
        ''' from shell property change notifications.
        ''' </summary>
        Private ReadOnly shellService As IVsShell

        ''' <summary>
        ''' The callback delegate invoked once the IDE has finished initializing.
        ''' </summary>
        Private ReadOnly callback As Action

        ''' <summary>
        ''' The subscription cookie returned by <see cref="IVsShell.AdviseShellPropertyChanges"/>,
        ''' used to unsubscribe when the IDE is no longer in a zombie state.
        ''' </summary>
        Private cookie As UInteger

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DteInitializer"/> class
        ''' and subscribes to shell property change notifications.
        ''' </summary>
        ''' 
        ''' <param name="shellService">
        ''' The <see cref="IVsShell"/> instance used to monitor IDE initialization state.
        ''' </param>
        ''' 
        ''' <param name="callback">
        ''' The delegate invoked when the IDE has finished loading and is no longer in a zombie state.
        ''' </param>
        Friend Sub New(shellService As IVsShell, callback As Action)

            Logger.Debug($"DteInitializer.New — START — shellService IsNot Nothing: {shellService IsNot Nothing}")

            Me.shellService = shellService
            Me.callback = callback

            ' Sets an event handler to detect when the IDE is fully initialized.
            Dim hr As Integer = Me.shellService.AdviseShellPropertyChanges(Me, Me.cookie)
            Logger.Debug($"DteInitializer.New — AdviseShellPropertyChanges HRESULT: 0x{hr:X8}, cookie: {Me.cookie}")
            ErrorHandler.ThrowOnFailure(hr)

            Logger.Info("DteInitializer.New — subscribed to shell property changes successfully")
        End Sub

#End Region

#Region " IVsShellPropertyEvents Implementation "

        ''' <summary>
        ''' Invoked by the shell when a property value changes.
        ''' <para></para>
        ''' Detects when the IDE transitions out of its zombie state and triggers the callback.
        ''' </summary>
        ''' 
        ''' <param name="propid">
        ''' The identifier of the property that changed.
        ''' </param>
        ''' 
        ''' <param name="var">
        ''' The new value of the property.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see cref="VSConstants.S_OK"/> if the method succeeds; otherwise, an error code.
        ''' </returns>
        Private Function IVsShellPropertyEvents_OnShellPropertyChange(propid As Integer, var As Object) As Integer _
        Implements IVsShellPropertyEvents.OnShellPropertyChange

            Logger.Debug($"OnShellPropertyChange — propid: {propid}, value: '{var}'")

            If propid = __VSSPROPID.VSSPROPID_Zombie AndAlso Not CBool(var) Then
                Logger.Info("OnShellPropertyChange — IDE exited zombie state, unsubscribing and firing callback")

                ' Releases the event handler to detect when the IDE is fully initialized.
                Dim hr As Integer = Me.shellService.UnadviseShellPropertyChanges(Me.cookie)
                Logger.Debug($"OnShellPropertyChange — UnadviseShellPropertyChanges HRESULT: 0x{hr:X8}")

                Me.cookie = 0
                ErrorHandler.ThrowOnFailure(hr)

                Logger.Debug("OnShellPropertyChange — invoking DTE initialization callback")
                Me.callback()
                Logger.Info("OnShellPropertyChange — DTE initialization callback invoked successfully")
            End If

            Return VSConstants.S_OK

        End Function

#End Region

    End Class

End Namespace

#End Region