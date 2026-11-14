Option Strict On
Option Explicit On
Option Infer Off

Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop
Imports Xunit
Imports Moq
Imports Snippet_Tool_2026.MyPackage.Helpers

''' <summary>Runs the real DTE shell listener against a mocked Visual Studio COM boundary.</summary>
Public Class DteInitializerTests

    ''' <summary>Creates a shell that assigns a known subscription cookie.</summary>
    Private Shared Function CreateShell() As Mock(Of IVsShell)
        Dim shell As New Mock(Of IVsShell)(MockBehavior.Strict)
        Dim cookie As UInteger = 42UI
        shell.Setup(Function(value As IVsShell) value.AdviseShellPropertyChanges(It.IsAny(Of IVsShellPropertyEvents)(), cookie)).Returns(VSConstants.S_OK)
        shell.Setup(Function(value As IVsShell) value.UnadviseShellPropertyChanges(42UI)).Returns(VSConstants.S_OK)
        Return shell
    End Function

    ''' <summary>Ignores unrelated shell events and the zombie state.</summary>
    <Theory>
    <InlineData(0, False)>
    <InlineData(1, False)>
    <InlineData(-1, True)>
    Public Sub IrrelevantEventsDoNotInvokeCallback(propertyId As Integer, value As Boolean)
        Dim shell As Mock(Of IVsShell) = CreateShell()
        Dim calls As Integer = 0
        Dim listener As New DteInitializer(shell.Object, Sub() calls += 1)
        Dim events As IVsShellPropertyEvents = listener
        Assert.Equal(VSConstants.S_OK, events.OnShellPropertyChange(propertyId, value))
        Assert.Equal(0, calls)
        shell.Verify(Function(instance As IVsShell) instance.UnadviseShellPropertyChanges(It.IsAny(Of UInteger)()), Times.Never())
    End Sub

    ''' <summary>Does not initialize while the IDE remains in the zombie state.</summary>
    <Fact>
    Public Sub ZombieStateDoesNotInitialize()
        Dim shell As Mock(Of IVsShell) = CreateShell()
        Dim calls As Integer = 0
        Dim listener As New DteInitializer(shell.Object, Sub() calls += 1)
        DirectCast(listener, IVsShellPropertyEvents).OnShellPropertyChange(CInt(__VSSPROPID.VSSPROPID_Zombie), True)
        Assert.Equal(0, calls)
    End Sub

    ''' <summary>Unsubscribes with the issued cookie before invoking the deferred initialization callback.</summary>
    <Fact>
    Public Sub ReadyStateUnsubscribesBeforeCallback()
        Dim shell As Mock(Of IVsShell) = CreateShell()
        Dim calls As Integer = 0
        Dim listener As New DteInitializer(shell.Object,
            Sub()
                shell.Verify(Function(instance As IVsShell) instance.UnadviseShellPropertyChanges(42UI), Times.Once())
                calls += 1
            End Sub)
        Assert.Equal(VSConstants.S_OK, DirectCast(listener, IVsShellPropertyEvents).OnShellPropertyChange(CInt(__VSSPROPID.VSSPROPID_Zombie), False))
        Assert.Equal(1, calls)
    End Sub

    ''' <summary>Propagates subscription failure instead of pretending initialization is deferred.</summary>
    <Fact>
    Public Sub SubscriptionFailureIsReported()
        Dim shell As New Mock(Of IVsShell)(MockBehavior.Strict)
        Dim cookie As UInteger = 0UI
        shell.Setup(Function(value As IVsShell) value.AdviseShellPropertyChanges(It.IsAny(Of IVsShellPropertyEvents)(), cookie)).Returns(VSConstants.E_FAIL)
        Assert.Throws(Of COMException)(Sub() NewListener(shell.Object))
    End Sub

    ''' <summary>Constructs a listener with an empty callback for failure-path tests.</summary>
    Private Shared Sub NewListener(shell As IVsShell)
        Dim listener As New DteInitializer(shell, Sub()
                                                End Sub)
    End Sub

    ''' <summary>Reports unsubscription failure without invoking the callback.</summary>
    <Fact>
    Public Sub UnsubscriptionFailureDoesNotInvokeCallback()
        Dim shell As Mock(Of IVsShell) = CreateShell()
        shell.Setup(Function(value As IVsShell) value.UnadviseShellPropertyChanges(42UI)).Returns(VSConstants.E_FAIL)
        Dim calls As Integer = 0
        Dim listener As New DteInitializer(shell.Object, Sub() calls += 1)
        Assert.Throws(Of COMException)(Sub() DirectCast(listener, IVsShellPropertyEvents).OnShellPropertyChange(CInt(__VSSPROPID.VSSPROPID_Zombie), False))
        Assert.Equal(0, calls)
    End Sub
End Class
