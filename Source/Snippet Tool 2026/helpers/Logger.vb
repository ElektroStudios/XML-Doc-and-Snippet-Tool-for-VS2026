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
#If DEBUG Then
Imports System.Diagnostics
Imports System.IO
#End If
Imports System.Runtime.CompilerServices

#End Region

#Region " Logger "

''' <summary>
''' Provides static logging functionality for the Snippet Tool extension.
''' Writes timestamped, leveled log entries to a session-scoped text file
''' under <c>%LocalAppData%\Snippet Tool\Logs\</c>.
''' </summary>
Public NotInheritable Class Logger

#If DEBUG Then
    Private Shared ReadOnly EnableLogger As Boolean = False ' Set to True to enable logging in Debug builds

    Private Shared ReadOnly _lock As New Object()
#End If

    ''' <summary>
    ''' Gets the full path of the current session log file.
    ''' </summary>
    ''' 
    ''' <returns>
    ''' The absolute path to the log file created by <see cref="Initialize"/>,
    ''' or <see langword="Nothing"/> if <see cref="Initialize"/> has not been called yet.
    ''' </returns>
    Public Shared ReadOnly Property LogFilePath As String

    ''' <summary>
    ''' Initializes the logger and creates the session log file.
    ''' Must be called once before any other logging method.
    ''' </summary>
    ''' 
    ''' <param name="logFileName">
    ''' Base name of the log file. A timestamp prefix is prepended automatically
    ''' to ensure each session produces a unique file.
    ''' Defaults to <c>"Snippet Tool.txt"</c>.
    ''' </param>
    ''' 
    ''' <remarks>
    ''' The log file is created under <c>%LocalAppData%\Snippet Tool\Logs\</c>.
    ''' The directory is created if it does not already exist.
    ''' </remarks>
    Public Shared Sub Initialize(Optional logFileName As String = "Snippet Tool.txt")

#If DEBUG Then
        If Not EnableLogger Then 
            Return
        End If

        Dim folder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Snippet Tool", "Logs")

        Directory.CreateDirectory(folder)

        Dim timestamp As String = Date.Now.ToString("yyyyMMdd_HHmmss")
        _logFilePath = Path.Combine(folder, $"{timestamp}_{logFileName}")

        WriteEntry("INFO", "=== Logger initialized ===")
        WriteEntry("INFO", $"CLR Version: {Environment.Version}")
        WriteEntry("INFO", $"VS PID: {Process.GetCurrentProcess().Id}")
#End If
    End Sub

    ''' <summary>
    ''' Writes an <c>INFO</c>-level entry to the log file.
    ''' </summary>
    ''' 
    ''' <param name="message">
    ''' The message to log.
    ''' </param>
    ''' 
    ''' <param name="memberName">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="lineNumber">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    Public Shared Sub Info(message As String,
                           <CallerMemberName> Optional memberName As String = "",
                           <CallerFilePath> Optional filePath As String = "",
                           <CallerLineNumber> Optional lineNumber As Integer = 0)

        WriteEntry("INFO", message, memberName, filePath, lineNumber)
    End Sub

    ''' <summary>
    ''' Writes a <c>WARN</c>-level entry to the log file.
    ''' Use for recoverable conditions that may indicate unexpected state.
    ''' </summary>
    ''' 
    ''' <param name="message">
    ''' The warning message to log.
    ''' </param>
    ''' 
    ''' <param name="memberName">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="lineNumber">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    Public Shared Sub Warn(message As String,
                           <CallerMemberName> Optional memberName As String = "",
                           <CallerFilePath> Optional filePath As String = "",
                           <CallerLineNumber> Optional lineNumber As Integer = 0)

        WriteEntry("WARN", message, memberName, filePath, lineNumber)
    End Sub

    ''' <summary>
    ''' Writes an <c>ERROR</c>-level entry to the log file.
    ''' Optionally includes full exception details and stack trace.
    ''' </summary>
    ''' 
    ''' <param name="message">
    ''' The error message to log.
    ''' </param>
    ''' 
    ''' <param name="ex">
    ''' Optional exception to append to the log entry.
    ''' When provided, the exception type, message, and stack trace are included.
    ''' </param>
    ''' 
    ''' <param name="memberName">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="lineNumber">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    Public Shared Sub [Error](message As String,
                              Optional ex As Exception = Nothing,
                              <CallerMemberName> Optional memberName As String = "",
                              <CallerFilePath> Optional filePath As String = "",
                              <CallerLineNumber> Optional lineNumber As Integer = 0)

        Dim fullMsg As String = If(ex IsNot Nothing, $"{message} | Exception: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}StackTrace: {ex.StackTrace}", message)
        WriteEntry("ERROR", fullMsg, memberName, filePath, lineNumber)
    End Sub

    ''' <summary>
    ''' Writes a <c>DEBUG</c>-level entry to the log file.
    ''' This method is compiled out entirely in Release builds.
    ''' </summary>
    ''' 
    ''' <param name="message">
    ''' The debug message to log.
    ''' </param>
    ''' 
    ''' <param name="memberName">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    ''' 
    ''' <param name="lineNumber">
    ''' Automatically supplied by the compiler. Do not pass explicitly.
    ''' </param>
    Public Shared Sub Debug(message As String,
                            <CallerMemberName> Optional memberName As String = "",
                            <CallerFilePath> Optional filePath As String = "",
                            <CallerLineNumber> Optional lineNumber As Integer = 0)

        WriteEntry("DEBUG", message, memberName, filePath, lineNumber)
    End Sub

    ''' <summary>
    ''' Core internal method that formats and appends a log entry to the session file.
    ''' All public logging methods delegate to this method.
    ''' </summary>
    ''' 
    ''' <param name="level">
    ''' The log level label (e.g. <c>"INFO"</c>, <c>"WARN"</c>, <c>"ERROR"</c>, <c>"DEBUG"</c>).
    ''' </param>
    ''' 
    ''' <param name="message">
    ''' The message body to write.
    ''' </param>
    ''' 
    ''' <param name="memberName">
    ''' Name of the calling member, injected by <see cref="CallerMemberNameAttribute"/>.
    ''' </param>
    ''' 
    ''' <param name="filePath">
    ''' Source file path of the caller, injected by <see cref="CallerFilePathAttribute"/>.
    ''' </param>
    ''' 
    ''' <param name="lineNumber">
    ''' Line number of the caller, injected by <see cref="CallerLineNumberAttribute"/>.
    ''' </param>
    ''' 
    ''' <remarks>
    ''' Write failures are silently swallowed to prevent the logger from disrupting extension operation.
    ''' Access to the file is synchronized via <c>_lock</c> to ensure thread safety.
    ''' </remarks>
    Private Shared Sub WriteEntry(level As String,
                                  message As String,
                                  Optional memberName As String = "",
                                  Optional filePath As String = "",
                                  Optional lineNumber As Integer = 0)

#If DEBUG Then
        If Not EnableLogger OrElse String.IsNullOrEmpty(_logFilePath) Then 
            Return
        End If

        Dim fileName As String = If(String.IsNullOrEmpty(filePath), "", Path.GetFileName(filePath))
        Dim location As String = If(String.IsNullOrEmpty(memberName), "", $" [{fileName}::{memberName} L{lineNumber}]")

        Dim entry As String = $"[{Date.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level,-5}]{location} {message}"

        SyncLock _lock
            Try
                File.AppendAllText(_logFilePath, entry & Environment.NewLine)
            Catch
            End Try
        End SyncLock
#End If
    End Sub

End Class

#End Region
