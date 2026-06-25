Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports Bricscad.ApplicationServices
Imports Application = Bricscad.ApplicationServices.Application

Namespace Arcxis_Cad_Tools
    ''' <summary>
    ''' Append-only activity log for correlating plugin actions with BricsCAD crashes.
    ''' Primary log: C:\BricsCADCrash\Logs\plugin-activity-{username}.log (same tree as ProcDump/Event1000).
    ''' Fallback:    %LOCALAPPDATA%\Arcxis\BricsCAD\plugin-activity-{username}.log
    ''' </summary>
    Public Module ArcxisActivityLog
        Private Const PrimaryLogDirectory As String = "C:\BricsCADCrash\Logs"
        Private ReadOnly FallbackLogDirectory As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Arcxis",
            "BricsCAD")
        Private ReadOnly SyncRoot As New Object()
        Private ReadOnly LogFilePath As String = ResolveLogFilePath()
        Private Const MaxLogBytes As Long = 2 * 1024 * 1024
        Private LastWriteUtc As DateTime = DateTime.MinValue
        Private ReadOnly ResumeIdleThreshold As TimeSpan = TimeSpan.FromMinutes(5)

        Public ReadOnly Property LogPath As String
            Get
                Return LogFilePath
            End Get
        End Property

        ''' <summary>
        ''' Call from a throttled Application.Idle handler to log return-from-idle
        ''' even when the user only interacts with BricsCAD (pan/zoom, etc.).
        ''' </summary>
        Public Sub CheckIdleResume()
            Try
                SyncLock SyncRoot
                    WriteResumeIfNeeded()
                End SyncLock
            Catch
            End Try
        End Sub

        Public Sub WriteLog(message As String)
            Try
                SyncLock SyncRoot
                    WriteResumeIfNeeded()
                    AppendLine(message)
                End SyncLock
            Catch
                ' Logging must never interrupt BricsCAD.
            End Try
        End Sub

        Public Sub WriteLog(ex As System.Exception, context As String)
            WriteLog(String.Format("{0}|EX|{1}|{2}", context, ex.GetType().Name, ex.Message))
        End Sub

        Public Sub LogPluginInit()
            Dim version = Assembly.GetExecutingAssembly().GetName().Version
            WriteLog(String.Format("PLUGIN INIT|v{0}", version))
        End Sub

        Public Sub LogPluginTerminate()
            WriteLog("PLUGIN TERMINATE")
        End Sub

        Public Sub LogCommandStart(commandName As String)
            WriteLog(String.Format("CMD START|{0}", commandName))
        End Sub

        Public Sub LogCommandEnd(commandName As String)
            WriteLog(String.Format("CMD END|{0}", commandName))
        End Sub

        Public Sub LogUi(action As String)
            WriteLog(String.Format("UI|{0}", action))
        End Sub

        Private Sub WriteResumeIfNeeded()
            If LastWriteUtc = DateTime.MinValue Then
                LastWriteUtc = DateTime.UtcNow
                Return
            End If

            Dim idle = DateTime.UtcNow - LastWriteUtc
            If idle < ResumeIdleThreshold Then Return

            Dim idleMinutes = CInt(Math.Floor(idle.TotalMinutes))
            AppendLine(String.Format("RESUME|idle {0}m", idleMinutes))
        End Sub

        Private Sub AppendLine(message As String)
            Dim logDirectory = Path.GetDirectoryName(LogFilePath)
            If Not String.IsNullOrEmpty(logDirectory) Then
                Directory.CreateDirectory(logDirectory)
            End If
            RotateIfNeeded()

            Dim docName = GetActiveDocumentName()
            Dim line = String.Format(
                "{0:yyyy-MM-dd HH:mm:ss.fff}|T{1}|{2}|{3}",
                DateTime.Now,
                Thread.CurrentThread.ManagedThreadId,
                docName,
                message)

            File.AppendAllText(LogFilePath, line & Environment.NewLine)
            LastWriteUtc = DateTime.UtcNow
        End Sub

        Private Sub RotateIfNeeded()
            If Not File.Exists(LogFilePath) Then Return
            If New FileInfo(LogFilePath).Length <= MaxLogBytes Then Return

            Dim rotatedPath = LogFilePath & "." & DateTime.Now.ToString("yyyyMMdd-HHmmss") & ".old"
            File.Move(LogFilePath, rotatedPath)
        End Sub

        Private Function GetActiveDocumentName() As String
            Try
                Dim doc = Application.DocumentManager.MdiActiveDocument
                If doc Is Nothing Then Return "-"
                If String.IsNullOrEmpty(doc.Name) Then Return "Drawing"
                Return Path.GetFileName(doc.Name)
            Catch
                Return "-"
            End Try
        End Function

        Private Function ResolveLogFilePath() As String
            Dim fileName = String.Format("plugin-activity-{0}.log", SanitizeFileName(Environment.UserName))

            If CanWriteToDirectory(PrimaryLogDirectory) Then
                Return Path.Combine(PrimaryLogDirectory, fileName)
            End If

            Return Path.Combine(FallbackLogDirectory, fileName)
        End Function

        Private Function SanitizeFileName(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return "unknown"

            Dim invalid = Path.GetInvalidFileNameChars()
            Dim sanitized = value
            For Each ch In invalid
                sanitized = sanitized.Replace(ch, "_"c)
            Next
            Return sanitized
        End Function

        Private Function CanWriteToDirectory(directoryPath As String) As Boolean
            Try
                Directory.CreateDirectory(directoryPath)
                Dim probePath = Path.Combine(directoryPath, ".arcxis-log-probe")
                File.WriteAllText(probePath, "ok")
                File.Delete(probePath)
                Return True
            Catch
                Return False
            End Try
        End Function
    End Module
End Namespace
