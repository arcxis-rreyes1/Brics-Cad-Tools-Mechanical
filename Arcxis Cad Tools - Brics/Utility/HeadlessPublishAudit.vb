Imports System.IO
Imports System.Text
Imports System.Threading
Imports Bricscad.ApplicationServices

Public NotInheritable Class HeadlessPublishAudit
    Private Shared ReadOnly _syncRoot As New Object()

    Private Sub New()
    End Sub

    Public Shared Function StartJob(moduleName As String, outputFile As String, sheetCount As Integer, Optional details As String = "") As String
        Dim jobId As String = Guid.NewGuid().ToString("N")
        WriteLine(jobId, "STARTED", moduleName, outputFile, sheetCount, details)
        Return jobId
    End Function

    Public Shared Sub MarkSkipped(jobId As String, moduleName As String, outputFile As String, Optional details As String = "")
        WriteLine(jobId, "SKIPPED", moduleName, outputFile, Nothing, details)
    End Sub

    Public Shared Sub MarkInfo(jobId As String, moduleName As String, outputFile As String, Optional details As String = "")
        WriteLine(jobId, "INFO", moduleName, outputFile, Nothing, details)
    End Sub

    Public Shared Sub MarkSuccess(jobId As String, moduleName As String, outputFile As String, Optional details As String = "")
        WriteLine(jobId, "SUCCESS", moduleName, outputFile, Nothing, details)
    End Sub

    Public Shared Sub MarkFailed(jobId As String, moduleName As String, outputFile As String, Optional details As String = "")
        WriteLine(jobId, "FAILED", moduleName, outputFile, Nothing, details)
    End Sub

    Public Shared Function WaitForOutput(outputFile As String, timeoutMs As Integer) As Boolean
        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
        Do While sw.ElapsedMilliseconds < timeoutMs
            Try
                If File.Exists(outputFile) Then
                    Using fs As New FileStream(outputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    End Using
                    Return True
                End If
            Catch
                ' Output exists but is still being written; keep polling.
            End Try
            Thread.Sleep(200)
        Loop

        Return False
    End Function

    Private Shared Sub WriteLine(jobId As String, status As String, moduleName As String, outputFile As String, sheetCount As Integer?, details As String)
        Try
            Dim logPath As String = GetLogPath(outputFile)
            Dim line As String = BuildLine(jobId, status, moduleName, outputFile, sheetCount, details)

            SyncLock _syncRoot
                Dim dir As String = Path.GetDirectoryName(logPath)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                File.AppendAllText(logPath, line & Environment.NewLine, Encoding.UTF8)
            End SyncLock
        Catch
            ' Logging should never break the print flow.
        End Try
    End Sub

    Private Shared Function BuildLine(jobId As String, status As String, moduleName As String, outputFile As String, sheetCount As Integer?, details As String) As String
        Dim nowStamp As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
        Dim dwgName As String = GetSafeSystemVariable("DWGNAME")
        Dim pid As String = Diagnostics.Process.GetCurrentProcess().Id.ToString()
        Dim sheetText As String = If(sheetCount.HasValue, sheetCount.Value.ToString(), "")

        Return String.Join("|", New String() {
            nowStamp,
            Sanitize(jobId),
            Sanitize(status),
            Sanitize(moduleName),
            Sanitize(dwgName),
            Sanitize(pid),
            Sanitize(outputFile),
            Sanitize(sheetText),
            Sanitize(details)
        })
    End Function

    Private Shared Function GetLogPath(outputFile As String) As String
        Try
            If Not String.IsNullOrWhiteSpace(outputFile) Then
                Dim outDir As String = Path.GetDirectoryName(outputFile)
                If Not String.IsNullOrWhiteSpace(outDir) Then
                    Return Path.Combine(outDir, "HeadlessPublishAudit.log")
                End If
            End If
        Catch
        End Try

        Dim fallbackDir As String = GetSafeSystemVariable("DWGPREFIX")
        If String.IsNullOrWhiteSpace(fallbackDir) Then
            fallbackDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If

        Return Path.Combine(fallbackDir, "HeadlessPublishAudit.log")
    End Function

    Private Shared Function GetSafeSystemVariable(name As String) As String
        Try
            Return CStr(Application.GetSystemVariable(name))
        Catch
            Return String.Empty
        End Try
    End Function

    Private Shared Function Sanitize(value As String) As String
        If value Is Nothing Then Return String.Empty
        Return value.Replace("|", "/").Replace(vbCr, " ").Replace(vbLf, " ")
    End Function
End Class
