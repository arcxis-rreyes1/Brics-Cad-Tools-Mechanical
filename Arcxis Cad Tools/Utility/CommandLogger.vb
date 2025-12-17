Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors

' Put this OUTSIDE any class or namespace
<Assembly: ExtensionApplication(GetType(LoggingTrial.CommandLogger))>
Namespace LoggingTrial
    Public Class CommandLogger
        Implements IExtensionApplication

        Private Shared logDir As String = Path.GetTempPath()
        Private Shared loggingEnabled As Boolean = True

        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Dim ed = Application.DocumentManager.MdiActiveDocument.Editor
            'ed.WriteMessage(vbLf & "[Logger initialized]")

            ' Ensure log directory exists at startup
            Try
                If Not Directory.Exists(logDir) Then
                    Directory.CreateDirectory(logDir)
                End If
            Catch ex As Exception
                'ed.WriteMessage(vbLf & $"[Logging Directory Error] {ex.Message}")
            End Try

            ' Attach to all open documents
            For Each doc In Application.DocumentManager
                AttachCommandEvents(doc)
            Next

            ' Watch for new docs opening or closing
            AddHandler Application.DocumentManager.DocumentCreated, AddressOf OnDocumentCreated
            AddHandler Application.DocumentManager.DocumentToBeDestroyed, AddressOf OnDocumentClosed
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
        End Sub

        Private Sub OnDocumentCreated(sender As Object, e As DocumentCollectionEventArgs)
            ' Delete old log file if it exists in temp folder
            Dim logFilePath = GetLogFilePath(e.Document)
            If File.Exists(logFilePath) Then
                Try
                    File.Delete(logFilePath)
                Catch ex As Exception
                    ' Optionally log or ignore
                End Try
            End If
            AttachCommandEvents(e.Document)
            If loggingEnabled Then LogMessage($"[OPEN] {e.Document.Name} at {DateTime.Now}", e.Document)
        End Sub

        Private Sub OnDocumentClosed(sender As Object, e As DocumentCollectionEventArgs)
            If loggingEnabled Then LogMessage($"[CLOSE] {e.Document.Name} at {DateTime.Now}", e.Document)

            ' Copy log file to network path and delete local log file
            Try
                Dim logFilePath = GetLogFilePath(e.Document)
                If File.Exists(logFilePath) Then
                    Dim networkDir = "\\egnytedrive\energyinspectors\Shared\FS2\K\DPIS Drawings\ToPrint\Failed\old\LispHistory"
                    If Not Directory.Exists(networkDir) Then
                        Directory.CreateDirectory(networkDir)
                    End If
                    Dim fileName = Path.GetFileName(logFilePath)
                    Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                    Dim ext = Path.GetExtension(fileName)
                    Dim destPath = Path.Combine(networkDir, fileName)
                    Dim counter = 1
                    While File.Exists(destPath)
                        destPath = Path.Combine(networkDir, $"{baseName}_{counter}{ext}")
                        counter += 1
                    End While
                    File.Copy(logFilePath, destPath, True)
                    File.Delete(logFilePath)
                End If
            Catch ex As Exception
                Dim ed = Application.DocumentManager.MdiActiveDocument.Editor
                ed.WriteMessage(vbLf & $"[Log Move Error] {ex.Message}")
            End Try
        End Sub

        Private Sub AttachCommandEvents(doc As Document)
            AddHandler doc.CommandWillStart, AddressOf OnCommandStart
            AddHandler doc.CommandEnded, AddressOf OnCommandEnd
            AddHandler doc.CommandCancelled, AddressOf OnCommandCancelled
        End Sub


        Private Sub OnCommandStart(sender As Object, e As CommandEventArgs)
            Dim doc = TryCast(sender, Document)
            Dim isLisp = (e.GlobalCommandName.Equals("LISP", StringComparison.OrdinalIgnoreCase) OrElse e.GlobalCommandName.Equals("LISPEXEC", StringComparison.OrdinalIgnoreCase))
            Dim tag = If(isLisp, " [LISP]", "")
            If loggingEnabled Then LogMessage($"  Command Start: {e.GlobalCommandName}{tag} at {DateTime.Now}", doc)
        End Sub

        Private Sub OnCommandEnd(sender As Object, e As CommandEventArgs)
            Dim doc = TryCast(sender, Document)
            Dim isLisp = (e.GlobalCommandName.Equals("LISP", StringComparison.OrdinalIgnoreCase) OrElse e.GlobalCommandName.Equals("LISPEXEC", StringComparison.OrdinalIgnoreCase))
            Dim tag = If(isLisp, " [LISP]", "")
            If loggingEnabled Then LogMessage($"  Command End: {e.GlobalCommandName}{tag} at {DateTime.Now}", doc)
        End Sub

        Private Sub OnCommandCancelled(sender As Object, e As CommandEventArgs)
            Dim doc = TryCast(sender, Document)
            Dim isLisp = (e.GlobalCommandName.Equals("LISP", StringComparison.OrdinalIgnoreCase) OrElse e.GlobalCommandName.Equals("LISPEXEC", StringComparison.OrdinalIgnoreCase))
            Dim tag = If(isLisp, " [LISP]", "")
            If loggingEnabled Then LogMessage($"  Command Cancelled: {e.GlobalCommandName}{tag} at {DateTime.Now}", doc)
        End Sub

        ' LogMessage now includes the file name in the log entry and uses per-document log file
        Private Sub LogMessage(message As String, Optional doc As Document = Nothing)
            Try
                Dim fileName As String = "UnknownFile"
                If doc Is Nothing Then
                    doc = Application.DocumentManager.MdiActiveDocument
                End If
                If doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.Name) Then
                    fileName = Path.GetFileName(doc.Name)
                End If
                Dim logFilePath = GetLogFilePath(doc)
                Dim dir = Path.GetDirectoryName(logFilePath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                Using sw As StreamWriter = File.AppendText(logFilePath)
                    sw.WriteLine($"[{fileName}] {message}")
                End Using
            Catch ex As Exception
                Dim ed = Application.DocumentManager.MdiActiveDocument.Editor
                ed.WriteMessage(vbLf & $"[Logging Error] {ex.Message}")
            End Try
        End Sub

        ' Helper to get per-document log file path
        Private Function GetLogFilePath(doc As Document) As String
            Dim user = Environment.UserName
            Dim filePart As String = "UnknownFile"
            If doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.Name) Then
                filePart = Path.GetFileNameWithoutExtension(doc.Name)
            End If
            Return Path.Combine(logDir, $"AcadCommandLog_{user}_{filePart}.csv")
        End Function

        <CommandMethod("TOGGLELOGGING")>
        Public Sub ToggleLogging()
            loggingEnabled = Not loggingEnabled
            Dim msg = If(loggingEnabled, "[Logging enabled]", "[Logging disabled]")
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & msg)
        End Sub

        <CommandMethod("TESTLOG")>
        Public Sub TestLog()
            LogMessage("This is a manual test log at " & DateTime.Now.ToString())
        End Sub
    End Class
End Namespace