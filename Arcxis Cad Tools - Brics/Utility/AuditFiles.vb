Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.AuditFiles))>
Namespace Arcxis_Cad_Tools

    Public Class AuditFiles

        ''' <summary>
        ''' Command to audit all DWG files in a selected folder
        ''' </summary>
        <CommandMethod("AuditFilesInFolder")>
        Public Sub AuditFilesInFolder()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acEd As Editor = acDoc.Editor

            Try
                ' Prompt user to select folder
                Dim folderPath As String = PromptForFolder()
                If String.IsNullOrEmpty(folderPath) Then
                    acEd.WriteMessage(vbLf & "Operation cancelled by user.")
                    Return
                End If

                ' Get all DWG files in the selected folder
                Dim dwgFiles As String() = Directory.GetFiles(folderPath, "*.dwg", SearchOption.TopDirectoryOnly)

                If dwgFiles.Length = 0 Then
                    acEd.WriteMessage(vbLf & "No DWG files found in the selected folder.")
                    Return
                End If

                acEd.WriteMessage(vbLf & $"Found {dwgFiles.Length} DWG file(s) to audit.")
                acEd.WriteMessage(vbLf & "Starting audit process...")

                Dim successCount As Integer = 0
                Dim errorCount As Integer = 0
                Dim errorFiles As New List(Of String)()

                ' Process each DWG file
                For Each dwgFile As String In dwgFiles
                    Try
                        acEd.WriteMessage(vbLf & $"Processing: {Path.GetFileName(dwgFile)}")

                        ' Audit the file
                        If AuditSingleFile(dwgFile, acEd) Then
                            successCount += 1
                            acEd.WriteMessage(" - Success")
                        Else
                            errorCount += 1
                            errorFiles.Add(Path.GetFileName(dwgFile))
                            acEd.WriteMessage(" - Failed")
                        End If

                    Catch ex As System.Exception
                        errorCount += 1
                        errorFiles.Add(Path.GetFileName(dwgFile))
                        acEd.WriteMessage(vbLf & $"Error processing {Path.GetFileName(dwgFile)}: {ex.Message}")
                    End Try
                Next

                ' Display summary
                acEd.WriteMessage(vbLf & "=== Audit Complete ===")
                acEd.WriteMessage(vbLf & $"Successfully audited: {successCount} file(s)")
                acEd.WriteMessage(vbLf & $"Failed: {errorCount} file(s)")

                If errorFiles.Count > 0 Then
                    acEd.WriteMessage(vbLf & "Failed files:")
                    For Each fileName As String In errorFiles
                        acEd.WriteMessage(vbLf & $"  - {fileName}")
                    Next
                End If

            Catch ex As System.Exception
                acEd.WriteMessage(vbLf & $"Fatal error: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Audits a single DWG file
        ''' </summary>
        Private Function AuditSingleFile(filePath As String, acEd As Editor) As Boolean
            Dim tempDb As Database = Nothing

            Try
                ' Create a new database and read the DWG file
                tempDb = New Database(False, True)
                tempDb.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, False, "")

                ' Perform audit with AuditInfo
                Dim auditInfo As New AuditInfo()
                auditInfo.FixErrors = True
                tempDb.Audit(auditInfo)

                ' Save the audited file
                tempDb.SaveAs(filePath, DwgVersion.Current)

                Return True

            Catch ex As System.Exception
                acEd.WriteMessage(vbLf & $"  Error: {ex.Message}")
                Return False

            Finally
                ' Dispose of the database
                If tempDb IsNot Nothing Then
                    tempDb.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Prompts the user to select a folder
        ''' </summary>
        Private Function PromptForFolder() As String
            Dim startPath As String = GetCurrentDwgFolder()
            Dim owner As New WindowWrapper(Application.MainWindow.Handle)
            Dim chosen As String = Nothing

            Using dlg As New FolderBrowserDialog()
                dlg.Description = "Select folder containing DWG files to audit"
                If Directory.Exists(startPath) Then
                    dlg.SelectedPath = startPath
                End If
                dlg.ShowNewFolderButton = False

                Dim result = dlg.ShowDialog(owner)
                If result = DialogResult.OK AndAlso Directory.Exists(dlg.SelectedPath) Then
                    chosen = dlg.SelectedPath

                    ' Set focus back to BricsCAD
                    Application.MainWindow.Focus()
                    Try
                        Environment.CurrentDirectory = startPath
                    Catch
                        ' Ignore if fails
                    End Try

                    Return chosen
                End If
            End Using

            Return Nothing
        End Function

        ''' <summary>
        ''' Gets the folder of the current drawing
        ''' </summary>
        Private Function GetCurrentDwgFolder() As String
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            If Not String.IsNullOrWhiteSpace(db.Filename) Then
                Return Path.GetDirectoryName(db.Filename)
            End If
            ' Unsaved drawing fallback
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

        ''' <summary>
        ''' Command to audit the currently open drawing
        ''' </summary>
        <CommandMethod("AuditCurrentFile")>
        Public Sub AuditCurrentFile()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor

            Try
                Using acLckDoc As DocumentLock = acDoc.LockDocument()
                    acEd.WriteMessage(vbLf & "Auditing current drawing...")

                    ' Perform audit on the current database with AuditInfo
                    Dim auditInfo As New AuditInfo()
                    auditInfo.FixErrors = True
                    acDb.Audit(auditInfo)

                    acEd.WriteMessage(vbLf & "Audit complete.")
                    acEd.WriteMessage(vbLf & "Note: Save the drawing to preserve changes.")
                End Using

            Catch ex As System.Exception
                acEd.WriteMessage(vbLf & $"Error during audit: {ex.Message}")
            End Try
        End Sub

    End Class

    ''' <summary>
    ''' Helper class for setting the parent window of dialogs
    ''' </summary>
    Public Class WindowWrapper
        Implements IWin32Window

        Private _handle As IntPtr

        Public Sub New(handle As IntPtr)
            _handle = handle
        End Sub

        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _handle
            End Get
        End Property
    End Class

End Namespace
