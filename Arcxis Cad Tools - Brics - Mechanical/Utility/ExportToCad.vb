Imports System.IO
Imports System.Runtime.InteropServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.ExportToCad))>
Namespace Arcxis_Cad_Tools

    Public Class ExportToCad

        ' Canonical drive that every machine uses to reach the shared network storage.
        ' xref paths that resolve to the same share through another drive letter or a raw
        ' UNC path are rewritten to this drive so all drawings store an identical location.
        Private Const SharedDriveLetter As String = "E:"

        ' Canonical root of the shared storage. Every shared file lives under here, and the
        ' folders directly beneath it are used to recognize stored paths that belong to the
        ' shared storage but came in under a different drive letter or UNC root.
        Private Const SharedRoot As String = "E:\Shared"

        ' Known roots that map onto the shared storage. A path that starts with a key is
        ' rewritten with the matching value (case-insensitive) and then verified on disk,
        ' so e.g. the Egnyte UNC or a legacy drive resolves to its real location under
        ' E:\Shared. Keep keys ending with "\" so only whole path segments match. Add new
        ' entries here as other legacy sources get migrated.
        Private Shared ReadOnly ForeignRootRemaps As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"\\egnytedrive\egnyte shared\", "E:\Shared\"},
            {"W:\", "E:\Shared\Migrated from Other Sources\Ei Server\"}
        }

        ' Cached top-level folder names directly under E:\Shared. Populated once per session
        ' from disk and used to rebase stored paths that reference a recognizable shared
        ' folder regardless of the drive letter/UNC prefix they were saved with.
        Private Shared _sharedTopFolders As HashSet(Of String)

        ' Cached UNC target of the shared drive (e.g. "\\server\share"). Resolved once per
        ' session so the WNetGetConnection lookup isn't repeated for every xref/underlay.
        ' Nothing is a valid cached answer (drive isn't a mapped network drive here).
        Private Shared _sharedUnc As String
        Private Shared _sharedUncResolved As Boolean

        Private Shared Function GetSharedUnc() As String
            If Not _sharedUncResolved Then
                _sharedUnc = GetDriveUncTarget(SharedDriveLetter)
                _sharedUncResolved = True
            End If
            Return _sharedUnc
        End Function

        <CommandMethod("B2C", CommandFlags.Modal)>
        Public Shared Sub BricsToAcad()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return

            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database

            If String.IsNullOrWhiteSpace(db.Filename) Then
                ed.WriteMessage(vbLf & "Drawing must be saved before creating a copy.")
                Return
            End If

            Try

                'Turn on and thaw all layers
                Using tr As Transaction = db.TransactionManager.StartTransaction()

                    Dim lt As LayerTable =
                CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

                    For Each layerId As ObjectId In lt

                        Dim ltr As LayerTableRecord =
                    CType(tr.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)

                        If ltr.IsOff Then ltr.IsOff = False
                        If ltr.IsFrozen Then ltr.IsFrozen = False
                        If ltr.IsLocked Then ltr.IsLocked = False

                    Next

                    tr.Commit()

                End Using

                Dim targetFolder As String =
            ArcxisPaths.NetworkUNCPathForEgnyte &
            "\FS2\K\DPIS Drawings\ToPrint\ConvertToCAD"

                If Not IO.Directory.Exists(targetFolder) Then
                    IO.Directory.CreateDirectory(targetFolder)
                End If

                Dim targetFile As String =
            IO.Path.Combine(
                targetFolder,
                IO.Path.GetFileName(db.Filename)
            )

                db.SaveAs(
            targetFile,
            True,
            DwgVersion.Current,
            db.SecurityParameters
        )

                ed.WriteMessage(vbLf & "Copy created:")
                ed.WriteMessage(vbLf & targetFile)

            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error creating copy: " & ex.Message)
            End Try

        End Sub


        <CommandMethod("xsr", CommandFlags.Modal)>
        Public Shared Sub MakeAllXrefPathsAbsolute()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            If String.IsNullOrWhiteSpace(db.Filename) Then
                ed.WriteMessage(vbLf & "Drawing must be saved first.")
                Return
            End If

            Dim xrefsToDetach As New List(Of ObjectId)
            ' Xrefs the user had turned off (unloaded) before we ran. db.ResolveXrefs below
            ' reloads everything, so we re-unload these afterward to preserve their state.
            Dim xrefsWereUnloaded As New ObjectIdCollection()

            Dim xrefRemapped As Integer = 0
            Dim xrefIssues As Integer = 0
            Dim pdfRemapped As Integer = 0
            Dim pdfIssues As Integer = 0

            ' Hold the document lock across the whole operation so every change (path
            ' rewrites, detach, resolve, unload restore, PDF pass) is collapsed into a
            ' single undo step rather than several.
            Using doc.LockDocument()

                Using tr As Transaction = db.TransactionManager.StartTransaction()

                    Dim xrefGraph As XrefGraph = db.GetHostDwgXrefGraph(True)

                    For i As Integer = 1 To xrefGraph.NumNodes - 1

                        ' Isolate each xref: a failure on one node must not abort the whole
                        ' transaction and throw away every other path fix in this drawing.
                        Try

                            Dim node As XrefGraphNode = xrefGraph.GetXrefNode(i)
                            If node Is Nothing OrElse node.BlockTableRecordId.IsNull Then Continue For

                            Dim btr As BlockTableRecord =
                            TryCast(tr.GetObject(node.BlockTableRecordId, OpenMode.ForWrite, False), BlockTableRecord)

                            If btr Is Nothing OrElse Not btr.IsFromExternalReference Then Continue For

                            ' Current engine state of this xref. Only genuinely missing/broken
                            ' references (file not found or unresolved) may be detached. A resolved
                            ' xref can be valid even when btr.PathName doesn't point straight at the
                            ' file (found via a support/project path), and an Unloaded xref was
                            ' intentionally turned off by the user - neither must ever be detached
                            ' just because we couldn't recompute a path. Load state is left as-is.
                            Dim status As XrefStatus = btr.XrefStatus
                            Dim isMissing As Boolean =
                                (status = XrefStatus.FileNotFound OrElse status = XrefStatus.Unresolved)

                            ' Remember if it was turned off so we can turn it back off after resolve.
                            If status = XrefStatus.Unloaded Then
                                xrefsWereUnloaded.Add(btr.ObjectId)
                            End If

                            Dim currentPath As String = btr.PathName

                            Dim ownerFolder As String = GetXrefOwnerFolder(node, db)
                            If String.IsNullOrWhiteSpace(ownerFolder) Then
                                ownerFolder = Path.GetDirectoryName(db.Filename)
                            End If

                            ' Locate the real file and normalize it onto E:\Shared. Nothing means
                            ' it could not be found anywhere.
                            Dim onShared As Boolean
                            Dim finalPath As String =
                                ResolveAndNormalize(currentPath, db, ownerFolder, FindFileHint.XRefDrawing, onShared)

                            If String.IsNullOrWhiteSpace(finalPath) Then
                                ' Detach only truly missing/broken references. Resolved-but-unverified
                                ' and Unloaded (turned off) xrefs are kept exactly as they are.
                                xrefIssues += 1
                                If Not isMissing Then
                                    ed.WriteMessage(vbLf & "Kept (" & status.ToString() & ", path unverified): " & btr.Name)
                                ElseIf IsTopLevelXref(node, db) Then
                                    ' db.DetachXref only works on xrefs attached directly to the host.
                                    xrefsToDetach.Add(btr.ObjectId)
                                    ed.WriteMessage(vbLf & "Missing xref will be detached: " & btr.Name & " | " & If(currentPath, "(no path)"))
                                Else
                                    ' Nested xref cannot be detached from the host; leave it for its owner.
                                    ed.WriteMessage(vbLf & "Missing nested xref (skipped, detach from its parent drawing): " & btr.Name & " | " & If(currentPath, "(no path)"))
                                End If
                                Continue For
                            End If

                            ' Flag files that resolved somewhere other than the shared drive.
                            If Not onShared Then
                                xrefIssues += 1
                                ed.WriteMessage(vbLf & "Kept (not on " & SharedDriveLetter & "\ shared path): " & btr.Name & " | " & finalPath)
                            End If

                            ' Rewrite the stored path to the verified absolute path only when it
                            ' changes. Successful remaps are silent; only failures are logged.
                            If Not String.Equals(currentPath, finalPath, StringComparison.OrdinalIgnoreCase) Then
                                btr.PathName = finalPath
                                xrefRemapped += 1
                            End If

                        Catch ex As Exception
                            xrefIssues += 1
                            ed.WriteMessage(vbLf & "Skipped an xref due to error: " & ex.Message)
                        End Try

                    Next

                    tr.Commit()

                End Using

                If xrefsToDetach.Count > 0 Then
                    For Each id As ObjectId In xrefsToDetach
                        Try
                            db.DetachXref(id)
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & "Could not detach xref: " & id.ToString() & " | " & ex.Message)
                        End Try
                    Next
                End If

                Try
                    db.ResolveXrefs(True, False)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "ResolveXrefs failed: " & ex.Message)
                End Try

                ' Restore the load state: ResolveXrefs reloads everything, so put back the ones
                ' the user had unloaded (turned off) before running xsr. Unload one at a time so
                ' a single bad id can't stop the rest from being restored.
                For Each id As ObjectId In xrefsWereUnloaded
                    Try
                        Dim oneId As New ObjectIdCollection()
                        oneId.Add(id)
                        db.UnloadXrefs(oneId)
                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "Could not restore unloaded xref state: " & id.ToString() & " | " & ex.Message)
                    End Try
                Next

                ' Apply the same resolve + E:\Shared normalization to PDF underlays.
                NormalizePdfUnderlayPaths(db, ed, pdfRemapped, pdfIssues)

            End Using

            ed.WriteMessage(vbLf & String.Format(
                "Finished xref cleanup: {0} xref + {1} PDF path(s) remapped, {2} issue(s).",
                xrefRemapped, pdfRemapped, xrefIssues + pdfIssues))

        End Sub


        ' Resolves and normalizes every host-level PDF underlay path the same way xsr
        ' handles DWG xrefs: locate the real file, then rewrite the stored path onto the
        ' canonical E:\Shared drive when it maps to that share. Missing files and files
        ' that aren't on the shared drive are left untouched and logged. Nested underlays
        ' living inside xref'd drawings are not touched (they belong to those drawings).
        Private Shared Sub NormalizePdfUnderlayPaths(db As Database, ed As Editor,
                                                     ByRef remapped As Integer, ByRef issues As Integer)

            Dim hostFolder As String = Nothing
            Try
                hostFolder = Path.GetDirectoryName(db.Filename)
            Catch
            End Try

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim namedDic As DBDictionary =
                        TryCast(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
                    If namedDic Is Nothing Then
                        tr.Commit()
                        Return
                    End If

                    Dim pdfDicKey As String = UnderlayDefinition.GetDictionaryKey(GetType(PdfDefinition))
                    If String.IsNullOrWhiteSpace(pdfDicKey) OrElse Not namedDic.Contains(pdfDicKey) Then
                        tr.Commit()
                        Return
                    End If

                    Dim pdfDic As DBDictionary =
                        TryCast(tr.GetObject(namedDic.GetAt(pdfDicKey), OpenMode.ForRead), DBDictionary)
                    If pdfDic Is Nothing Then
                        tr.Commit()
                        Return
                    End If

                    For Each entry As DBDictionaryEntry In pdfDic

                        ' Isolate each underlay so one bad entry can't abort the whole pass
                        ' (and lose every other fix when the commit is skipped).
                        Try

                            Dim pdfDef As PdfDefinition =
                                TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), PdfDefinition)
                            If pdfDef Is Nothing Then Continue For

                            Dim currentPath As String = pdfDef.SourceFileName
                            If String.IsNullOrWhiteSpace(currentPath) Then Continue For

                            ' Same resolve + E:\Shared normalization used for DWG xrefs.
                            Dim onShared As Boolean
                            Dim finalPath As String =
                                ResolveAndNormalize(currentPath, db, hostFolder, FindFileHint.[Default], onShared)

                            If String.IsNullOrWhiteSpace(finalPath) Then
                                issues += 1
                                ed.WriteMessage(vbLf & "PDF not found on this machine or " & SharedDriveLetter & "\ (left as-is): " & entry.Key & " | " & currentPath)
                                Continue For
                            End If

                            If Not onShared Then
                                issues += 1
                                ed.WriteMessage(vbLf & "Kept PDF (not on " & SharedDriveLetter & "\ shared path): " & entry.Key & " | " & finalPath)
                            End If

                            ' Successful remaps are silent; only failures are logged.
                            If Not String.Equals(currentPath, finalPath, StringComparison.OrdinalIgnoreCase) Then
                                pdfDef.UpgradeOpen()
                                pdfDef.SourceFileName = finalPath
                                pdfDef.DowngradeOpen()
                                remapped += 1
                            End If

                        Catch ex As Exception
                            issues += 1
                            ed.WriteMessage(vbLf & "Skipped a PDF underlay due to error: " & ex.Message)
                        End Try

                    Next

                    tr.Commit()

                Catch ex As Exception
                    ed.WriteMessage(vbLf & "PDF underlay normalization failed: " & ex.Message)
                End Try
            End Using

        End Sub


        ' Shared "resolve then normalize" step used by both the DWG xref pass and the PDF
        ' underlay pass so the two stay in sync. Resolves the stored path to a verified
        ' absolute file, then normalizes it onto the shared drive. Returns Nothing when the
        ' file can't be located anywhere; otherwise returns the final path and sets
        ' onShared to whether that path lives on E:\Shared (False means it resolved to some
        ' other, non-shared location and the caller should flag it).
        Private Shared Function ResolveAndNormalize(
        storedPath As String,
        db As Database,
        ownerFolder As String,
        hint As FindFileHint,
        ByRef onShared As Boolean
    ) As String

            onShared = False

            Dim finalPath As String = ResolveXrefAbsolutePath(storedPath, db, ownerFolder, hint)
            If String.IsNullOrWhiteSpace(finalPath) Then Return Nothing

            ' NormalizeToSharedDrive only rewrites when the path maps onto the shared drive,
            ' so a changed result is by definition on E:\Shared; otherwise check directly.
            Dim normalizedPath As String = NormalizeToSharedDrive(finalPath)
            If Not String.Equals(normalizedPath, finalPath, StringComparison.OrdinalIgnoreCase) Then
                finalPath = normalizedPath
                onShared = True
            Else
                onShared = IsOnSharedDrive(finalPath)
            End If

            Return finalPath
        End Function

        ' Resolves the xref file to an absolute path, preferring the canonical shared
        ' location so every machine stores the same path:
        ' 1) known shared equivalents of the stored path (root remaps + E:\Shared folder
        '    rebase) when the file is actually there, 2) the stored path as-is, 3) the
        ' engine's own file search (honors support/project paths), then 4) relative to
        ' the owning drawing folder. Returns Nothing when the file cannot be located.
        Private Shared Function ResolveXrefAbsolutePath(
        storedPath As String,
        db As Database,
        ownerFolder As String,
        Optional hint As FindFileHint = FindFileHint.XRefDrawing
    ) As String

            If String.IsNullOrWhiteSpace(storedPath) Then Return Nothing

            ' The stored path may point at the shared storage through a drive letter or a
            ' migrated root that isn't valid on this machine (e.g. W:\ from another user's
            ' setup, or a folder that was later moved under E:\Shared). Because every
            ' machine reaches that storage through the shared drive, try the known shared
            ' equivalents first and use one when the file is actually there. This is
            ' self-verifying, so a candidate that doesn't exist simply falls through.
            For Each candidate As String In GetSharedDriveCandidates(storedPath)
                If SafeFileExists(candidate) Then
                    Try
                        Return Path.GetFullPath(candidate)
                    Catch
                        Return candidate
                    End Try
                End If
            Next

            If Path.IsPathRooted(storedPath) AndAlso SafeFileExists(storedPath) Then
                Try
                    Return Path.GetFullPath(storedPath)
                Catch
                    Return storedPath
                End Try
            End If

            Try
                Dim found As String =
                HostApplicationServices.Current.FindFile(storedPath, db, hint)
                If Not String.IsNullOrWhiteSpace(found) AndAlso SafeFileExists(found) Then
                    Return Path.GetFullPath(found)
                End If
            Catch
                ' FindFile throws when the file cannot be located; fall through.
            End Try

            If Not Path.IsPathRooted(storedPath) AndAlso Not String.IsNullOrWhiteSpace(ownerFolder) Then
                Try
                    Dim combined As String = Path.GetFullPath(Path.Combine(ownerFolder, storedPath))
                    If SafeFileExists(combined) Then Return combined
                Catch
                End Try
            End If

            Return Nothing
        End Function

        ' File.Exists returns False (or throws) on UNC/network access errors, not just for
        ' genuinely missing files. This wrapper keeps those transient failures from being
        ' treated as "missing" and detaching valid xrefs.
        Private Shared Function SafeFileExists(path As String) As Boolean
            If String.IsNullOrWhiteSpace(path) Then Return False
            Try
                Return File.Exists(path)
            Catch
                Return False
            End Try
        End Function

        ' Win32 lookup that returns the UNC target (\\server\share) a mapped drive letter
        ' points at. Returns a non-zero error code when the drive is local or unmapped.
        <DllImport("mpr.dll", CharSet:=CharSet.Auto)>
        Private Shared Function WNetGetConnection(
            <MarshalAs(UnmanagedType.LPTStr)> localName As String,
            <MarshalAs(UnmanagedType.LPTStr)> remoteName As System.Text.StringBuilder,
            ByRef length As Integer
        ) As Integer
        End Function

        ' Returns the UNC target a drive letter maps to (e.g. "E:" -> "\\fs\share"), or
        ' Nothing when the drive is local/unmapped. Any trailing backslash is trimmed.
        Private Shared Function GetDriveUncTarget(driveWithColon As String) As String
            Try
                Dim sb As New System.Text.StringBuilder(1024)
                Dim length As Integer = sb.Capacity
                Dim result As Integer = WNetGetConnection(driveWithColon, sb, length)
                If result = 0 AndAlso sb.Length > 0 Then
                    Return sb.ToString().TrimEnd("\"c)
                End If
            Catch
            End Try
            Return Nothing
        End Function

        ' Expands a drive-letter path to its UNC-equivalent full path so two paths that
        ' reach the same share through different drive letters can be compared. Returns
        ' the path unchanged when it is already UNC, and Nothing when it is a local path
        ' (no mappable UNC target).
        Private Shared Function ToUncEquivalentPath(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then Return Nothing
            If path.StartsWith("\\", StringComparison.Ordinal) Then Return path
            If path.Length >= 2 AndAlso path(1) = ":"c Then
                Dim unc As String = GetDriveUncTarget(path.Substring(0, 2))
                If Not String.IsNullOrWhiteSpace(unc) Then
                    Dim remainder As String = path.Substring(2)
                    If Not remainder.StartsWith("\", StringComparison.Ordinal) Then remainder = "\" & remainder
                    Return unc & remainder
                End If
            End If
            Return Nothing
        End Function

        ' True when the absolute path already lives on the canonical shared drive, either
        ' directly (E:\...) or via another route that maps to the same network share.
        Private Shared Function IsOnSharedDrive(absolutePath As String) As Boolean
            If String.IsNullOrWhiteSpace(absolutePath) Then Return False
            If absolutePath.StartsWith(SharedDriveLetter & "\", StringComparison.OrdinalIgnoreCase) Then Return True

            Dim sharedUnc As String = GetSharedUnc()
            If String.IsNullOrWhiteSpace(sharedUnc) Then Return False

            Dim pathUnc As String = ToUncEquivalentPath(absolutePath)
            If String.IsNullOrWhiteSpace(pathUnc) Then Return False

            Return pathUnc.Equals(sharedUnc, StringComparison.OrdinalIgnoreCase) OrElse
                   pathUnc.StartsWith(sharedUnc & "\", StringComparison.OrdinalIgnoreCase)
        End Function

        ' Rewrites an absolute path onto the canonical shared drive when it points at the
        ' same network share as that drive (reached through a different drive letter or a
        ' raw UNC path). Paths not on the shared share are returned unchanged.
        Private Shared Function NormalizeToSharedDrive(absolutePath As String) As String
            If String.IsNullOrWhiteSpace(absolutePath) Then Return absolutePath

            Dim sharedUnc As String = GetSharedUnc()
            ' Shared drive isn't a mapped network drive here (or lookup failed); leave the
            ' path as resolved rather than guess.
            If String.IsNullOrWhiteSpace(sharedUnc) Then Return absolutePath

            Dim pathUnc As String = ToUncEquivalentPath(absolutePath)
            If String.IsNullOrWhiteSpace(pathUnc) Then Return absolutePath

            If pathUnc.Equals(sharedUnc, StringComparison.OrdinalIgnoreCase) Then
                Return SharedDriveLetter & "\"
            End If

            If pathUnc.StartsWith(sharedUnc & "\", StringComparison.OrdinalIgnoreCase) Then
                ' Remainder keeps its leading backslash, so this yields e.g. "E:\Shared\...".
                Dim remainder As String = pathUnc.Substring(sharedUnc.Length)
                Return SharedDriveLetter & remainder
            End If

            Return absolutePath
        End Function

        ' Returns the folder names directly under E:\Shared (e.g. "Arcxis", "Engineering",
        ' "Migrated from Other Sources"), read from disk once and cached. Empty when the
        ' shared drive isn't available on this machine, which simply disables the segment
        ' rebase below rather than guessing.
        Private Shared Function GetSharedTopFolders() As HashSet(Of String)
            If _sharedTopFolders IsNot Nothing Then Return _sharedTopFolders

            Dim folders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Try
                If Directory.Exists(SharedRoot) Then
                    For Each dir As String In Directory.GetDirectories(SharedRoot)
                        Dim name As String = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
                        If Not String.IsNullOrWhiteSpace(name) Then folders.Add(name)
                    Next
                End If
            Catch
                ' Leave the set empty on any access error; segment rebase just won't fire.
            End Try

            _sharedTopFolders = folders
            Return _sharedTopFolders
        End Function

        ' Builds the candidate shared-storage locations for a stored path, in priority
        ' order: first explicit root remaps (Egnyte UNC / legacy drives), then a rebase
        ' onto E:\Shared whenever the stored path contains one of E:\Shared's real
        ' top-level folders as a segment. Callers verify each candidate on disk, so these
        ' are only guesses at where the file really lives - a candidate that doesn't exist
        ' is skipped and nothing gets rewritten to a non-existent path.
        Private Shared Function GetSharedDriveCandidates(storedPath As String) As List(Of String)
            Dim candidates As New List(Of String)()
            If String.IsNullOrWhiteSpace(storedPath) Then Return candidates

            For Each remap As KeyValuePair(Of String, String) In ForeignRootRemaps
                If storedPath.StartsWith(remap.Key, StringComparison.OrdinalIgnoreCase) Then
                    candidates.Add(remap.Value & storedPath.Substring(remap.Key.Length))
                End If
            Next

            Dim topFolders As HashSet(Of String) = GetSharedTopFolders()
            If topFolders.Count > 0 Then
                Dim segments() As String = storedPath.Split({"\"c, "/"c}, StringSplitOptions.RemoveEmptyEntries)
                For i As Integer = 0 To segments.Length - 1
                    If Not topFolders.Contains(segments(i)) Then Continue For
                    Dim rebased As New System.Text.StringBuilder(SharedRoot)
                    For j As Integer = i To segments.Length - 1
                        rebased.Append("\"c).Append(segments(j))
                    Next
                    candidates.Add(rebased.ToString())
                Next
            End If

            Return candidates
        End Function

        ' True when the xref is attached directly to the host drawing (i.e. one of its
        ' incoming graph edges comes from the host database). Only such xrefs can be
        ' removed with db.DetachXref; nested xrefs must be detached from their parent.
        Private Shared Function IsTopLevelXref(node As XrefGraphNode, hostDb As Database) As Boolean
            If node Is Nothing OrElse hostDb Is Nothing Then Return False
            Try
                For j As Integer = 0 To node.NumIn - 1
                    Dim parent As XrefGraphNode = TryCast(node.In(j), XrefGraphNode)
                    If parent Is Nothing Then Continue For
                    If ReferenceEquals(parent.Database, hostDb) Then Return True
                Next
            Catch
            End Try
            Return False
        End Function

        Private Shared Function GetXrefOwnerFolder(
        node As XrefGraphNode,
        hostDb As Database
    ) As String

            Try
                Dim parentNode As XrefGraphNode =
                TryCast(node.In(0), XrefGraphNode)

                If parentNode Is Nothing Then
                    Return Path.GetDirectoryName(hostDb.Filename)
                End If

                If parentNode.Database Is Nothing Then
                    Return Path.GetDirectoryName(hostDb.Filename)
                End If

                If String.IsNullOrWhiteSpace(parentNode.Database.Filename) Then
                    Return Path.GetDirectoryName(hostDb.Filename)
                End If

                Return Path.GetDirectoryName(parentNode.Database.Filename)

            Catch
                Return Path.GetDirectoryName(hostDb.Filename)
            End Try

        End Function
    End Class
End Namespace

