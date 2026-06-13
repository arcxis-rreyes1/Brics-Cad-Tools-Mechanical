Imports System.IO
Imports System.Threading
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Bricscad.ApplicationServices
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.ExportToCad))>
Namespace Arcxis_Cad_Tools

    Public Class ExportToCad

        Private Const MasterSealBlockName As String = "Master Seal File"
        Private Const MasterSealRelativePath As String = "Onyx File System\Templates\Engineering\SealsOriginal\Master Seal File.dwg"
        Private Const Fs2SealShareMarker As String = "\\fs2\seal$"

        Private Shared Function IsFs2SealSharePath(path As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(path) AndAlso
                path.IndexOf(Fs2SealShareMarker, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Sub RemoveFs2SealShareRasterImages(db As Database, tr As Transaction, bt As BlockTable)
            Try
                Dim imgDictId As ObjectId = RasterImageDef.GetImageDictionary(db)
                If imgDictId.IsNull Then Return

                Dim imgDict As DBDictionary = TryCast(tr.GetObject(imgDictId, OpenMode.ForRead), DBDictionary)
                If imgDict Is Nothing Then Return

                Dim keysToRemove As New List(Of String)()
                For Each entry As DBDictionaryEntry In imgDict
                    Dim defId As ObjectId = entry.Value
                    Dim defObj As RasterImageDef = TryCast(tr.GetObject(defId, OpenMode.ForRead), RasterImageDef)
                    If defObj Is Nothing Then Continue For

                    Dim src As String = If(defObj.SourceFileName, String.Empty)
                    If IsFs2SealSharePath(src) Then
                        keysToRemove.Add(entry.Key)
                    End If
                Next

                If keysToRemove.Count = 0 Then Return

                imgDict.UpgradeOpen()
                For Each key In keysToRemove
                    If Not imgDict.Contains(key) Then Continue For
                    Dim defId As ObjectId = imgDict.GetAt(key)

                    For Each btrId As ObjectId In bt
                        Dim btrRec As BlockTableRecord = TryCast(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                        If btrRec Is Nothing Then Continue For

                        For Each objId As ObjectId In btrRec
                            If objId.ObjectClass.DxfName <> "RASTERIMAGE" Then Continue For

                            Dim ri As RasterImage = TryCast(tr.GetObject(objId, OpenMode.ForWrite), RasterImage)
                            If ri IsNot Nothing AndAlso ri.ImageDefId = defId Then
                                ri.Erase()
                            End If
                        Next
                    Next

                    If Not defId.IsNull Then
                        Dim defToErase As DBObject = TryCast(tr.GetObject(defId, OpenMode.ForWrite), DBObject)
                        If defToErase IsNot Nothing Then defToErase.Erase()
                    End If

                    Try
                        imgDict.Remove(key)
                    Catch
                    End Try
                Next
                imgDict.DowngradeOpen()
            Catch
            End Try
        End Sub

        <CommandMethod("SealFix", CommandFlags.Modal)>
        Public Shared Sub RunSealFixCommand()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            FixMasterSealFileXrefPath(doc)
        End Sub

        Public Shared Sub FixMasterSealFileXrefPath(doc As Document)
            If doc Is Nothing Then Return

            If String.IsNullOrEmpty(Module_Arcxis_TB.NetworkUNCPathForEgnyte) Then
                Module_Arcxis_TB.UNCPath()
            End If

            If String.IsNullOrEmpty(Module_Arcxis_TB.NetworkUNCPathForEgnyte) Then Return

            Dim expectedPath As String = Path.Combine(
                Module_Arcxis_TB.NetworkUNCPathForEgnyte,
                MasterSealRelativePath)

            Dim db As Database = doc.Database
            Dim xrefIdsToReload As New List(Of ObjectId)()
            Dim xrefIdsToDetach As New List(Of ObjectId)()

            Using doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = TryCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    If bt Is Nothing Then
                        tr.Commit()
                        Return
                    End If

                    For Each id As ObjectId In bt
                        Dim btr As BlockTableRecord = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                        If btr Is Nothing OrElse Not btr.IsFromExternalReference Then Continue For

                        Dim currentPath As String = If(btr.PathName, String.Empty)
                        Dim isMasterSeal As Boolean =
                            String.Equals(btr.Name, MasterSealBlockName, StringComparison.OrdinalIgnoreCase)

                        If isMasterSeal Then
                            If String.Equals(currentPath, expectedPath, StringComparison.OrdinalIgnoreCase) Then Continue For

                            btr.UpgradeOpen()
                            btr.PathName = expectedPath
                            xrefIdsToReload.Add(btr.ObjectId)
                            Continue For
                        End If

                        If IsFs2SealSharePath(currentPath) Then
                            xrefIdsToDetach.Add(btr.ObjectId)
                        End If
                    Next

                    RemoveFs2SealShareRasterImages(db, tr, bt)

                    tr.Commit()
                End Using
            End Using

            For Each xrefId As ObjectId In xrefIdsToDetach
                Try
                    db.DetachXref(xrefId)
                Catch
                End Try
            Next

            If xrefIdsToReload.Count = 0 Then Return

            Dim chunk As New ObjectIdCollection()
            For Each xrefId As ObjectId In xrefIdsToReload
                chunk.Add(xrefId)
            Next

            Try
                db.ReloadXrefs(chunk)
            Catch
            End Try
        End Sub

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
            Module_Arcxis_TB.NetworkUNCPathForEgnyte &
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


        <CommandMethod("xsr")>
        Public Shared Sub MakeAllXrefPathsAbsolute()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            If String.IsNullOrWhiteSpace(db.Filename) Then
                ed.WriteMessage(vbLf & "Drawing must be saved first.")
                Return
            End If

            Dim xrefsToDetach As New List(Of ObjectId)

            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim xrefGraph As XrefGraph = db.GetHostDwgXrefGraph(True)

                For i As Integer = 1 To xrefGraph.NumNodes - 1

                    Dim node As XrefGraphNode = xrefGraph.GetXrefNode(i)
                    If node Is Nothing OrElse node.BlockTableRecordId.IsNull Then Continue For

                    Dim btr As BlockTableRecord =
                    TryCast(tr.GetObject(node.BlockTableRecordId, OpenMode.ForWrite, False), BlockTableRecord)

                    If btr Is Nothing OrElse Not btr.IsFromExternalReference Then Continue For

                    Dim currentPath As String = btr.PathName
                    If String.IsNullOrWhiteSpace(currentPath) Then
                        xrefsToDetach.Add(btr.ObjectId)
                        Continue For
                    End If

                    Dim ownerFolder As String = GetXrefOwnerFolder(node, db)

                    If String.IsNullOrWhiteSpace(ownerFolder) Then
                        ownerFolder = Path.GetDirectoryName(db.Filename)
                    End If

                    Dim finalPath As String = currentPath

                    If Not Path.IsPathRooted(currentPath) Then
                        finalPath = Path.GetFullPath(Path.Combine(ownerFolder, currentPath))
                        btr.PathName = finalPath

                        ed.WriteMessage(vbLf & btr.Name & ": " & currentPath & " -> " & finalPath)
                    End If

                    If Not File.Exists(finalPath) Then
                        xrefsToDetach.Add(btr.ObjectId)
                        ed.WriteMessage(vbLf & "Missing xref will be detached: " & btr.Name & " | " & finalPath)
                    End If

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
                ed.WriteMessage(vbLf & "Reloaded remaining xrefs.")
            Catch ex As Exception
                ed.WriteMessage(vbLf & "ResolveXrefs failed: " & ex.Message)
            End Try

            ed.WriteMessage(vbLf & "Finished xref cleanup.")

        End Sub


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

