Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Text.RegularExpressions
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.CleanupSubmissionCommands))>
Namespace Arcxis_Cad_Tools

    Friend Class CleanupPoint3d
        Public Property X As Double
        Public Property Y As Double
        Public Property Z As Double

        Public Shared Function FromPoint3d(pt As Point3d) As CleanupPoint3d
            Return New CleanupPoint3d With {
                .X = pt.X,
                .Y = pt.Y,
                .Z = pt.Z
            }
        End Function
    End Class

    Friend Class CleanupBounds
        Public Property Min As CleanupPoint3d
        Public Property Max As CleanupPoint3d

        Public Shared Function TryFromEntity(ent As Entity) As CleanupBounds
            Try
                Dim ext = ent.GeometricExtents
                Return New CleanupBounds With {
                    .Min = CleanupPoint3d.FromPoint3d(ext.MinPoint),
                    .Max = CleanupPoint3d.FromPoint3d(ext.MaxPoint)
                }
            Catch
                Return Nothing
            End Try
        End Function
    End Class

    Friend Class CleanupAttributeCapture
        Public Property Tag As String
        Public Property Value As String
    End Class

    Friend Class CleanupBlockChildCapture
        Public Property EntityType As String
        Public Property Layer As String
        Public Property TextString As String
        Public Property TextStringNormalized As String
    End Class

    Friend Class CleanupBlockDefinitionCapture
        Public Property Name As String
        Public Property IsAnonymous As Boolean
        Public Property OnlyText As Boolean
        Public Property HasMatchingTextPattern As Boolean
        Public Property Children As List(Of CleanupBlockChildCapture)
    End Class

    Friend Class CleanupInsertCapture
        Public Property BlockName As String
        Public Property BlockNameIsAnonymous As Boolean
        Public Property Rotation As Double
        Public Property Scale As CleanupPoint3d
        Public Property Attributes As List(Of CleanupAttributeCapture)
        Public Property BlockDefinition As CleanupBlockDefinitionCapture
    End Class

    Friend Class CleanupTextCapture
        Public Property TextString As String
        Public Property TextStringNormalized As String
        Public Property Height As Double
        Public Property Rotation As Double
        Public Property Style As String
    End Class

    Friend Class CleanupEntityCapture
        Public Property CaptureId As String
        Public Property Handle As String
        Public Property EntityType As String
        Public Property DxfName As String
        Public Property Layer As String
        Public Property ColorIndex As Integer
        Public Property Linetype As String
        Public Property Origin As CleanupPoint3d
        Public Property Bounds As CleanupBounds
        Public Property Insert As CleanupInsertCapture
        Public Property Text As CleanupTextCapture
        Public Property MText As CleanupTextCapture
        Public Property Radius As Double?
        Public Property VertexCount As Integer?
        Public Property Closed As Boolean?
        Public Property UserNote As String
    End Class

    ' Geometric composition of a loose multi-entity symbol — mirrors the Python catalog
    ' matcher's effective type counts + primary layers + geometry anchors so a selection
    ' drawn inline (not inserted as a block) can be matched by shape rather than block name.
    Friend Class CleanupComposition
        Public Property EntityCount As Integer
        Public Property EffectiveTypeCounts As Dictionary(Of String, Integer)
        Public Property Layers As Dictionary(Of String, Integer)
        Public Property AnchorRadii As List(Of Double)
        Public Property Bounds As CleanupBounds
    End Class

    Friend Class CleanupMatchCriteria
        Public Property MatchKind As String
        Public Property EntityType As String
        Public Property Layer As String
        Public Property LayerMatch As String
        Public Property BlockName As String
        Public Property BlockNameMatch As String
        Public Property BlockNamePattern As String
        Public Property TextPattern As String
        Public Property BlockContentsOnlyEntityTypes As List(Of String)
        Public Property BlockContentsAnyTextMatches As String
        Public Property Composition As CleanupComposition
        Public Property RadiusTolerance As Double?
    End Class

    Friend Class CleanupProposedRule
        Public Property RuleId As String
        Public Property DerivedFromCaptureIds As List(Of String)
        Public Property Action As String
        Public Property Priority As Integer
        Public Property Description As String
        Public Property Enabled As Boolean
        Public Property Status As String
        Public Property Match As CleanupMatchCriteria
    End Class

    Friend Class CleanupSourceDrawing
        Public Property FileName As String
        Public Property FullPath As String
        Public Property Layout As String
    End Class

    Friend Class CleanupSubmissionAttachments
        Public Property EvidenceDxf As String
        Public Property EvidenceBlockName As String
    End Class

    Friend Class CleanupSubmission
        Public Property SchemaVersion As String
        Public Property SubmissionId As String
        Public Property SubmittedAt As String
        Public Property SubmittedBy As String
        Public Property SourceDrawing As CleanupSourceDrawing
        Public Property Entities As List(Of CleanupEntityCapture)
        Public Property ProposedRules As List(Of CleanupProposedRule)
        Public Property Attachments As CleanupSubmissionAttachments
    End Class

    Friend NotInheritable Class CleanupSubmissionCapture
        Private Const DefaultTextPattern As String = "^(?i)\s*[^-]+\s*-\s*[^-]+\s*-\s*[^-]+\s*$"

        Public Shared Function CaptureEntities(ids As IEnumerable(Of ObjectId), tr As Transaction) As List(Of CleanupEntityCapture)
            Dim results As New List(Of CleanupEntityCapture)()

            For Each id In ids
                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing OrElse ent.IsErased Then Continue For
                results.Add(CaptureEntity(ent, tr))
            Next

            Return results
        End Function

        Private Shared Function CaptureEntity(ent As Entity, tr As Transaction) As CleanupEntityCapture
            Dim capture As New CleanupEntityCapture With {
                .CaptureId = Guid.NewGuid().ToString("D"),
                .Handle = ent.Handle.ToString(),
                .EntityType = NormalizeEntityType(ent),
                .DxfName = ent.GetType().Name,
                .Layer = ent.Layer,
                .ColorIndex = ent.ColorIndex,
                .Linetype = ent.Linetype,
                .Origin = CleanupPoint3d.FromPoint3d(GetOrigin(ent)),
                .Bounds = CleanupBounds.TryFromEntity(ent)
            }

            Dim txt = TryCast(ent, DBText)
            If txt IsNot Nothing Then
                capture.Text = New CleanupTextCapture With {
                    .TextString = txt.TextString,
                    .TextStringNormalized = NormalizeText(txt.TextString),
                    .Height = txt.Height,
                    .Rotation = txt.Rotation,
                    .Style = ResolveTextStyleName(txt.TextStyleId, tr)
                }
            End If

            Dim mt = TryCast(ent, MText)
            If mt IsNot Nothing Then
                capture.MText = New CleanupTextCapture With {
                    .TextString = mt.Contents,
                    .TextStringNormalized = NormalizeText(mt.Text),
                    .Height = mt.TextHeight,
                    .Rotation = mt.Rotation,
                    .Style = ResolveTextStyleName(mt.TextStyleId, tr)
                }
            End If

            Dim br = TryCast(ent, BlockReference)
            If br IsNot Nothing Then
                capture.Insert = CaptureInsert(br, tr)
            End If

            ' Geometric detail used to build a composition match for loose symbols.
            Dim circ = TryCast(ent, Circle)
            If circ IsNot Nothing Then capture.Radius = circ.Radius

            Dim arcEnt = TryCast(ent, Arc)
            If arcEnt IsNot Nothing Then capture.Radius = arcEnt.Radius

            Dim pl = TryCast(ent, Polyline)
            If pl IsNot Nothing Then
                capture.VertexCount = pl.NumberOfVertices
                capture.Closed = pl.Closed
            End If

            Return capture
        End Function

        Private Shared Function CaptureInsert(br As BlockReference, tr As Transaction) As CleanupInsertCapture
            Dim btr = TryCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
            Dim blockName = If(btr Is Nothing, String.Empty, btr.Name)

            Return New CleanupInsertCapture With {
                .BlockName = blockName,
                .BlockNameIsAnonymous = IsAnonymousBlockName(blockName),
                .Rotation = br.Rotation,
                .Scale = CleanupPoint3d.FromPoint3d(New Point3d(br.ScaleFactors.X, br.ScaleFactors.Y, br.ScaleFactors.Z)),
                .Attributes = CaptureAttributes(br, tr),
                .BlockDefinition = CaptureBlockDefinition(btr, tr)
            }
        End Function

        Private Shared Function CaptureAttributes(br As BlockReference, tr As Transaction) As List(Of CleanupAttributeCapture)
            Dim attrs As New List(Of CleanupAttributeCapture)()

            For Each attId As ObjectId In br.AttributeCollection
                Dim att = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                If att Is Nothing Then Continue For
                attrs.Add(New CleanupAttributeCapture With {
                    .Tag = att.Tag,
                    .Value = att.TextString
                })
            Next

            Return attrs
        End Function

        Private Shared Function CaptureBlockDefinition(btr As BlockTableRecord, tr As Transaction) As CleanupBlockDefinitionCapture
            If btr Is Nothing Then Return Nothing

            Dim children As New List(Of CleanupBlockChildCapture)()
            Dim onlyText As Boolean = True
            Dim hasPattern As Boolean = False

            For Each entId As ObjectId In btr
                Dim subEnt = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                If subEnt Is Nothing Then Continue For

                Dim childType = NormalizeEntityType(subEnt)
                If childType <> "TEXT" AndAlso childType <> "MTEXT" Then
                    onlyText = False
                End If

                Dim child As New CleanupBlockChildCapture With {
                    .EntityType = childType,
                    .Layer = subEnt.Layer
                }

                Dim subTxt = TryCast(subEnt, DBText)
                If subTxt IsNot Nothing Then
                    child.TextString = subTxt.TextString
                    child.TextStringNormalized = NormalizeText(subTxt.TextString)
                    If Regex.IsMatch(subTxt.TextString.Trim(), DefaultTextPattern) Then hasPattern = True
                End If

                Dim subMt = TryCast(subEnt, MText)
                If subMt IsNot Nothing Then
                    child.TextString = subMt.Contents
                    child.TextStringNormalized = NormalizeText(subMt.Text)
                    If Regex.IsMatch(subMt.Text.Trim(), DefaultTextPattern) Then hasPattern = True
                End If

                children.Add(child)
            Next

            Return New CleanupBlockDefinitionCapture With {
                .Name = btr.Name,
                .IsAnonymous = IsAnonymousBlockName(btr.Name),
                .OnlyText = onlyText AndAlso children.Count > 0,
                .HasMatchingTextPattern = hasPattern,
                .Children = children
            }
        End Function

        Public Shared Function ProposeRules(captures As List(Of CleanupEntityCapture), description As String, matchMode As String) As List(Of CleanupProposedRule)
            ' INSERT- and text-only selections match reliably per entity (block name / text pattern).
            ' Any selection containing loose geometry is treated as ONE symbol and matched by its
            ' geometric composition (effective type counts + anchor radii + layers), mirroring the
            ' Python catalog matcher — one rule for the whole selection, not one broad rule per entity.
            Dim allBlockOrText = captures.Count > 0 AndAlso captures.All(
                Function(c) c.Insert IsNot Nothing OrElse c.Text IsNot Nothing OrElse c.MText IsNot Nothing)

            If Not allBlockOrText Then
                Return New List(Of CleanupProposedRule) From {
                    ProposeCompositionRule(captures, description)
                }
            End If

            Return ProposeEntityRules(captures, description, matchMode)
        End Function

        Private Shared Function ProposeEntityRules(captures As List(Of CleanupEntityCapture), description As String, matchMode As String) As List(Of CleanupProposedRule)
            Dim rules As New List(Of CleanupProposedRule)()

            For Each capture In captures
                Dim match As New CleanupMatchCriteria With {
                    .MatchKind = "entity",
                    .EntityType = capture.EntityType,
                    .LayerMatch = "any"
                }

                If capture.Insert IsNot Nothing Then
                    If matchMode.Equals("Prefix", StringComparison.OrdinalIgnoreCase) AndAlso capture.Insert.BlockNameIsAnonymous Then
                        match.BlockNameMatch = "startsWith"
                        match.BlockNamePattern = "^A\$"
                    ElseIf matchMode.Equals("Regex", StringComparison.OrdinalIgnoreCase) Then
                        match.BlockNameMatch = "regex"
                        match.BlockNamePattern = Regex.Escape(capture.Insert.BlockName)
                    Else
                        match.BlockName = capture.Insert.BlockName
                        match.BlockNameMatch = "exact"
                    End If

                    If capture.Insert.BlockDefinition IsNot Nothing AndAlso
                       capture.Insert.BlockDefinition.OnlyText AndAlso
                       capture.Insert.BlockDefinition.HasMatchingTextPattern Then
                        match.BlockContentsOnlyEntityTypes = New List(Of String) From {"TEXT"}
                        match.BlockContentsAnyTextMatches = DefaultTextPattern
                    End If
                End If

                If capture.Text IsNot Nothing Then
                    If matchMode.Equals("Regex", StringComparison.OrdinalIgnoreCase) Then
                        match.TextPattern = DefaultTextPattern
                    Else
                        match.TextPattern = "^" & Regex.Escape(capture.Text.TextStringNormalized) & "$"
                    End If
                End If

                If capture.MText IsNot Nothing AndAlso String.IsNullOrWhiteSpace(match.TextPattern) Then
                    If matchMode.Equals("Regex", StringComparison.OrdinalIgnoreCase) Then
                        match.TextPattern = DefaultTextPattern
                    Else
                        match.TextPattern = "^" & Regex.Escape(capture.MText.TextStringNormalized) & "$"
                    End If
                End If

                rules.Add(New CleanupProposedRule With {
                    .RuleId = Guid.NewGuid().ToString("D"),
                    .DerivedFromCaptureIds = New List(Of String) From {capture.CaptureId},
                    .Action = "delete",
                    .Priority = 100,
                    .Description = description,
                    .Enabled = False,
                    .Status = "pending_review",
                    .Match = match
                })
            Next

            Return rules
        End Function

        Private Shared Function ProposeCompositionRule(captures As List(Of CleanupEntityCapture), description As String) As CleanupProposedRule
            Dim composition = BuildComposition(captures)
            Dim match As New CleanupMatchCriteria With {
                .MatchKind = "composition",
                .LayerMatch = "any",
                .Composition = composition,
                .RadiusTolerance = 0.5
            }

            Return New CleanupProposedRule With {
                .RuleId = Guid.NewGuid().ToString("D"),
                .DerivedFromCaptureIds = captures.Select(Function(c) c.CaptureId).ToList(),
                .Action = "delete",
                .Priority = 100,
                .Description = description,
                .Enabled = False,
                .Status = "pending_review",
                .Match = match
            }
        End Function

        Private Shared Function BuildComposition(captures As List(Of CleanupEntityCapture)) As CleanupComposition
            Dim typeCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim layers As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim radii As New List(Of Double)()
            Dim bounds As CleanupBounds = Nothing

            For Each capture In captures
                ' Effective type counts: a closed polyline with >= 4 vertices counts as that many
                ' LINEs, matching how loose copies are drawn (see Python _effective_type_counts).
                If String.Equals(capture.EntityType, "POLYLINE", StringComparison.OrdinalIgnoreCase) AndAlso
                   capture.Closed.GetValueOrDefault() AndAlso capture.VertexCount.GetValueOrDefault() >= 4 Then
                    AddCount(typeCounts, "LINE", capture.VertexCount.Value)
                Else
                    AddCount(typeCounts, capture.EntityType, 1)
                End If

                If Not String.IsNullOrEmpty(capture.Layer) Then
                    AddCount(layers, capture.Layer, 1)
                End If

                If capture.Radius.HasValue Then
                    radii.Add(Math.Round(capture.Radius.Value, 3))
                End If

                bounds = MergeBounds(bounds, capture.Bounds)
            Next

            Return New CleanupComposition With {
                .EntityCount = captures.Count,
                .EffectiveTypeCounts = typeCounts,
                .Layers = layers,
                .AnchorRadii = radii.Distinct().OrderBy(Function(r) r).ToList(),
                .Bounds = bounds
            }
        End Function

        Private Shared Sub AddCount(counts As Dictionary(Of String, Integer), key As String, amount As Integer)
            Dim current As Integer = 0
            counts.TryGetValue(key, current)
            counts(key) = current + amount
        End Sub

        Private Shared Function MergeBounds(existing As CleanupBounds, addition As CleanupBounds) As CleanupBounds
            If addition Is Nothing OrElse addition.Min Is Nothing OrElse addition.Max Is Nothing Then Return existing
            If existing Is Nothing Then
                Return New CleanupBounds With {
                    .Min = New CleanupPoint3d With {.X = addition.Min.X, .Y = addition.Min.Y, .Z = addition.Min.Z},
                    .Max = New CleanupPoint3d With {.X = addition.Max.X, .Y = addition.Max.Y, .Z = addition.Max.Z}
                }
            End If

            existing.Min.X = Math.Min(existing.Min.X, addition.Min.X)
            existing.Min.Y = Math.Min(existing.Min.Y, addition.Min.Y)
            existing.Min.Z = Math.Min(existing.Min.Z, addition.Min.Z)
            existing.Max.X = Math.Max(existing.Max.X, addition.Max.X)
            existing.Max.Y = Math.Max(existing.Max.Y, addition.Max.Y)
            existing.Max.Z = Math.Max(existing.Max.Z, addition.Max.Z)
            Return existing
        End Function

        Private Shared Function NormalizeEntityType(ent As Entity) As String
            If TypeOf ent Is Line Then Return "LINE"
            If TypeOf ent Is Circle Then Return "CIRCLE"
            If TypeOf ent Is Arc Then Return "ARC"
            If TypeOf ent Is DBText Then Return "TEXT"
            If TypeOf ent Is MText Then Return "MTEXT"
            If TypeOf ent Is BlockReference Then Return "INSERT"
            If TypeOf ent Is Polyline OrElse TypeOf ent Is Polyline2d Then Return "POLYLINE"
            If TypeOf ent Is Hatch Then Return "HATCH"
            Return ent.GetType().Name.ToUpperInvariant()
        End Function

        Private Shared Function GetOrigin(ent As Entity) As Point3d
            Dim ln = TryCast(ent, Line)
            If ln IsNot Nothing Then Return ln.StartPoint

            Dim txt = TryCast(ent, DBText)
            If txt IsNot Nothing Then Return txt.Position

            Dim mt = TryCast(ent, MText)
            If mt IsNot Nothing Then Return mt.Location

            Dim br = TryCast(ent, BlockReference)
            If br IsNot Nothing Then Return br.Position

            Try
                Dim ext = ent.GeometricExtents
                Return New Point3d(
                    (ext.MinPoint.X + ext.MaxPoint.X) / 2.0,
                    (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0,
                    (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0)
            Catch
                Return Point3d.Origin
            End Try
        End Function

        Private Shared Function ResolveTextStyleName(styleId As ObjectId, tr As Transaction) As String
            If styleId.IsNull Then Return String.Empty
            Dim ts = TryCast(tr.GetObject(styleId, OpenMode.ForRead), TextStyleTableRecord)
            Return If(ts Is Nothing, String.Empty, ts.Name)
        End Function

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim().ToUpperInvariant()
        End Function

        Private Shared Function IsAnonymousBlockName(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                name.StartsWith("A$", StringComparison.OrdinalIgnoreCase)
        End Function
    End Class

    Friend NotInheritable Class CleanupSubmissionExport
        Private Shared ReadOnly JsonOptions As JsonSerializerOptions = CreateJsonOptions()

        Private Shared Function CreateJsonOptions() As JsonSerializerOptions
            Return New JsonSerializerOptions With {
                .WriteIndented = True,
                .DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            }
        End Function

        Public Shared Function GetDefaultOutputFolder() As String
            Dim baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Arcxis",
                "CleanupSubmissions")
            Directory.CreateDirectory(baseDir)
            Return baseDir
        End Function

        Public Shared Function ExportEvidenceDxf(sourceDb As Database, ids As ObjectIdCollection, outputPath As String, Optional blockName As String = Nothing) As String
            If String.IsNullOrWhiteSpace(blockName) Then
                blockName = "ARCXIS_CLEANUP_EVIDENCE"
            End If

            Using sideDb As New Database(True, False)
                sourceDb.Wblock(sideDb, ids, Point3d.Origin, DuplicateRecordCloning.Ignore)
                Dim actualBlockName = WrapModelSpaceEntitiesInBlock(sideDb, blockName)
                ' Write a true DXF (SaveAs writes DWG regardless of extension, which ezdxf cannot read).
                sideDb.DxfOut(outputPath, 16, DwgVersion.Current)
                Return actualBlockName
            End Using
        End Function

        Private Shared Function WrapModelSpaceEntitiesInBlock(db As Database, blockName As String) As String
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForWrite), BlockTable)
                Dim ms = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                Dim entityIds As New ObjectIdCollection()
                For Each entId As ObjectId In ms
                    entityIds.Add(entId)
                Next

                If entityIds.Count = 0 Then
                    tr.Commit()
                    Return SanitizeBlockName(blockName)
                End If

                Dim uniqueName = GetUniqueBlockName(bt, SanitizeBlockName(blockName))
                Dim btr As New BlockTableRecord With {
                    .Name = uniqueName,
                    .Origin = Point3d.Origin
                }
                bt.Add(btr)
                tr.AddNewlyCreatedDBObject(btr, True)

                ' Clone each entity into the block and erase the original. A same-database
                ' Entity.Clone() preserves layer/color/linetype (the layer record already
                ' lives in this side DB); WblockCloneObjects with a fresh IdMapping flattened
                ' cloned entities onto layer "0", losing the source layer the matcher needs.
                For Each entId As ObjectId In entityIds
                    Dim ent = TryCast(tr.GetObject(entId, OpenMode.ForWrite), Entity)
                    If ent Is Nothing OrElse ent.IsErased Then Continue For
                    Dim clone = TryCast(ent.Clone(), Entity)
                    If clone IsNot Nothing Then
                        btr.AppendEntity(clone)
                        tr.AddNewlyCreatedDBObject(clone, True)
                    End If
                    ent.Erase()
                Next

                Dim blkRef As New BlockReference(Point3d.Origin, btr.ObjectId)
                ms.AppendEntity(blkRef)
                tr.AddNewlyCreatedDBObject(blkRef, True)

                tr.Commit()
                Return uniqueName
            End Using
        End Function

        Private Shared Function SanitizeBlockName(name As String) As String
            Dim sanitized = Regex.Replace(name, "[^A-Za-z0-9_$-]", "_")
            If String.IsNullOrWhiteSpace(sanitized) Then Return "ARCXIS_CLEANUP_EVIDENCE"
            If sanitized.Length > 31 Then sanitized = sanitized.Substring(0, 31)
            Return sanitized
        End Function

        Private Shared Function GetUniqueBlockName(bt As BlockTable, baseName As String) As String
            If Not bt.Has(baseName) Then Return baseName

            Dim suffix = 1
            While bt.Has(baseName & "_" & suffix.ToString(CultureInfo.InvariantCulture))
                suffix += 1
            End While

            Dim candidate = baseName & "_" & suffix.ToString(CultureInfo.InvariantCulture)
            If candidate.Length > 31 Then
                candidate = baseName.Substring(0, Math.Max(1, 31 - suffix.ToString(CultureInfo.InvariantCulture).Length - 1)) &
                    "_" & suffix.ToString(CultureInfo.InvariantCulture)
            End If

            Return candidate
        End Function

        Public Shared Function WriteSubmissionBundle(submission As CleanupSubmission, ids As ObjectIdCollection, sourceDb As Database, outputFolder As String) As String
            Directory.CreateDirectory(outputFolder)

            Dim folderName = submission.SubmissionId
            Dim bundleFolder = Path.Combine(outputFolder, folderName)
            Directory.CreateDirectory(bundleFolder)

            Dim evidenceFileName = "evidence.dxf"
            Dim evidencePath = Path.Combine(bundleFolder, evidenceFileName)
            Dim blockName = "ARCXIS_CLEANUP_" & submission.SubmissionId.Replace("-", "").Substring(0, 8)
            Dim actualBlockName = ExportEvidenceDxf(sourceDb, ids, evidencePath, blockName)

            submission.Attachments = New CleanupSubmissionAttachments With {
                .EvidenceDxf = evidenceFileName,
                .EvidenceBlockName = actualBlockName
            }

            Dim manifestPath = Path.Combine(bundleFolder, "manifest.json")
            File.WriteAllText(manifestPath, JsonSerializer.Serialize(submission, JsonOptions))

            Dim zipPath = Path.Combine(outputFolder, folderName & ".cleanup-submission.zip")
            If File.Exists(zipPath) Then File.Delete(zipPath)
            ZipFile.CreateFromDirectory(bundleFolder, zipPath)
            Directory.Delete(bundleFolder, True)

            Return zipPath
        End Function
    End Class

    Friend Class CleanupSubmitConfig
        Public Property ServerUrl As String
        Public Property ApiKey As String
        Public Property ClientSlug As String
    End Class

    ''' <summary>
    ''' Sends a submission bundle to the central catalog server's intake endpoint
    ''' (POST {ServerUrl}/api/symbols/upload, multipart). Config lives at
    ''' %LOCALAPPDATA%\Arcxis\CleanupSubmit\config.json. Never throws — the local
    ''' bundle is always kept, so an offline/unconfigured machine still works.
    ''' </summary>
    Friend NotInheritable Class CleanupSubmissionUpload

        ' Per-user override (this machine / this user only).
        Private Shared Function ConfigPath() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Arcxis", "CleanupSubmit", "config.json")
        End Function

        ' Machine-wide (all users on this PC).
        Private Shared Function ProgramDataConfigPath() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Arcxis", "CleanupSubmit", "config.json")
        End Function

        ''' <summary>
        ''' The "set once for everyone" config beside the plugin's Support Files on the
        ''' Egnyte share. Built on Module_Arcxis_TB.NetworkUNCPathForEgnyte — the SAME UNC
        ''' base (\\egnytedrive\...\shared) the module already uses for Support Files — so it
        ''' is drive-letter-agnostic; only the folder tail is hardcoded.
        ''' </summary>
        Private Shared Function SharedConfigPath() As String
            Try
                If String.IsNullOrEmpty(Module_Arcxis_TB.NetworkUNCPathForEgnyte) Then
                    Module_Arcxis_TB.UNCPath()
                End If
                Dim base = Module_Arcxis_TB.NetworkUNCPathForEgnyte
                If String.IsNullOrEmpty(base) Then Return Nothing
                Return Path.Combine(base,
                    "Arcxis", "Engineering", "Drafting Standards", "CAD Lisp Routines",
                    "BricsCad", "Support Files", "config.json")
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryReadConfig(p As String) As CleanupSubmitConfig
            Try
                If String.IsNullOrWhiteSpace(p) OrElse Not File.Exists(p) Then Return Nothing
                Dim cfg = JsonSerializer.Deserialize(Of CleanupSubmitConfig)(
                    File.ReadAllText(p),
                    New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})
                If cfg Is Nothing OrElse String.IsNullOrWhiteSpace(cfg.ServerUrl) Then Return Nothing
                Return cfg
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Resolves submit config in priority order — per-user override, then the shared
        ''' Egnyte file (org default), then machine-wide ProgramData, then environment
        ''' variables. First match wins; Nothing means local-only (no upload).
        ''' </summary>
        Public Shared Function LoadConfig() As CleanupSubmitConfig
            For Each p In {ConfigPath(), SharedConfigPath(), ProgramDataConfigPath()}
                Dim cfg = TryReadConfig(p)
                If cfg IsNot Nothing Then Return cfg
            Next
            Dim envUrl = Environment.GetEnvironmentVariable("ARCXIS_SUBMIT_SERVER_URL")
            If Not String.IsNullOrWhiteSpace(envUrl) Then
                Return New CleanupSubmitConfig With {
                    .ServerUrl = envUrl,
                    .ApiKey = Environment.GetEnvironmentVariable("ARCXIS_SUBMIT_API_KEY"),
                    .ClientSlug = Environment.GetEnvironmentVariable("ARCXIS_SUBMIT_CLIENT_SLUG")
                }
            End If
            Return Nothing
        End Function

        Public Shared Function TryUpload(zipPath As String, cfg As CleanupSubmitConfig) As Tuple(Of Boolean, String)
            Try
                Dim url = cfg.ServerUrl.TrimEnd("/"c) & "/api/symbols/upload"
                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(30)
                    Using form As New MultipartFormDataContent()
                        Dim fileContent As New ByteArrayContent(File.ReadAllBytes(zipPath))
                        fileContent.Headers.ContentType = New MediaTypeHeaderValue("application/zip")
                        form.Add(fileContent, "file", Path.GetFileName(zipPath))
                        If Not String.IsNullOrWhiteSpace(cfg.ClientSlug) Then
                            form.Add(New StringContent(cfg.ClientSlug), "clientSlug")
                        End If
                        Using req As New HttpRequestMessage(HttpMethod.Post, url)
                            req.Content = form
                            If Not String.IsNullOrWhiteSpace(cfg.ApiKey) Then
                                req.Headers.Add("X-Api-Key", cfg.ApiKey)
                            End If
                            Dim resp = client.Send(req)
                            If resp.IsSuccessStatusCode Then
                                Return Tuple.Create(True, "Uploaded to " & url)
                            End If
                            Return Tuple.Create(False, "Server responded " & CInt(resp.StatusCode).ToString())
                        End Using
                    End Using
                End Using
            Catch ex As System.Exception
                Return Tuple.Create(False, "Upload failed: " & ex.Message)
            End Try
        End Function
    End Class

    Public Class CleanupSubmissionCommands

        <CommandMethod("SUBMITCLEANUPRULE", CommandFlags.Modal)>
        Public Sub SubmitCleanupRule()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim selOpts As New PromptSelectionOptions With {
                .MessageForAdding = vbLf & "Select entities to submit as cleanup rule evidence: "
            }

            Dim selResult = ed.GetSelection(selOpts)
            If selResult.Status <> PromptStatus.OK Then Return

            Dim descOpts As New PromptStringOptions(vbLf & "Describe why this should be deleted: ") With {
                .AllowSpaces = True
            }
            Dim descResult = ed.GetString(descOpts)
            If descResult.Status <> PromptStatus.OK Then Return

            Dim kwOpts As New PromptKeywordOptions(vbLf & "Match mode [Exact/Prefix/Regex] <Exact>: ", "Exact Prefix Regex")
            kwOpts.Keywords.Default = "Exact"
            Dim kwResult = ed.GetKeywords(kwOpts)
            If kwResult.Status <> PromptStatus.OK Then Return

            Dim ids As New ObjectIdCollection()
            For Each id As SelectedObject In selResult.Value
                ids.Add(id.ObjectId)
            Next

            Using doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim captures = CleanupSubmissionCapture.CaptureEntities(
                        selResult.Value.GetObjectIds(), tr)

                    Dim submission As New CleanupSubmission With {
                        .SchemaVersion = "1.0",
                        .SubmissionId = Guid.NewGuid().ToString("D"),
                        .SubmittedAt = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                        .SubmittedBy = Environment.UserName,
                        .SourceDrawing = New CleanupSourceDrawing With {
                            .FileName = Path.GetFileName(db.Filename),
                            .FullPath = db.Filename,
                            .Layout = LayoutManager.Current.CurrentLayout
                        },
                        .Entities = captures,
                        .ProposedRules = CleanupSubmissionCapture.ProposeRules(captures, descResult.StringResult, kwResult.StringResult)
                    }

                    tr.Commit()

                    Dim zipPath = CleanupSubmissionExport.WriteSubmissionBundle(
                        submission, ids, db, CleanupSubmissionExport.GetDefaultOutputFolder())

                    ed.WriteMessage(vbLf & "Cleanup submission created:")
                    ed.WriteMessage(vbLf & zipPath)

                    ' Send to the central catalog server if configured; otherwise keep local only.
                    Dim submitCfg = CleanupSubmissionUpload.LoadConfig()
                    If submitCfg Is Nothing Then
                        ed.WriteMessage(vbLf & "(Saved locally. No submit-server configured — import on the server to add it to the catalog.)")
                    Else
                        Dim up = CleanupSubmissionUpload.TryUpload(zipPath, submitCfg)
                        If up.Item1 Then
                            ed.WriteMessage(vbLf & "Sent to catalog server: " & up.Item2)
                        Else
                            ed.WriteMessage(vbLf & "Saved locally; server upload skipped — " & up.Item2)
                        End If
                    End If
                End Using
            End Using
        End Sub

    End Class

End Namespace
