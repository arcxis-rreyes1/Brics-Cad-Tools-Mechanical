Imports System
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.MirrorFlip))>
Namespace Arcxis_Cad_Tools

    ''' <summary>
    ''' Creates a mirrored copy of the selected geometry and shifts it over by a number
    ''' of pages (1800 units per page). The mirror axis is taken through the center of
    ''' the selection's bounding box. Framing blocks and storage notes have their scale
    ''' flipped so their text/symbols stay readable after the mirror.
    ''' Command: FC
    ''' </summary>
    Public Class MirrorFlip

        Private Const PageWidth As Double = 1800.0

        Private Shared _lastPagesOver As Integer = 0
        Private Shared _lastVerticalFlip As Boolean = True
        Private Shared _lastUsePickedCenter As Boolean = False

        ' Blocks whose X scale is forced to +1 when the (snapped) rotation is 0 or 270,
        ' otherwise the Y scale is forced to -1.
        Private Shared ReadOnly UseXScaleAt0And270 As String() = {
            "FR01 4", "FR01 5A", "FR01 5B", "FR01 6", "FR01 8", "FR01 10", "FR01 11"
        }

        ' Blocks whose X scale is forced to +1 when the (snapped) rotation is 0 or 90,
        ' otherwise the Y scale is forced to -1.
        Private Shared ReadOnly UseXScaleAt0And90 As String() = {
            "FR01 1", "FR01 2A", "FR01 2B", "FR01 7", "FR01 9",
            "LIMITEDATTICSTORAGENOTE", "CEILINGLIMITEDATTICSTORAGENOTE", "BEDROOMSTORAGENOTE"
        }

        Public Sub MirrorFlipCommand()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return

            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim pagesOver As Integer? = PromptPagesOver(ed)
            If Not pagesOver.HasValue Then
                ed.WriteMessage(vbLf & "Command cancelled — no page count entered.")
                Return
            End If
            _lastPagesOver = pagesOver.Value
            Dim distanceOver As Double = pagesOver.Value * PageWidth

            Dim verticalFlip As Boolean = PromptVerticalFlip(ed)
            _lastVerticalFlip = verticalFlip

            Dim usePickedCenter As Boolean = PromptUsePickedCenter(ed)
            _lastUsePickedCenter = usePickedCenter

            Using docLock As DocumentLock = doc.LockDocument()
                Do
                    Dim mirrorCenter As Point3d = Point3d.Origin
                    If usePickedCenter AndAlso Not PromptMirrorCenterFromBox(ed, mirrorCenter) Then Exit Do

                    Dim selectionResult As PromptSelectionResult =
                        ed.GetSelection(New PromptSelectionOptions() With {
                            .MessageForAdding = vbLf & "Select objects to MIRROR+SHIFT (Enter to finish): "
                        })

                    If selectionResult.Status <> PromptStatus.OK Then Exit Do

                    Dim mirroredCount As Integer = MirrorAndShift(
                        db, selectionResult.Value.GetObjectIds(), verticalFlip, distanceOver, usePickedCenter, mirrorCenter)

                    If mirroredCount = 0 Then
                        ed.WriteMessage(vbLf & "Nothing usable was selected.")
                    Else
                        ed.WriteMessage(vbLf & String.Format("{0} object(s) mirrored and shifted.", mirroredCount))
                    End If
                Loop
            End Using
        End Sub

        Private Shared Function MirrorAndShift(db As Database,
                                               sourceIds As ObjectId(),
                                               verticalFlip As Boolean,
                                               distanceOver As Double,
                                               usePickedCenter As Boolean,
                                               pickedCenter As Point3d) As Integer
            Dim mirroredCount As Integer = 0

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim center As Point3d = pickedCenter
                If Not usePickedCenter AndAlso Not TryGetMirrorCenter(tr, sourceIds, center) Then
                    tr.Commit()
                    Return 0
                End If

                Dim mirrorAxis As Line3d = BuildMirrorAxis(center, verticalFlip)
                Dim mirror As Matrix3d = Matrix3d.Mirroring(mirrorAxis)
                Dim shift As Matrix3d = Matrix3d.Displacement(New Vector3d(distanceOver, 0.0, 0.0))
                Dim transform As Matrix3d = shift * mirror

                Dim idsToClone As New ObjectIdCollection()
                For Each sourceId As ObjectId In sourceIds
                    idsToClone.Add(sourceId)
                Next

                Dim idMap As New IdMapping()
                db.DeepCloneObjects(idsToClone, db.CurrentSpaceId, idMap, False)

                For Each sourceId As ObjectId In sourceIds
                    If Not idMap.Contains(sourceId) Then Continue For

                    Dim pair As IdPair = idMap(sourceId)
                    If Not pair.IsCloned OrElse pair.Value.IsNull Then Continue For

                    Dim sourceEntity As Entity = TryCast(tr.GetObject(sourceId, OpenMode.ForRead), Entity)
                    Dim mirroredEntity As Entity = TryCast(tr.GetObject(pair.Value, OpenMode.ForWrite), Entity)
                    If sourceEntity Is Nothing OrElse mirroredEntity Is Nothing Then Continue For

                    Dim blockRef As BlockReference = TryCast(mirroredEntity, BlockReference)
                    If blockRef Is Nothing AndAlso TransformTextPositionPreservingOrientation(db, sourceEntity, mirroredEntity, transform, verticalFlip) Then
                        mirroredCount += 1
                        Continue For
                    End If

                    mirroredEntity.TransformBy(transform)

                    blockRef = TryCast(mirroredEntity, BlockReference)
                    If blockRef IsNot Nothing Then
                        FixInsertScales(blockRef)
                    Else
                        CorrectAnnotationTextOrientation(mirroredEntity)
                    End If

                    mirroredCount += 1
                Next

                tr.Commit()
            End Using

            Return mirroredCount
        End Function

        ''' <summary>
        ''' Determines the mirror center from the selection. Block references are skipped
        ''' first because a stray/off insert can have extents far from the visible cluster
        ''' and would drag the center off. The center is taken from line/polyline geometry,
        ''' falling back to any non-block geometry, then to everything as a last resort.
        ''' </summary>
        Private Shared Function TryGetMirrorCenter(tr As Transaction,
                                                   sourceIds As ObjectId(),
                                                   ByRef center As Point3d) As Boolean
            If TryGetCenterFromExtents(tr, sourceIds, AddressOf IsLineOrPolyline, center) Then Return True
            If TryGetCenterFromExtents(tr, sourceIds, AddressOf IsNotBlockReference, center) Then Return True
            Return TryGetCenterFromExtents(tr, sourceIds, Function(e As Entity) True, center)
        End Function

        ''' <summary>Combined bounding-box center of the entities accepted by the filter.</summary>
        Private Shared Function TryGetCenterFromExtents(tr As Transaction,
                                                        sourceIds As ObjectId(),
                                                        include As Func(Of Entity, Boolean),
                                                        ByRef center As Point3d) As Boolean
            Dim haveBox As Boolean = False
            Dim minPt As Point3d = Point3d.Origin
            Dim maxPt As Point3d = Point3d.Origin

            For Each id As ObjectId In sourceIds
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing OrElse Not include(ent) Then Continue For

                Dim ext As Extents3d
                Try
                    ext = ent.GeometricExtents
                Catch
                    Continue For
                End Try

                If Not haveBox Then
                    minPt = ext.MinPoint
                    maxPt = ext.MaxPoint
                    haveBox = True
                Else
                    minPt = New Point3d(Math.Min(minPt.X, ext.MinPoint.X),
                                        Math.Min(minPt.Y, ext.MinPoint.Y),
                                        Math.Min(minPt.Z, ext.MinPoint.Z))
                    maxPt = New Point3d(Math.Max(maxPt.X, ext.MaxPoint.X),
                                        Math.Max(maxPt.Y, ext.MaxPoint.Y),
                                        Math.Max(maxPt.Z, ext.MaxPoint.Z))
                End If
            Next

            If Not haveBox Then Return False

            center = New Point3d((minPt.X + maxPt.X) / 2.0,
                                 (minPt.Y + maxPt.Y) / 2.0,
                                 (minPt.Z + maxPt.Z) / 2.0)
            Return True
        End Function

        Private Shared Function IsLineOrPolyline(ent As Entity) As Boolean
            Return TypeOf ent Is Line _
                OrElse TypeOf ent Is Polyline _
                OrElse TypeOf ent Is Polyline2d _
                OrElse TypeOf ent Is Polyline3d
        End Function

        Private Shared Function IsNotBlockReference(ent As Entity) As Boolean
            Return Not (TypeOf ent Is BlockReference)
        End Function

        Private Shared Function PromptPagesOver(ed As Editor) As Integer?
            Dim prompt As New PromptIntegerOptions(vbLf & "How many pages over? <" & _lastPagesOver.ToString() & ">: ")
            prompt.AllowNone = True
            prompt.AllowZero = True
            prompt.AllowNegative = False
            prompt.DefaultValue = _lastPagesOver
            prompt.UseDefaultValue = True

            Dim result As PromptIntegerResult = ed.GetInteger(prompt)
            If result.Status = PromptStatus.None Then Return _lastPagesOver
            If result.Status <> PromptStatus.OK Then Return Nothing
            Return result.Value
        End Function

        Private Shared Function PromptVerticalFlip(ed As Editor) As Boolean
            Dim defaultKeyword As String = If(_lastVerticalFlip, "Yes", "No")
            Dim prompt As New PromptKeywordOptions(vbLf & "Are you doing a vertical flip? [Yes/No] <" & defaultKeyword & ">: ")
            prompt.Keywords.Add("Yes")
            prompt.Keywords.Add("No")
            prompt.Keywords.Default = defaultKeyword
            prompt.AllowNone = True

            Dim result As PromptResult = ed.GetKeywords(prompt)
            If result.Status <> PromptStatus.OK OrElse String.IsNullOrEmpty(result.StringResult) Then
                Return _lastVerticalFlip
            End If
            Return String.Equals(result.StringResult, "Yes", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function PromptUsePickedCenter(ed As Editor) As Boolean
            Dim defaultKeyword As String = If(_lastUsePickedCenter, "Box", "Auto")
            Dim prompt As New PromptKeywordOptions(vbLf & "Mirror center method [Auto/Box] <" & defaultKeyword & ">: ")
            prompt.Keywords.Add("Auto")
            prompt.Keywords.Add("Box")
            prompt.Keywords.Default = defaultKeyword
            prompt.AllowNone = True

            Dim result As PromptResult = ed.GetKeywords(prompt)
            If result.Status <> PromptStatus.OK OrElse String.IsNullOrEmpty(result.StringResult) Then
                Return _lastUsePickedCenter
            End If
            Return String.Equals(result.StringResult, "Box", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function PromptMirrorCenterFromBox(ed As Editor, ByRef center As Point3d) As Boolean
            Dim topLeftResult As PromptPointResult = ed.GetPoint(vbLf & "Select TOP LEFT corner for mirror center (Enter to finish): ")
            If topLeftResult.Status <> PromptStatus.OK Then Return False

            Dim bottomRightOptions As New PromptPointOptions(vbLf & "Select BOTTOM RIGHT corner for mirror center: ")
            bottomRightOptions.UseBasePoint = True
            bottomRightOptions.BasePoint = topLeftResult.Value

            Dim bottomRightResult As PromptPointResult = ed.GetPoint(bottomRightOptions)
            If bottomRightResult.Status <> PromptStatus.OK Then Return False

            center = New Point3d((topLeftResult.Value.X + bottomRightResult.Value.X) / 2.0,
                                 (topLeftResult.Value.Y + bottomRightResult.Value.Y) / 2.0,
                                 (topLeftResult.Value.Z + bottomRightResult.Value.Z) / 2.0)
            Return True
        End Function

        ''' <summary>
        ''' Vertical flip mirrors across a vertical line (flips left/right); otherwise a
        ''' horizontal line (flips up/down). Both pass through the supplied center point.
        ''' </summary>
        Private Shared Function BuildMirrorAxis(center As Point3d, verticalFlip As Boolean) As Line3d
            If verticalFlip Then
                Return New Line3d(center, center + New Vector3d(0.0, 200.0, 0.0))
            End If
            Return New Line3d(center, center + New Vector3d(200.0, 0.0, 0.0))
        End Function

        Private Shared Function NearestRightAngleDegrees(rotationRadians As Double) As Integer
            Dim angleDegrees As Double = rotationRadians * 180.0 / Math.PI
            Dim n As Integer = CInt(Math.Floor(angleDegrees / 90.0 + 0.5))
            Dim snapped As Integer = n * 90
            Return ((snapped Mod 360) + 360) Mod 360
        End Function

        Private Shared Sub FixInsertScales(blockRef As BlockReference)
            Dim nearestAngle As Integer = NearestRightAngleDegrees(blockRef.Rotation)
            Dim blockName As String = blockRef.Name
            Dim useXAngles As Integer() = Nothing

            If Array.IndexOf(UseXScaleAt0And270, blockName) >= 0 Then
                useXAngles = {0, 270}
            ElseIf Array.IndexOf(UseXScaleAt0And90, blockName) >= 0 Then
                useXAngles = {0, 90}
            End If

            If useXAngles Is Nothing Then Return

            Dim scales As Scale3d = blockRef.ScaleFactors
            If Array.IndexOf(useXAngles, nearestAngle) >= 0 Then
                blockRef.ScaleFactors = New Scale3d(1.0, scales.Y, scales.Z)
            Else
                blockRef.ScaleFactors = New Scale3d(scales.X, -1.0, scales.Z)
            End If
        End Sub

        ''' <summary>
        ''' Emulates MIRRTEXT = 0 for standalone text: mirror the anchor location, but
        ''' preserve the original text orientation instead of geometrically reversing it.
        ''' </summary>
        Private Shared Function TransformTextPositionPreservingOrientation(db As Database,
                                                                           source As Entity,
                                                                           mirrored As Entity,
                                                                           transform As Matrix3d,
                                                                           verticalFlip As Boolean) As Boolean
            Dim srcText As DBText = TryCast(source, DBText)
            If srcText IsNot Nothing Then
                Dim copy As DBText = CType(mirrored, DBText)
                copy.Position = srcText.Position.TransformBy(transform)
                If srcText.HorizontalMode <> TextHorizontalMode.TextLeft OrElse srcText.VerticalMode <> TextVerticalMode.TextBase Then
                    copy.AlignmentPoint = srcText.AlignmentPoint.TransformBy(transform)
                End If
                copy.Rotation = srcText.Rotation
                copy.IsMirroredInX = srcText.IsMirroredInX
                copy.IsMirroredInY = srcText.IsMirroredInY
                copy.WidthFactor = srcText.WidthFactor
                copy.Oblique = srcText.Oblique
                copy.AdjustAlignment(db)
                Return True
            End If

            Dim srcMText As MText = TryCast(source, MText)
            If srcMText IsNot Nothing Then
                Dim copy As MText = CType(mirrored, MText)
                Dim targetCenter As Point3d
                copy.Location = srcMText.Location.TransformBy(transform)
                SetMTextRotation(copy, MirrorTextRotation(srcMText.Rotation, verticalFlip))

                If TryGetTransformedExtentsCenter(srcMText, transform, targetCenter) Then
                    Dim copyCenter As Point3d
                    If TryGetExtentsCenter(copy, copyCenter) Then
                        copy.TransformBy(Matrix3d.Displacement(copyCenter.GetVectorTo(targetCenter)))
                    End If
                End If

                Return True
            End If

            Return False
        End Function

        Private Shared Function MirrorTextRotation(rotation As Double, verticalFlip As Boolean) As Double
            Dim mirrored As Double
            If verticalFlip Then
                mirrored = Math.PI - rotation
            Else
                mirrored = -rotation
            End If

            Return MostReadableRotation(mirrored)
        End Function

        Private Shared Function MostReadableRotation(rotation As Double) As Double
            Dim normalized As Double = NormalizeRadians(rotation)
            Dim degrees As Double = normalized * 180.0 / Math.PI
            Const tolerance As Double = 0.0001

            If degrees > 90.0 + tolerance AndAlso degrees <= 270.0 + tolerance Then
                normalized += Math.PI
            End If
            Return NormalizeRadians(normalized)
        End Function

        Private Shared Sub SetMTextRotation(mtext As MText, rotation As Double)
            mtext.Rotation = rotation
            mtext.Direction = New Vector3d(Math.Cos(rotation), Math.Sin(rotation), 0.0)
        End Sub

        ''' <summary>
        ''' Dimensions and multileaders need their leader/dimension geometry mirrored, but
        ''' their annotation text should be forced back into the readable half-plane.
        ''' </summary>
        Private Shared Sub CorrectAnnotationTextOrientation(ent As Entity)
            If TypeOf ent Is Dimension Then
                If TryMakeNumericRotationPropertyReadable(ent, "TextRotation") Then
                    RecomputeDimensionBlockIfAvailable(ent)
                    ent.RecordGraphicsModified(True)
                End If
                Return
            End If

            If ent.GetType().Name.IndexOf("MLeader", StringComparison.OrdinalIgnoreCase) < 0 Then Return

            Dim updated As Boolean = TryMakeNumericRotationPropertyReadable(ent, "TextRotation")
            updated = TryMakeMLeaderMTextReadable(ent) OrElse updated

            If updated Then ent.RecordGraphicsModified(True)
        End Sub

        Private Shared Function TryMakeNumericRotationPropertyReadable(target As Object, propertyName As String) As Boolean
            Dim prop As Reflection.PropertyInfo = target.GetType().GetProperty(propertyName)
            If prop Is Nothing OrElse Not prop.CanRead OrElse Not prop.CanWrite Then Return False
            If prop.PropertyType IsNot GetType(Double) Then Return False

            Dim current As Double
            Try
                current = CDbl(prop.GetValue(target, Nothing))
            Catch
                Return False
            End Try

            Dim readable As Double = MostReadableRotation(current)
            If Math.Abs(NormalizeRadians(readable - current)) < 0.0000001 Then Return False

            Try
                prop.SetValue(target, readable, Nothing)
            Catch
                Return False
            End Try

            Return True
        End Function

        Private Shared Function TryMakeMLeaderMTextReadable(target As Object) As Boolean
            Dim prop As Reflection.PropertyInfo = target.GetType().GetProperty("MText")
            If prop Is Nothing OrElse Not prop.CanRead Then Return False

            Dim mt As MText
            Try
                mt = TryCast(prop.GetValue(target, Nothing), MText)
            Catch
                Return False
            End Try

            If mt Is Nothing Then Return False

            Dim readable As Double = MostReadableRotation(mt.Rotation)
            If Math.Abs(NormalizeRadians(readable - mt.Rotation)) < 0.0000001 Then Return False

            SetMTextRotation(mt, readable)
            If prop.CanWrite Then
                Try
                    prop.SetValue(target, mt, Nothing)
                Catch
                    Return False
                End Try
            End If

            Return True
        End Function

        Private Shared Sub RecomputeDimensionBlockIfAvailable(target As Object)
            Dim method As Reflection.MethodInfo = target.GetType().GetMethod("RecomputeDimensionBlock", {GetType(Boolean)})
            If method Is Nothing Then Return
            Try
                method.Invoke(target, {True})
            Catch
            End Try
        End Sub

        Private Shared Function NormalizeRadians(angle As Double) As Double
            Dim twoPi As Double = 2.0 * Math.PI
            Dim normalized As Double = angle Mod twoPi
            If normalized < 0.0 Then normalized += twoPi
            Return normalized
        End Function

        Private Shared Function TryGetTransformedExtentsCenter(ent As Entity,
                                                              transform As Matrix3d,
                                                              ByRef center As Point3d) As Boolean
            Dim ext As Extents3d
            Try
                ext = ent.GeometricExtents
            Catch
                Return False
            End Try

            Dim minPt As Point3d = Point3d.Origin
            Dim maxPt As Point3d = Point3d.Origin
            Dim havePoint As Boolean = False

            For Each pt As Point3d In ExtentsCorners(ext)
                Dim transformed As Point3d = pt.TransformBy(transform)
                If Not havePoint Then
                    minPt = transformed
                    maxPt = transformed
                    havePoint = True
                Else
                    minPt = New Point3d(Math.Min(minPt.X, transformed.X),
                                        Math.Min(minPt.Y, transformed.Y),
                                        Math.Min(minPt.Z, transformed.Z))
                    maxPt = New Point3d(Math.Max(maxPt.X, transformed.X),
                                        Math.Max(maxPt.Y, transformed.Y),
                                        Math.Max(maxPt.Z, transformed.Z))
                End If
            Next

            If Not havePoint Then Return False

            center = New Point3d((minPt.X + maxPt.X) / 2.0,
                                 (minPt.Y + maxPt.Y) / 2.0,
                                 (minPt.Z + maxPt.Z) / 2.0)
            Return True
        End Function

        Private Shared Function TryGetExtentsCenter(ent As Entity, ByRef center As Point3d) As Boolean
            Dim ext As Extents3d
            Try
                ext = ent.GeometricExtents
            Catch
                Return False
            End Try

            center = New Point3d((ext.MinPoint.X + ext.MaxPoint.X) / 2.0,
                                 (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0,
                                 (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0)
            Return True
        End Function

        Private Shared Function ExtentsCorners(ext As Extents3d) As Point3d()
            Return {
                New Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MinPoint.Z),
                New Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MaxPoint.Z),
                New Point3d(ext.MinPoint.X, ext.MaxPoint.Y, ext.MinPoint.Z),
                New Point3d(ext.MinPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z),
                New Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z),
                New Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MaxPoint.Z),
                New Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, ext.MinPoint.Z),
                New Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z)
            }
        End Function

    End Class

End Namespace
