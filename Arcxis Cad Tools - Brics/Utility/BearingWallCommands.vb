Imports System
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.BearingWallCommands))>
Namespace Arcxis_Cad_Tools
    Public Class BearingWallCommands
        Private Const BearingLayerName As String = "S-FRM-BEARING"
        Private Const HatchPatternName As String = "ANSI31"
        Private Const HatchScaleValue As Double = 36.0

        <CommandMethod("BW", CommandFlags.Modal)>
        Public Sub DrawBearingWall()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then
                Return
            End If

            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim originalOsmode As Integer = CInt(Application.GetSystemVariable("OSMODE"))
            Dim originalCmdecho As Integer = CInt(Application.GetSystemVariable("CMDECHO"))
            Dim originalCmddia As Integer = CInt(Application.GetSystemVariable("CMDDIA"))
            Dim originalAttdia As Integer = CInt(Application.GetSystemVariable("ATTDIA"))
            Dim originalOrthomode As Integer = CInt(Application.GetSystemVariable("ORTHOMODE"))
            Dim originalLayer As String = CStr(Application.GetSystemVariable("CLAYER"))

            Using docLock As DocumentLock = doc.LockDocument()
                Try
                    Application.SetSystemVariable("CMDECHO", 0)
                    Application.SetSystemVariable("CMDDIA", 0)
                    Application.SetSystemVariable("ATTDIA", 0)
                    Application.SetSystemVariable("ORTHOMODE", 0)
                    Application.SetSystemVariable("OSMODE", 1)

                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        EnsureLayerExistsAndOn(tr, db, BearingLayerName)
                        tr.Commit()
                    End Using

                    Application.SetSystemVariable("CLAYER", BearingLayerName)

                    Dim startPointResult As PromptPointResult = PromptForPoint(ed, vbLf & "Pick START of bearing wall")
                    If startPointResult.Status <> PromptStatus.OK Then
                        Return
                    End If

                    Application.SetSystemVariable("OSMODE", 161)
                    Dim endPointResult As PromptPointResult = PromptForPoint(ed, vbLf & "Pick END of bearing wall", startPointResult.Value)
                    If endPointResult.Status <> PromptStatus.OK Then
                        Return
                    End If

                    Dim startPoint As Point3d = startPointResult.Value
                    Dim endPoint As Point3d = endPointResult.Value
                    Dim wallVector As Vector3d = endPoint - startPoint
                    If wallVector.Length <= Tolerance.Global.EqualPoint Then
                        ed.WriteMessage(vbLf & "Bearing wall start and end points must be different.")
                        Return
                    End If

                    Application.SetSystemVariable("OSMODE", 512)
                    Dim firstSideResult As PromptPointResult = PromptForPoint(ed, vbLf & "Pick on one side of wall")
                    If firstSideResult.Status <> PromptStatus.OK Then
                        Return
                    End If

                    Dim secondSideResult As PromptPointResult = PromptForPoint(ed, vbLf & "Pick other side of wall")
                    If secondSideResult.Status <> PromptStatus.OK Then
                        Return
                    End If

                    Dim unitDirection As Vector3d = wallVector.GetNormal()
                    Dim unitNormal As New Vector3d(-unitDirection.Y, unitDirection.X, 0.0)
                    Dim firstOffset As Double = SignedOffsetFromBaseline(startPoint, firstSideResult.Value, unitNormal)
                    Dim secondOffset As Double = SignedOffsetFromBaseline(startPoint, secondSideResult.Value, unitNormal)
                    Dim wallThickness As Double = Math.Abs(firstOffset - secondOffset)

                    If wallThickness <= Tolerance.Global.EqualPoint Then
                        ed.WriteMessage(vbLf & "Wall side picks must define a non-zero thickness.")
                        Return
                    End If

                    Dim firstStart As Point3d = OffsetPoint(startPoint, unitNormal, firstOffset)
                    Dim firstEnd As Point3d = OffsetPoint(endPoint, unitNormal, firstOffset)
                    Dim secondEnd As Point3d = OffsetPoint(endPoint, unitNormal, secondOffset)
                    Dim secondStart As Point3d = OffsetPoint(startPoint, unitNormal, secondOffset)

                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        Dim blockTable As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim modelSpace As BlockTableRecord = CType(tr.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                        Dim boundary As New Polyline()
                        boundary.SetDatabaseDefaults()
                        boundary.Layer = BearingLayerName
                        boundary.AddVertexAt(0, New Point2d(firstStart.X, firstStart.Y), 0.0, 0.0, 0.0)
                        boundary.AddVertexAt(1, New Point2d(firstEnd.X, firstEnd.Y), 0.0, 0.0, 0.0)
                        boundary.AddVertexAt(2, New Point2d(secondEnd.X, secondEnd.Y), 0.0, 0.0, 0.0)
                        boundary.AddVertexAt(3, New Point2d(secondStart.X, secondStart.Y), 0.0, 0.0, 0.0)
                        boundary.Closed = True
                        modelSpace.AppendEntity(boundary)
                        tr.AddNewlyCreatedDBObject(boundary, True)

                        Dim hatch As New Hatch()
                        hatch.SetDatabaseDefaults()
                        hatch.Layer = BearingLayerName
                        hatch.SetHatchPattern(HatchPatternType.PreDefined, HatchPatternName)
                        hatch.PatternScale = HatchScaleValue
                        hatch.PatternAngle = Math.Atan2(unitDirection.Y, unitDirection.X)
                        hatch.Associative = True
                        modelSpace.AppendEntity(hatch)
                        tr.AddNewlyCreatedDBObject(hatch, True)

                        Dim loopIds As New ObjectIdCollection()
                        loopIds.Add(boundary.ObjectId)
                        hatch.AppendLoop(HatchLoopTypes.External, loopIds)
                        hatch.EvaluateHatch(True)

                        tr.Commit()
                    End Using
                Finally
                    Application.SetSystemVariable("CLAYER", originalLayer)
                    Application.SetSystemVariable("ORTHOMODE", originalOrthomode)
                    Application.SetSystemVariable("ATTDIA", originalAttdia)
                    Application.SetSystemVariable("CMDDIA", originalCmddia)
                    Application.SetSystemVariable("CMDECHO", originalCmdecho)
                    Application.SetSystemVariable("OSMODE", originalOsmode)
                End Try
            End Using
        End Sub

        Private Shared Function PromptForPoint(ed As Editor, message As String, Optional basePoint As Point3d? = Nothing) As PromptPointResult
            Dim options As New PromptPointOptions(message)
            If basePoint.HasValue Then
                options.UseBasePoint = True
                options.BasePoint = basePoint.Value
            End If

            Return ed.GetPoint(options)
        End Function

        Private Shared Function SignedOffsetFromBaseline(basePoint As Point3d, pickedPoint As Point3d, unitNormal As Vector3d) As Double
            Dim pickVector As Vector3d = pickedPoint - basePoint
            Return pickVector.DotProduct(unitNormal)
        End Function

        Private Shared Function OffsetPoint(basePoint As Point3d, direction As Vector3d, distance As Double) As Point3d
            Return New Point3d(
                basePoint.X + (direction.X * distance),
                basePoint.Y + (direction.Y * distance),
                basePoint.Z + (direction.Z * distance))
        End Function

        Private Shared Sub EnsureLayerExistsAndOn(tr As Transaction, db As Database, layerName As String)
            Dim layerTable As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

            If Not layerTable.Has(layerName) Then
                layerTable.UpgradeOpen()
                Dim layer As New LayerTableRecord()
                layer.Name = layerName
                layerTable.Add(layer)
                tr.AddNewlyCreatedDBObject(layer, True)
            End If

            Dim layerId As ObjectId = layerTable(layerName)
            Dim layerRecord As LayerTableRecord = CType(tr.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)
            If layerRecord.IsOff Then
                layerRecord.IsOff = False
            End If

            If layerRecord.IsFrozen Then
                layerRecord.IsFrozen = False
            End If
        End Sub
    End Class
End Namespace