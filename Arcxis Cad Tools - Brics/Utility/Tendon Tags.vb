Imports System
Imports Bricscad.Runtime
Imports Teigha.Runtime
Imports Bricscad.ApplicationServices
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Bricscad.EditorInput
Imports Teigha.Colors

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.TendonTagging))>
Namespace Arcxis_Cad_Tools
    Public Class TendonTagging

        <CommandMethod("DST")>
        Public Sub DimensionTendons()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = Bricscad.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            'Dim ltsc As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("ltscale")
            Dim acLine As Line
            '' Set system variable to new value
            'Bricscad.ApplicationServices.Application.SetSystemVariable("FIELDDISPLAY", 0)

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    '' Get the current value from a system variable
                    Dim ech As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("cmdecho")
                    Dim otm As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("orthomode")
                    Dim oldos As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("osmode")
                    '' Set system variable to new value
                    Bricscad.ApplicationServices.Application.SetSystemVariable("orthomode", 1)
                    Bricscad.ApplicationServices.Application.SetSystemVariable("osmode", 16384)


                    Dim pts As New Point3dCollection()

                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

                    '' Prompt for the start point
                    pPtOpts.Message = vbLf & "Pick the FIRST point of the fenceline: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptStart As Point3d = pPtRes.Value
                    pts.Add(ptStart)

                    '' Exit if the user presses ESC or cancels the command
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                    '' Prompt for the end point
                    pPtOpts.Message = vbLf & "Pick the LAST point of the fenceline: "
                    pPtOpts.UseBasePoint = True
                    pPtOpts.BasePoint = ptStart
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptEnd As Point3d = pPtRes.Value
                    pts.Add(ptEnd)


                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                    Dim distRes = aced.GetPoint(vbLf & "Enter dimension offset distance: ")
                    Dim offsetpoint As Point3d = distRes.Value
                    Dim offsetdist As Double

                    If distRes.Status <> PromptStatus.OK Then Exit Sub

                    Dim acBlkTbl As BlockTable = DirectCast(acTrans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim acBlkTblRec As BlockTableRecord = DirectCast(acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)


                    '''''creates the line which user input to find slabline constraints
                    acLine = New Line(ptStart, ptEnd)
                    acLine.Layer = "0"
                    acLine.Linetype = "Hidden"
                    Dim FLineAng As Double = acLine.Angle
                    Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI

                    If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then

                        offsetdist = pts(1).Y - offsetpoint.Y

                    Else

                        offsetdist = pts(1).X - offsetpoint.X

                    End If

                    Dim crossingPts As New Point3dCollection()

                    Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE,LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "S-FND-STEND,S-FND-SLABDP")}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Dim firstPt As New Point2d

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each SSSObj As ObjectId In SS.GetObjectIds()

                                Dim ent As Entity = TryCast(acTrans.GetObject(SSSObj, OpenMode.ForRead), Entity)
                                If ent Is Nothing OrElse Not IsLineOrPolyline(ent) Then Continue For

                                AddIntersectionsWithLine(ent, acLine, crossingPts)

                            Next

                            For i = 0 To crossingPts.Count - 2
                                Dim pt1 = crossingPts(i)
                                Dim pt2 = crossingPts(i + 1)

                                Dim midPt As New Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0)
                                Dim perpVec As Vector3d = acLine.Delta.GetPerpendicularVector().GetNormal() * offsetdist
                                Dim dimLinePt As Point3d = midPt + perpVec
                                Dim dimLine As New RotatedDimension() With {
                                .XLine1Point = pt1,
                                .XLine2Point = pt2,
                                .DimLinePoint = dimLinePt,
                                .DimensionStyle = acCurDb.Dimstyle,
                                .Layer = "S-FND-DIM"
                            }

                                acBlkTblRec.AppendEntity(dimLine)
                                acTrans.AddNewlyCreatedDBObject(dimLine, True)
                            Next

                        End If

                    End If

                    Bricscad.ApplicationServices.Application.SetSystemVariable("orthomode", otm)
                    Bricscad.ApplicationServices.Application.SetSystemVariable("osmode", oldos)

                    '' Save the changes and dispose of the transaction
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        <CommandMethod("WTD")>
        Public Sub TendonDraw()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = Bricscad.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            'Dim ltsc As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("ltscale")
            Dim acLine As Line
            Dim TendonPointsList As New List(Of Point3d)
            Dim NewTendonSpacing As Double

            Dim spacing As Double

            Dim tendoninfo2 As New PromptStringOptions(vbLf & "Specify Tendon Spacing(in inches)...")

            Dim results1 As PromptResult = Bricscad.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetString(tendoninfo2)

            Dim result1result As Double = results1.StringResult

            spacing = result1result


            '' Set system variable to new value
            'Bricscad.ApplicationServices.Application.SetSystemVariable("FIELDDISPLAY", 0)

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    '' Get the current value from a system variable
                    Dim ech As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("cmdecho")
                    Dim otm As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("orthomode")
                    Dim oldos As Integer = Bricscad.ApplicationServices.Application.GetSystemVariable("osmode")
                    '' Set system variable to new value
                    Bricscad.ApplicationServices.Application.SetSystemVariable("orthomode", 1)
                    Bricscad.ApplicationServices.Application.SetSystemVariable("osmode", 16384)


                    Dim pts As New Point3dCollection()

                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

                    '' Prompt for the start point
                    pPtOpts.Message = vbLf & "Pick the FIRST point of the fenceline: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptStart As Point3d = pPtRes.Value
                    pts.Add(ptStart)

                    '' Exit if the user presses ESC or cancels the command
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                    '' Prompt for the end point
                    pPtOpts.Message = vbLf & "Pick the LAST point of the fenceline: "
                    pPtOpts.UseBasePoint = True
                    pPtOpts.BasePoint = ptStart
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptEnd As Point3d = pPtRes.Value
                    pts.Add(ptEnd)

                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub


                    Dim acBlkTbl As BlockTable = DirectCast(acTrans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim acBlkTblRec As BlockTableRecord = DirectCast(acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)


                    '''''creates the line which user input to find slabline constraints
                    acLine = New Line(ptStart, ptEnd)
                    acLine.Layer = "0"
                    acLine.Linetype = "Hidden"
                    Dim FLineAng As Double = acLine.Angle
                    Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI


                    Dim SlabIptCol As New Point3dCollection()
                    Dim SlabLineIptCol As New Point3dCollection()
                    Dim IntWithTendon As New Point3dCollection()

                    Dim acTypValAr As TypedValue() = New TypedValue() {
                        New TypedValue(0, "LINE,LWPOLYLINE"),
                        New TypedValue(DxfCode.LayerName, "S-FND-SLAB,S-FND-STEND")
                    }
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Dim SlabEnt As Entity = Nothing

                    Dim CrossingPoint1 As New Point3d()
                    Dim Crossingpoint2 As New Point3d()

                    Const DistanceFromSlab As Double = 6

                    Dim LengthOfLine As Double
                    Dim NumberofTendons As Double
                    Dim StartPoint As New Point3d
                    Dim EndPoint As New Point3d
                    Dim StartWithTendon As Boolean = False
                    Dim EndWithTendon As Boolean = False

                    Dim minX As Double
                    Dim maxX As Double
                    Dim minY As Double
                    Dim maxY As Double

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each SSSObj As ObjectId In SS.GetObjectIds()

                                Dim ent As Entity = TryCast(acTrans.GetObject(SSSObj, OpenMode.ForRead), Entity)
                                If ent Is Nothing OrElse Not IsLineOrPolyline(ent) Then Continue For

                                If String.Equals(ent.Layer, "S-FND-SLAB", StringComparison.OrdinalIgnoreCase) Then
                                    SlabEnt = ent
                                    AddIntersectionsWithLine(ent, acLine, SlabIptCol)
                                    SetSlabExtents(ent, minX, maxX, minY, maxY)

                                ElseIf String.Equals(ent.Layer, "S-FND-STEND", StringComparison.OrdinalIgnoreCase) Then
                                    AddIntersectionsWithLine(ent, acLine, IntWithTendon)
                                End If

                            Next

                            ' tendon to tendon, tendon to slab, or slab to slab
                            If SlabIptCol.Count < 2 Then
                                If SlabIptCol.Count < 1 Then
                                    Dim peo As New PromptEntityOptions(vbLf & "Select outerslabline:")
                                    Dim per As PromptEntityResult = aced.GetEntity(peo)
                                    If per.Status <> PromptStatus.OK Then Exit Sub

                                    SlabEnt = TryCast(acTrans.GetObject(per.ObjectId, OpenMode.ForRead), Entity)
                                    If SlabEnt Is Nothing Then Exit Sub

                                    AddIntersectionsWithLine(SlabEnt, acLine, SlabIptCol)
                                    SetSlabExtents(SlabEnt, minX, maxX, minY, maxY)

                                    If IntWithTendon.Count < 1 Then
                                        acDoc.Editor.WriteMessage(vbLf & "FENCE MUST CROSS A TENDON!!")
                                        Exit Sub
                                    ElseIf IntWithTendon.Count = 1 AndAlso SlabIptCol.Count >= 1 Then
                                        SetTendonToSlabBounds(SlabIptCol(0), IntWithTendon(0), FLineAngD, DistanceFromSlab,
                                                              StartPoint, EndPoint, StartWithTendon, EndWithTendon,
                                                              LengthOfLine, NumberofTendons, NewTendonSpacing, spacing)
                                    ElseIf IntWithTendon.Count < 2 Then
                                        acDoc.Editor.WriteMessage(vbLf & "FENCE MUST CROSS TWO TENDONS!!")
                                        Exit Sub
                                    Else
                                        If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then
                                            If IntWithTendon(0).X < IntWithTendon(1).X Then
                                                CrossingPoint1 = IntWithTendon(0)
                                                Crossingpoint2 = IntWithTendon(1)
                                            Else
                                                CrossingPoint1 = IntWithTendon(1)
                                                Crossingpoint2 = IntWithTendon(0)
                                            End If

                                            StartPoint = New Point3d(CrossingPoint1.X, CrossingPoint1.Y, 0)
                                            EndPoint = New Point3d(Crossingpoint2.X, Crossingpoint2.Y, 0)
                                            LengthOfLine = CrossingPoint1.X - Crossingpoint2.X
                                            NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                            NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                        Else
                                            If IntWithTendon(0).Y > IntWithTendon(1).Y Then
                                                CrossingPoint1 = IntWithTendon(0)
                                                Crossingpoint2 = IntWithTendon(1)
                                            Else
                                                CrossingPoint1 = IntWithTendon(1)
                                                Crossingpoint2 = IntWithTendon(0)
                                            End If

                                            StartPoint = New Point3d(CrossingPoint1.X, CrossingPoint1.Y, 0)
                                            EndPoint = New Point3d(Crossingpoint2.X, Crossingpoint2.Y, 0)
                                            LengthOfLine = Math.Abs(CrossingPoint1.Y) - Math.Abs(Crossingpoint2.Y)
                                            NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                            NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                        End If

                                        StartWithTendon = True
                                        EndWithTendon = True
                                    End If

                                Else
                                    If IntWithTendon.Count < 1 Then
                                        acDoc.Editor.WriteMessage(vbLf & "FENCE MUST CROSS A TENDON!!")
                                        Exit Sub
                                    End If

                                    Dim boundaryTendonInt As Point3d = IntWithTendon(0)
                                    If IntWithTendon.Count > 1 Then
                                        Dim minDistance As Double = SlabIptCol(0).DistanceTo(boundaryTendonInt)
                                        For idx As Integer = 1 To IntWithTendon.Count - 1
                                            Dim currentDistance As Double = SlabIptCol(0).DistanceTo(IntWithTendon(idx))
                                            If currentDistance < minDistance Then
                                                minDistance = currentDistance
                                                boundaryTendonInt = IntWithTendon(idx)
                                            End If
                                        Next
                                    End If

                                    SetTendonToSlabBounds(SlabIptCol(0), boundaryTendonInt, FLineAngD, DistanceFromSlab,
                                                          StartPoint, EndPoint, StartWithTendon, EndWithTendon,
                                                          LengthOfLine, NumberofTendons, NewTendonSpacing, spacing)
                                End If

                            ElseIf SlabIptCol.Count = 2 Then

                                Dim useMixedSlabTendonBounds As Boolean = False
                                Dim slabBoundary As Point3d = SlabIptCol(0)
                                Dim tendonBoundary As Point3d = New Point3d()

                                If IntWithTendon.Count > 0 Then
                                    Dim nearestSlabToStart As Point3d = SlabIptCol(0)
                                    Dim nearestSlabToEnd As Point3d = SlabIptCol(0)
                                    Dim nearestTendonToStart As Point3d = IntWithTendon(0)
                                    Dim nearestTendonToEnd As Point3d = IntWithTendon(0)

                                    Dim minSlabStartDist As Double = ptStart.DistanceTo(nearestSlabToStart)
                                    Dim minSlabEndDist As Double = ptEnd.DistanceTo(nearestSlabToEnd)
                                    For sIdx As Integer = 1 To SlabIptCol.Count - 1
                                        Dim slabStartDist As Double = ptStart.DistanceTo(SlabIptCol(sIdx))
                                        If slabStartDist < minSlabStartDist Then
                                            minSlabStartDist = slabStartDist
                                            nearestSlabToStart = SlabIptCol(sIdx)
                                        End If

                                        Dim slabEndDist As Double = ptEnd.DistanceTo(SlabIptCol(sIdx))
                                        If slabEndDist < minSlabEndDist Then
                                            minSlabEndDist = slabEndDist
                                            nearestSlabToEnd = SlabIptCol(sIdx)
                                        End If
                                    Next

                                    Dim minTendonStartDist As Double = ptStart.DistanceTo(nearestTendonToStart)
                                    Dim minTendonEndDist As Double = ptEnd.DistanceTo(nearestTendonToEnd)
                                    For tIdx As Integer = 1 To IntWithTendon.Count - 1
                                        Dim tendonStartDist As Double = ptStart.DistanceTo(IntWithTendon(tIdx))
                                        If tendonStartDist < minTendonStartDist Then
                                            minTendonStartDist = tendonStartDist
                                            nearestTendonToStart = IntWithTendon(tIdx)
                                        End If

                                        Dim tendonEndDist As Double = ptEnd.DistanceTo(IntWithTendon(tIdx))
                                        If tendonEndDist < minTendonEndDist Then
                                            minTendonEndDist = tendonEndDist
                                            nearestTendonToEnd = IntWithTendon(tIdx)
                                        End If
                                    Next

                                    Dim startPrefersSlab As Boolean = (minSlabStartDist <= minTendonStartDist)
                                    Dim endPrefersSlab As Boolean = (minSlabEndDist <= minTendonEndDist)

                                    If startPrefersSlab <> endPrefersSlab Then
                                        useMixedSlabTendonBounds = True
                                        If startPrefersSlab Then
                                            slabBoundary = nearestSlabToStart
                                            tendonBoundary = nearestTendonToEnd
                                        Else
                                            slabBoundary = nearestSlabToEnd
                                            tendonBoundary = nearestTendonToStart
                                        End If
                                    End If
                                End If

                                If useMixedSlabTendonBounds Then
                                    If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then
                                        Dim slabOffsetPoint As Point3d
                                        If slabBoundary.X < tendonBoundary.X Then
                                            slabOffsetPoint = New Point3d(slabBoundary.X + DistanceFromSlab, slabBoundary.Y, 0)
                                            StartPoint = slabOffsetPoint
                                            EndPoint = New Point3d(tendonBoundary.X, tendonBoundary.Y, 0)
                                            EndWithTendon = True
                                        Else
                                            slabOffsetPoint = New Point3d(slabBoundary.X - DistanceFromSlab, slabBoundary.Y, 0)
                                            StartPoint = New Point3d(tendonBoundary.X, tendonBoundary.Y, 0)
                                            EndPoint = slabOffsetPoint
                                            StartWithTendon = True
                                        End If

                                        LengthOfLine = Math.Abs(StartPoint.X - EndPoint.X)
                                        NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                        NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                    Else
                                        Dim slabOffsetPoint As Point3d
                                        If slabBoundary.Y > tendonBoundary.Y Then
                                            slabOffsetPoint = New Point3d(slabBoundary.X, slabBoundary.Y - DistanceFromSlab, 0)
                                            StartPoint = slabOffsetPoint
                                            EndPoint = New Point3d(tendonBoundary.X, tendonBoundary.Y, 0)
                                            EndWithTendon = True
                                        Else
                                            slabOffsetPoint = New Point3d(slabBoundary.X, slabBoundary.Y + DistanceFromSlab, 0)
                                            StartPoint = New Point3d(tendonBoundary.X, tendonBoundary.Y, 0)
                                            EndPoint = slabOffsetPoint
                                            StartWithTendon = True
                                        End If

                                        LengthOfLine = Math.Abs(StartPoint.Y - EndPoint.Y)
                                        NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                        NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                    End If
                                Else
                                    If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then
                                        If SlabIptCol(0).X < SlabIptCol(1).X Then
                                            CrossingPoint1 = SlabIptCol(0)
                                            Crossingpoint2 = SlabIptCol(1)
                                        Else
                                            CrossingPoint1 = SlabIptCol(1)
                                            Crossingpoint2 = SlabIptCol(0)
                                        End If

                                        StartPoint = New Point3d(CrossingPoint1.X + DistanceFromSlab, CrossingPoint1.Y, 0)
                                        EndPoint = New Point3d(Crossingpoint2.X - DistanceFromSlab, Crossingpoint2.Y, 0)
                                        LengthOfLine = (CrossingPoint1.X + DistanceFromSlab) - (Crossingpoint2.X - DistanceFromSlab)
                                        NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                        NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                    Else
                                        If SlabIptCol(0).Y > SlabIptCol(1).Y Then
                                            CrossingPoint1 = SlabIptCol(0)
                                            Crossingpoint2 = SlabIptCol(1)
                                        Else
                                            CrossingPoint1 = SlabIptCol(1)
                                            Crossingpoint2 = SlabIptCol(0)
                                        End If

                                        StartPoint = New Point3d(CrossingPoint1.X, CrossingPoint1.Y - DistanceFromSlab, 0)
                                        EndPoint = New Point3d(Crossingpoint2.X, Crossingpoint2.Y + DistanceFromSlab, 0)
                                        LengthOfLine = (CrossingPoint1.Y - DistanceFromSlab) - (Crossingpoint2.Y + DistanceFromSlab)
                                        NumberofTendons = Math.Ceiling(Math.Abs(LengthOfLine) / spacing)
                                        NewTendonSpacing = Math.Abs(LengthOfLine / NumberofTendons)
                                    End If
                                End If
                            End If

                            If SlabEnt Is Nothing Then
                                acDoc.Editor.WriteMessage(vbLf & "No slab boundary found along fence line.")
                                Exit Sub
                            End If

                            Dim NewTendonPoint As New Point3d
                            TendonPointsList.Clear()

                            If StartWithTendon = False Then
                                TendonPointsList.Add(StartPoint)
                            End If

                            For i = 1 To NumberofTendons - 1
                                If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then
                                    NewTendonPoint = New Point3d(StartPoint.X + (i * NewTendonSpacing), StartPoint.Y, 0)
                                Else
                                    NewTendonPoint = New Point3d(StartPoint.X, StartPoint.Y - (i * NewTendonSpacing), 0)
                                End If
                                TendonPointsList.Add(NewTendonPoint)
                            Next

                            If EndWithTendon = False Then
                                TendonPointsList.Add(EndPoint)
                            End If

                            For Each point In TendonPointsList

                                If FLineAngD = 0 Or FLineAngD = 180 Or FLineAngD = 360 Then
                                    Dim TendoLine As Line
                                    Dim TendoStart As New Point3d
                                    Dim TendoEnd As New Point3d
                                    TendoStart = New Point3d(point.X, maxY, 0)
                                    TendoEnd = New Point3d(point.X, minY, 0)
                                    TendoLine = New Line(TendoStart, TendoEnd)

                                    SlabEnt.IntersectWith(TendoLine, Intersect.OnBothOperands, SlabLineIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If SlabLineIptCol.Count <= 2 Then

                                        TendoStart = New Point3d(point.X, SlabLineIptCol(0).Y, 0)
                                        TendoEnd = New Point3d(point.X, SlabLineIptCol(1).Y, 0)
                                        TendoLine = New Line(TendoStart, TendoEnd)

                                    ElseIf SlabLineIptCol.Count <= 4 Then
                                        Dim value As Double
                                        value = point.Y
                                        For h = 0 To SlabLineIptCol.Count - 2
                                            If Math.Abs(value) < Math.Max(Math.Abs(SlabLineIptCol(h).Y), Math.Abs(SlabLineIptCol(h + 1).Y)) And Math.Abs(value) > Math.Min(Math.Abs(SlabLineIptCol(h).Y), Math.Abs(SlabLineIptCol(h + 1).Y)) Then
                                                TendoStart = New Point3d(point.X, SlabLineIptCol(h).Y, 0)
                                                TendoEnd = New Point3d(point.X, SlabLineIptCol(h + 1).Y, 0)
                                                Exit For
                                            End If
                                        Next
                                        'TendoStart = New Point3d(point.X, SlabLineIptCol(0).Y, 0)
                                        'TendoEnd = New Point3d(point.X, SlabLineIptCol(1).Y, 0)
                                        TendoLine = New Line(TendoStart, TendoEnd)

                                    End If

                                    TendoLine.Layer = "S-FND-STEND"

                                    acBlkTblRec.AppendEntity(TendoLine)
                                    acTrans.AddNewlyCreatedDBObject(TendoLine, True)
                                    SlabLineIptCol.Clear()

                                Else

                                    Dim TendoLine As Line
                                    Dim TendoStart As New Point3d
                                    Dim TendoEnd As New Point3d
                                    TendoStart = New Point3d(maxX, point.Y, 0)
                                    TendoEnd = New Point3d(minX, point.Y, 0)
                                    TendoLine = New Line(TendoStart, TendoEnd)

                                    SlabEnt.IntersectWith(TendoLine, Intersect.OnBothOperands, SlabLineIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If SlabLineIptCol.Count <= 2 Then

                                        TendoStart = New Point3d(SlabLineIptCol(0).X, point.Y, 0)
                                        TendoEnd = New Point3d(SlabLineIptCol(1).X, point.Y, 0)
                                        TendoLine = New Line(TendoStart, TendoEnd)

                                    ElseIf SlabLineIptCol.Count <= 4 Then
                                        Dim value As Double
                                        value = point.X
                                        For h = 0 To SlabLineIptCol.Count - 2
                                            If Math.Abs(value) < Math.Max(Math.Abs(SlabLineIptCol(h).X), Math.Abs(SlabLineIptCol(h + 1).X)) And Math.Abs(value) > Math.Min(Math.Abs(SlabLineIptCol(h).X), Math.Abs(SlabLineIptCol(h + 1).X)) Then
                                                TendoStart = New Point3d(SlabLineIptCol(h).X, point.Y, 0)
                                                TendoEnd = New Point3d(SlabLineIptCol(h + 1).X, point.Y, 0)
                                                Exit For
                                            End If
                                        Next
                                        TendoLine = New Line(TendoStart, TendoEnd)

                                    End If

                                    TendoLine.Layer = "S-FND-STEND"

                                    acBlkTblRec.AppendEntity(TendoLine)
                                    acTrans.AddNewlyCreatedDBObject(TendoLine, True)
                                    SlabLineIptCol.Clear()

                                End If

                            Next

                            'acLine.StartPoint = StartPoint
                            'acLine.EndPoint = EndPoint
                            'acBlkTblRec.AppendEntity(acLine)
                            'acTrans.AddNewlyCreatedDBObject(acLine, True)

                        End If

                    End If

                    Bricscad.ApplicationServices.Application.SetSystemVariable("orthomode", otm)
                    Bricscad.ApplicationServices.Application.SetSystemVariable("osmode", oldos)

                    '' Save the changes and dispose of the transaction
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        <CommandMethod("WTT")>
        Public Sub WestTendonTag()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim CableList As New List(Of Entity)
            Dim SortedLineList As New List(Of Entity)
            Dim SortedCableStartXList As New List(Of Point3d)
            Dim SortedCableStartYList As New List(Of Point3d)
            Dim SortedCableStartXPoints As New List(Of Point3d)
            Dim SortedCableStartPts As New List(Of Point3d)
            Dim SortedCableStartYPoints As New List(Of Point3d)
            Dim CablePointsList As New List(Of Point3d)
            Dim SortedCablePointsList As New List(Of Point3d)
            Dim CableStartPointsList As New List(Of Point3d)
            Dim SortedCablePoints As New List(Of Point3d)
            Dim SortedCableXPointsList As New List(Of Point3d)
            Dim SortedCableStartPoints As New List(Of Double)
            Dim SortedCableYPointsList As New List(Of Point3d)
            Dim SortedCableYPoints As New List(Of Double)
            Dim CableTailIptList As New List(Of Point3d)
            Dim TailIptAng As New Dictionary(Of Point3d, Integer)
            Dim HeadIptAng As New Dictionary(Of Point3d, Integer)
            Dim CableStart As Point3d
            Dim CableEnd As Point3d
            Dim CableAng As Double
            Dim CableAngD As Integer
            Dim TailRot As Double
            Dim TailIpt As Point3d
            Dim HeadIpt As Point3d
            Dim HoriCableList As New List(Of Object)
            Dim HoriCablePointsList As New List(Of Point3d)
            Dim SortedHoriCablePointsList As New List(Of Point3d)
            Dim VertCableList As New List(Of Object)
            Dim VertCablePointsList As New List(Of Point3d)
            Dim SortedVertCablePointsList As New List(Of Point3d)
            Dim VertTendon2Num As Integer
            Dim VertTendon3Num As Integer

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim opts As New PromptSelectionOptions()
                opts.MessageForAdding = vbLf & "Select Foundation Plan: ..."

                Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE,LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "CABLE,S-FND-BTEND,S-FND-STEND")}
                Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                '' Request for objects to be selected in the drawing area
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(opts, acSelFtr)

                '' If the prompt status is OK, objects were selected
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value

                    '' Step through the objects in the selection set
                    For Each acSSObj As SelectedObject In acSSet

                        Dim typ As String = acSSObj.ObjectId.ObjectClass.Name()
                        Dim acEnt As Object = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead, False, True)


                        If typ = "AcDbLine" Then

                            CableStart = acEnt.StartPoint
                            CableEnd = acEnt.EndPoint
                            CableAng = acEnt.Angle
                            CableAngD = CableAng * 180.0 / Math.PI

                        Else typ = "AcDbPolyline"

                            CableStart = acEnt.StartPoint
                            CableEnd = acEnt.EndPoint
                            Dim lwp As Polyline = TryCast(acEnt, Polyline)
                            Dim lwpStartpt As Point3d = lwp.GetPoint3dAt(0)
                            Dim lwpEndpt As Point3d = lwp.GetPoint3dAt(1)
                            CableAng = Math.Atan2(lwpEndpt.Y - lwpStartpt.Y, lwpEndpt.X - lwpStartpt.X)
                            CableAngD = CableAng * 180.0 / Math.PI

                        End If

                        If CableAngD = 0 OrElse CableAngD = 360 Then

                            HoriCableList.Add(acEnt)
                            HoriCablePointsList.Add(CableStart)

                        ElseIf CableAngD = 90 OrElse CableAngD = -270 Then

                            VertCableList.Add(acEnt)
                            VertCablePointsList.Add(CableStart)

                        ElseIf CableAngD = 180 OrElse CableAngD = -180 Then

                            HoriCableList.Add(acEnt)
                            HoriCablePointsList.Add(CableEnd)

                        ElseIf CableAngD = 270 OrElse CableAngD = -90 Then

                            VertCableList.Add(acEnt)
                            VertCablePointsList.Add(CableEnd)

                        End If

                        CablePointsList.Add(CableStart)
                        CablePointsList.Add(CableEnd)

                    Next

                    SortedCableXPointsList = (From Xpnt In CablePointsList Order By Xpnt.X Select Xpnt).ToList

                    Dim CableXPT As Point3d = SortedCableXPointsList(0)
                    Dim CableX As Double = Math.Round(CableXPT.X, 10)

                    SortedCableYPointsList = (From Ypnt In CablePointsList Order By Ypnt.Y Select Ypnt).ToList

                    Dim CableYPT As Point3d = SortedCableYPointsList(SortedCableYPointsList.Count - 1)
                    Dim CableY As Double = Math.Round(CableYPT.Y, 10)
                    Dim HoriTendon2Num As Integer = 1
                    Dim HoriTendonLen As String = "0"
                    Dim HoriTendon3Num As Integer = 1
                    Dim CableLength As Double


                    SortedHoriCablePointsList = (From Ypnt In HoriCablePointsList Order By Ypnt.Y Select Ypnt).ToList

                    For Each pt As Point3d In SortedHoriCablePointsList

                        For Each CableEnt As Object In HoriCableList

                            CableStart = CableEnt.StartPoint
                            CableEnd = CableEnt.EndPoint
                            Dim ID As String = CableEnt.objectid.ToString
                            Dim CableColor As String = CableEnt.ColorIndex
                            Dim CableLayer As String = CableEnt.layer
                            Dim IDpart As String() = ID.Split("(")
                            Dim IDpart1 As String = IDpart(1)
                            Dim IDpart2 As String() = IDpart1.Split(")")
                            Dim CableID As String = IDpart2(0)
                            CableLength = (Math.Abs(CableStart.X - CableEnd.X) / 12)


                            If CableStart = pt Then

                                HoriTendonLen = "%<\AcExpr ((round(((%<\AcObjProp Object(%<\_ObjId " & CableID & ">%).Length \f ""%lu2%pr2"">%+18.12)/12)+.499))) \f ""%lu2%pr0%ps[(,)]"">%"
                                TailIpt = CableEnd
                                TailRot = 1.570796
                                HeadIpt = New Point3d(CableXPT.X, CableStart.Y, 0)


                            ElseIf CableEnd = pt Then

                                HoriTendonLen = "%<\AcExpr ((round(((%<\AcObjProp Object(%<\_ObjId " & CableID & ">%).Length \f ""%lu2%pr2"">%+18.12)/12)+.499))) \f ""%lu2%pr0%ps[(,)]"">%"
                                TailIpt = CableStart
                                TailRot = 1.570796
                                HeadIpt = New Point3d(CableXPT.X, CableEnd.Y, 0)

                            End If

                            If CableStart = pt OrElse CableEnd = pt Then

                                InsertTendonTail(acDoc, acCurDb, aced, acTrans, TailRot, TailIpt)

                                HoriTendonHeadInsertAdi(acDoc, acCurDb, aced, acTrans, HeadIpt, HoriTendon2Num, CableLength, "S-FND-DBL")
                                HoriTendon2Num = HoriTendon2Num + 1


                            End If

                        Next

                    Next

                    VertTendon2Num = HoriTendon2Num
                    Dim VertTendonLen As String = "0"
                    VertTendon3Num = HoriTendon3Num

                    SortedVertCablePointsList = (From Xpnt In VertCablePointsList Order By Xpnt.X Select Xpnt).ToList

                    For Each pt As Point3d In SortedVertCablePointsList

                        For Each CableEnt As Object In VertCableList

                            CableStart = CableEnt.StartPoint
                            CableEnd = CableEnt.EndPoint
                            Dim ID As String = CableEnt.objectid.ToString
                            Dim CableColor As String = CableEnt.ColorIndex
                            Dim CableLayer As String = CableEnt.layer
                            Dim IDpart As String() = ID.Split("(")
                            Dim IDpart1 As String = IDpart(1)
                            Dim IDpart2 As String() = IDpart1.Split(")")
                            Dim CableID As String = IDpart2(0)
                            CableLength = (Math.Abs(CableStart.Y - CableEnd.Y) / 12)

                            If CableStart = pt Then

                                VertTendonLen = "%<\AcExpr ((round(((%<\AcObjProp Object(%<\_ObjId " & CableID & ">%).Length \f ""%lu2%pr2"">%+18.12)/12)+.499))) \f ""%lu2%pr0%ps[(,)]"">%"
                                TailIpt = CableStart
                                TailRot = 0.0
                                HeadIpt = New Point3d(CableStart.X, CableYPT.Y, 0)

                            ElseIf CableEnd = pt Then

                                VertTendonLen = "%<\AcExpr ((round(((%<\AcObjProp Object(%<\_ObjId " & CableID & ">%).Length \f ""%lu2%pr2"">%+18.12)/12)+.499))) \f ""%lu2%pr0%ps[(,)]"">%"
                                TailIpt = CableEnd
                                TailRot = 0.0
                                HeadIpt = New Point3d(CableStart.X, CableYPT.Y, 0)

                            End If

                            If CableStart = pt OrElse CableEnd = pt Then

                                InsertTendonTail(acDoc, acCurDb, aced, acTrans, TailRot, TailIpt)

                                VertSlabTendonInsertAdi(acDoc, acCurDb, aced, acTrans, HeadIpt, VertTendon2Num, CableLength, "S-FND-DBL")
                                VertTendon2Num = VertTendon2Num + 1

                            End If

                        Next

                    Next

                End If

                '' Save the changes and dispose of the transaction
                acTrans.Commit()

            End Using

            acDoc.Editor.Regen()

        End Sub

        Private Sub HoriTendonHeadInsertAdi(acDoc As Document, acCurDb As Database, aced As Editor, acTrans As Transaction, HeadIpt As Point3d, HoriTendon2Num As Integer, HoriTendonLen As String, BlockLayer As String)

            Dim Dimscale As Double = 1
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim BlockName As String
            Dim BlockColor As Integer
            BlockName = "WestTendonHead"
            BlockColor = 3

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has(BlockLayer) = False Then
                Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                '' Assign the layer the ACI color 1 and a name
                acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, BlockColor)
                acLyrTblRec.Name = BlockLayer

                '' Upgrade the Layer table for write
                acLyrTbl.UpgradeOpen()

                '' Append the new layer to the Layer table and the transaction
                acLyrTbl.Add(acLyrTblRec)
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            ElseIf acLyrTbl.Has(BlockLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            End If

            ' to force update drawing screen
            acDoc.TransactionManager.EnableGraphicsFlush(True)
            Dim bt As BlockTable = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite), BlockTable)
            Dim tsTbl As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
            Dim textstylename As String = "TendonTagging"

            If tsTbl.Has(textstylename) Then

                Dim TxtStyleVar As Object
                'Get the TextStylye SYSTEM VARIABLE
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                'SET TextStyle SYSTEM VARIABLE
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "TendonTagging")

            Else

                Dim st As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                Dim strr As TextStyleTableRecord = New TextStyleTableRecord()

                strr.Name = "TendonTagging"
                st.Add(strr)
                strr.FileName = "romans.shx"
                strr.ObliquingAngle = 0.0
                strr.XScale = 0.8
                strr.TextSize = 0.0
                strr.IsVertical = False
                strr.IsShapeFile = False
                acTrans.AddNewlyCreatedDBObject(strr, True)

                'make as current
                Dim TxtStyleVar As Object
                'Get the TextStylye SYSTEM VARIABLE
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                'SET TextStyle SYSTEM VARIABLE
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "TendonTagging")

            End If

            ' if the block table doesn't already exists, exit
            Dim ed As Editor = acDoc.Editor
            Dim blktb As BlockTable = TryCast(acCurDb.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)

            ' Get the block definitions
            Dim blkDef1 As BlockTableRecord = TryCast(blktb(BlockName).GetObject(OpenMode.ForRead), BlockTableRecord)
            Dim blkDef2 As BlockTableRecord = TryCast(blktb("TendonNumber").GetObject(OpenMode.ForRead), BlockTableRecord)

            ' Open modelspace (or paperspace) for writing
            Dim BlkTbl As BlockTableRecord = DirectCast(acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            ' === Block 1: original block ===
            Using blkRef1 As New BlockReference(HeadIpt, blkDef1.ObjectId)
                blkRef1.Layer = BlockLayer
                blkRef1.ScaleFactors = New Scale3d(Dimscale, Dimscale, Dimscale)
                BlkTbl.AppendEntity(blkRef1)
                acTrans.AddNewlyCreatedDBObject(blkRef1, True)

                ' (optional) add attributes if any
            End Using

            ' === Block 2: "Tendon Number" block ===
            ' Let's say you want it a bit offset from the first one:

            Using blkRef2 As New BlockReference(HeadIpt, blkDef2.ObjectId)
                blkRef2.Layer = BlockLayer
                blkRef2.ScaleFactors = New Scale3d(Dimscale, Dimscale, Dimscale)
                BlkTbl.AppendEntity(blkRef2)
                acTrans.AddNewlyCreatedDBObject(blkRef2, True)

                For Each id As ObjectId In blkDef2
                    Dim dbObj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim attDef As AttributeDefinition = CType(dbObj, AttributeDefinition)
                        If Not attDef.Constant Then
                            Dim attRef As New AttributeReference()
                            If attDef.Tag = "TENDONNUMBER" Then
                                attRef.SetAttributeFromBlock(attDef, blkRef2.BlockTransform)
                                attRef.TextString = HoriTendon2Num ' <-- Your value goes here
                            ElseIf attDef.Tag = "ELONGATION" Then
                                Dim ElongNumber As Double
                                ElongNumber = Math.Round((0.007 * (HoriTendonLen * 12)) * 8) / 8
                                attRef.SetAttributeFromBlock(attDef, blkRef2.BlockTransform)
                                attRef.TextString = FormatInchesAsFraction(ElongNumber)
                            End If

                            blkRef2.AttributeCollection.AppendAttribute(attRef)
                            acTrans.AddNewlyCreatedDBObject(attRef, True)

                        End If
                    End If
                Next
            End Using
        End Sub

        Function FormatInchesAsFraction(value As Double) As String
            Dim wholeInches As Integer = Math.Floor(value)
            Dim fractionDecimal As Double = value - wholeInches

            ' Convert to nearest 1/8"
            Dim eighths As Integer = CInt(Math.Round(fractionDecimal * 8))

            ' Adjust if rounding causes whole number (e.g., 7.999 rounds to 8/8 = 1)
            If eighths = 8 Then
                wholeInches += 1
                eighths = 0
            End If

            ' Convert to fraction string
            Dim fractionText As String = ""
            If eighths > 0 Then
                fractionText = $"{eighths}/8"
                ' Optional: simplify (e.g., 4/8 -> 1/2)
                Select Case eighths
                    Case 2 : fractionText = "1/4"
                    Case 3 : fractionText = "3/8"
                    Case 4 : fractionText = "1/2"
                    Case 5 : fractionText = "5/8"
                    Case 6 : fractionText = "3/4"
                    Case 7 : fractionText = "7/8"
                End Select
            End If

            ' Combine parts
            Dim result As String = ""
            If wholeInches > 0 AndAlso fractionText <> "" Then
                result = $"{wholeInches} {fractionText}"
            ElseIf wholeInches > 0 Then
                result = $"{wholeInches}"
            Else
                result = fractionText
            End If

            Return result & """"

        End Function

        Private Sub VertSlabTendonInsertAdi(acDoc As Document, acCurDb As Database, aced As Editor, acTrans As Transaction, HeadIpt As Point3d, VertTendon2Num As Integer, VertTendonLen As String, BlockLayer As String)

            Dim Dimscale As Double = 1
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim BlockName As String
            Dim BlockColor As Integer
            BlockName = "WestTendonHead"
            BlockColor = 3

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has(BlockLayer) = False Then
                Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                '' Assign the layer the ACI color 1 and a name
                acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, BlockColor)
                acLyrTblRec.Name = BlockLayer

                '' Upgrade the Layer table for write
                acLyrTbl.UpgradeOpen()

                '' Append the new layer to the Layer table and the transaction
                acLyrTbl.Add(acLyrTblRec)
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            ElseIf acLyrTbl.Has(BlockLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            End If

            ' to force update drawing screen
            acDoc.TransactionManager.EnableGraphicsFlush(True)
            Dim bt As BlockTable = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite), BlockTable)
            Dim tsTbl As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
            Dim textstylename As String = "TendonTagging"

            If tsTbl.Has(textstylename) Then

                Dim TxtStyleVar As Object
                'Get the TextStylye SYSTEM VARIABLE
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                'SET TextStyle SYSTEM VARIABLE
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "TendonTagging")

            Else

                Dim st As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                Dim strr As TextStyleTableRecord = New TextStyleTableRecord()

                strr.Name = "TendonTagging"
                st.Add(strr)
                strr.FileName = "romans.shx"
                strr.ObliquingAngle = 0.0
                strr.XScale = 0.8
                strr.TextSize = 0.0
                strr.IsVertical = False
                strr.IsShapeFile = False
                acTrans.AddNewlyCreatedDBObject(strr, True)

                'make as current
                Dim TxtStyleVar As Object
                'Get the TextStylye SYSTEM VARIABLE
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                'SET TextStyle SYSTEM VARIABLE
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "TendonTagging")

            End If

            ' if the block table doesn't already exists, exit
            Dim ed As Editor = acDoc.Editor
            Dim blktb As BlockTable = TryCast(acCurDb.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)

            ' Get the block definitions
            Dim blkDef1 As BlockTableRecord = TryCast(blktb(BlockName).GetObject(OpenMode.ForRead), BlockTableRecord)
            Dim blkDef2 As BlockTableRecord = TryCast(blktb("VertTendonNumber").GetObject(OpenMode.ForRead), BlockTableRecord)

            ' Open modelspace (or paperspace) for writing
            Dim BlkTbl As BlockTableRecord = DirectCast(acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            ' === Block 1: original block ===
            Using blkRef1 As New BlockReference(HeadIpt, blkDef1.ObjectId)
                blkRef1.Layer = BlockLayer
                blkRef1.ScaleFactors = New Scale3d(Dimscale, Dimscale, Dimscale)
                blkRef1.Rotation = Math.PI * 1.5
                BlkTbl.AppendEntity(blkRef1)

                acTrans.AddNewlyCreatedDBObject(blkRef1, True)

                ' (optional) add attributes if any
            End Using

            ' === Block 2: "Tendon Number" block ===
            ' Let's say you want it a bit offset from the first one:

            Using blkRef2 As New BlockReference(HeadIpt, blkDef2.ObjectId)
                blkRef2.Layer = BlockLayer
                blkRef2.ScaleFactors = New Scale3d(Dimscale, Dimscale, Dimscale)
                BlkTbl.AppendEntity(blkRef2)
                acTrans.AddNewlyCreatedDBObject(blkRef2, True)

                For Each id As ObjectId In blkDef2
                    Dim dbObj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim attDef As AttributeDefinition = CType(dbObj, AttributeDefinition)
                        If Not attDef.Constant Then
                            Dim attRef As New AttributeReference()
                            If attDef.Tag = "TENDONNUMBER" Then
                                attRef.SetAttributeFromBlock(attDef, blkRef2.BlockTransform)
                                attRef.TextString = VertTendon2Num ' <-- Your value goes here
                            ElseIf attDef.Tag = "ELONGATION" Then
                                Dim ElongNumber As Double
                                ElongNumber = Math.Round((0.007 * (VertTendonLen * 12)) * 8) / 8
                                attRef.SetAttributeFromBlock(attDef, blkRef2.BlockTransform)
                                attRef.TextString = FormatInchesAsFraction(ElongNumber)
                            End If

                            blkRef2.AttributeCollection.AppendAttribute(attRef)
                            acTrans.AddNewlyCreatedDBObject(attRef, True)

                        End If
                    End If
                Next
            End Using

        End Sub

        Private Shared Function IsLineOrPolyline(ent As Entity) As Boolean
            Return TypeOf ent Is Line OrElse TypeOf ent Is Polyline
        End Function

        Private Shared Sub AddIntersectionsWithLine(ent As Entity, line As Line, points As Point3dCollection)
            ent.IntersectWith(line, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero)
        End Sub

        Private Shared Sub SetTendonToSlabBounds(slabInt As Point3d, boundaryTendonInt As Point3d, fLineAngD As Integer,
                                                 distanceFromSlab As Double, ByRef startPoint As Point3d, ByRef endPoint As Point3d,
                                                 ByRef startWithTendon As Boolean, ByRef endWithTendon As Boolean,
                                                 ByRef lengthOfLine As Double, ByRef numberOfTendons As Double,
                                                 ByRef newTendonSpacing As Double, spacing As Double)
            startWithTendon = False
            endWithTendon = False

            If fLineAngD = 0 Or fLineAngD = 180 Or fLineAngD = 360 Then
                Dim slabOffsetPoint As Point3d
                If slabInt.X < boundaryTendonInt.X Then
                    slabOffsetPoint = New Point3d(slabInt.X + distanceFromSlab, slabInt.Y, 0)
                    startPoint = slabOffsetPoint
                    endPoint = New Point3d(boundaryTendonInt.X, boundaryTendonInt.Y, 0)
                    endWithTendon = True
                Else
                    slabOffsetPoint = New Point3d(slabInt.X - distanceFromSlab, slabInt.Y, 0)
                    startPoint = New Point3d(boundaryTendonInt.X, boundaryTendonInt.Y, 0)
                    endPoint = slabOffsetPoint
                    startWithTendon = True
                End If

                lengthOfLine = Math.Abs(startPoint.X - endPoint.X)
            Else
                Dim slabOffsetPoint As Point3d
                If slabInt.Y > boundaryTendonInt.Y Then
                    slabOffsetPoint = New Point3d(slabInt.X, slabInt.Y - distanceFromSlab, 0)
                    startPoint = slabOffsetPoint
                    endPoint = New Point3d(boundaryTendonInt.X, boundaryTendonInt.Y, 0)
                    endWithTendon = True
                Else
                    slabOffsetPoint = New Point3d(slabInt.X, slabInt.Y + distanceFromSlab, 0)
                    startPoint = New Point3d(boundaryTendonInt.X, boundaryTendonInt.Y, 0)
                    endPoint = slabOffsetPoint
                    startWithTendon = True
                End If

                lengthOfLine = Math.Abs(startPoint.Y - endPoint.Y)
            End If

            numberOfTendons = Math.Ceiling(Math.Abs(lengthOfLine) / spacing)
            newTendonSpacing = Math.Abs(lengthOfLine / numberOfTendons)
        End Sub

        Private Shared Sub SetSlabExtents(ent As Entity, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
            Dim pline As Polyline = TryCast(ent, Polyline)
            If pline IsNot Nothing Then
                Dim firstPt As Point2d = pline.GetPoint2dAt(0)
                minX = firstPt.X
                maxX = firstPt.X
                minY = firstPt.Y
                maxY = firstPt.Y

                For i As Integer = 1 To pline.NumberOfVertices - 1
                    Dim pt As Point2d = pline.GetPoint2dAt(i)
                    If pt.X < minX Then minX = pt.X
                    If pt.X > maxX Then maxX = pt.X
                    If pt.Y < minY Then minY = pt.Y
                    If pt.Y > maxY Then maxY = pt.Y
                Next
            Else
                Dim ext As Extents3d = ent.GeometricExtents
                minX = ext.MinPoint.X
                maxX = ext.MaxPoint.X
                minY = ext.MinPoint.Y
                maxY = ext.MaxPoint.Y
            End If

            minX -= 12
            maxX += 12
            minY -= 12
            maxY += 12
        End Sub

        Private Sub InsertTendonTail(acDoc As Document, acCurDb As Database, aced As Editor, acTrans As Transaction, TailRot As Double, TailIpt As Point3d)

            Dim Dimscale As Double = 48.0
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim BlockName As String = "Tendon Tail"
            Dim BlockLayer As String = "S-FND-DBL"
            Dim BlockColor As Integer = 3

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has(BlockLayer) = False Then
                Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                '' Assign the layer the ACI color 1 and a name
                acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, BlockColor)
                acLyrTblRec.Name = BlockLayer

                '' Upgrade the Layer table for write
                acLyrTbl.UpgradeOpen()

                '' Append the new layer to the Layer table and the transaction
                acLyrTbl.Add(acLyrTblRec)
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            ElseIf acLyrTbl.Has(BlockLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(BlockLayer)

            End If

            'Dim ucs As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem

            ' to force update drawing screen
            acDoc.TransactionManager.EnableGraphicsFlush(True)
            Dim bt As BlockTable = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite), BlockTable)
            Dim tsTbl As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
            Dim textstylename As String = "CABLES"

            If tsTbl.Has(textstylename) Then

                Dim TxtStyleVar As Object
                'Get the TextStylye SYSTEM VARIABLE
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                'SET TextStyle SYSTEM VARIABLE
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "CABLES")

            Else

                Dim st As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                Dim strr As TextStyleTableRecord = New TextStyleTableRecord()

                strr.Name = "CABLES"
                st.Add(strr)
                strr.FileName = "romans.shx"
                strr.ObliquingAngle = 0.0
                strr.XScale = 1.0
                strr.TextSize = 0.0
                strr.IsVertical = False
                strr.IsShapeFile = False
                acTrans.AddNewlyCreatedDBObject(strr, True)
                Dim TxtStyleVar As Object
                TxtStyleVar = Bricscad.ApplicationServices.Application.GetSystemVariable("TextStyle")
                Bricscad.ApplicationServices.Application.SetSystemVariable("TextStyle", "CABLES")

            End If

            If Not bt.Has(BlockName) Then

                Dim inspt As New Point3d(0, 0, 0)
                Dim Btr As New BlockTableRecord()
                bt.Add(Btr)
                Btr.Name = BlockName
                Btr.Origin = inspt
                Btr.BlockScaling = BlockScaling.Uniform
                Btr.Units = UnitsValue.Inches
                Btr.Explodable = True
                acTrans.AddNewlyCreatedDBObject(Btr, True)

                Dim ln As New Teigha.DatabaseServices.Line()
                ln.StartPoint = New Point3d(-0.0673, 0.0577, 0)
                ln.EndPoint = New Point3d(0.0673, 0.0577, 0)
                ln.Layer = "0"

                Btr.AppendEntity(ln)
                acTrans.AddNewlyCreatedDBObject(ln, True)

            End If

            Dim blktb As BlockTable = TryCast(acCurDb.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)
            Dim blkDef As BlockTableRecord = TryCast(blktb(BlockName).GetObject(OpenMode.ForRead), BlockTableRecord)
            Dim BlkTbl As BlockTableRecord = DirectCast(acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            Using blkRef As New BlockReference(TailIpt, blkDef.ObjectId)

                blkRef.Layer = BlockLayer
                blkRef.ScaleFactors = New Scale3d(Dimscale, Dimscale, Dimscale)
                blkRef.Rotation = TailRot

                BlkTbl.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)

            End Using

        End Sub
    End Class
End Namespace