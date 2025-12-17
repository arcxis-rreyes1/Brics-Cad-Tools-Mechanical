' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Linq

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.AddTendonsPTS))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class AddTendonsPTS

        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx
        <CommandMethod("ATTN")>
        Public Sub AddTendonsPTS()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim CableStartList As New List(Of Point3d)
            Dim CableEndList As New List(Of Point3d)
            Dim TendonPointsList As New List(Of Point3d)
            Dim SortedTendonPointsList As New List(Of Point3d)
            Dim SortedTendonPoints As New List(Of Point3d)
            Dim SortedTendonXPointsList As New List(Of Point3d)

            Dim LineList As New List(Of Entity)
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
            Dim PtAndID As New Dictionary(Of Double, String)
            'Dim ltsc As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("ltscale")
            Dim acLine As Line

            '' Set system variable to new value
            'Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FIELDDISPLAY", 0)

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    '' Get the current value from a system variable
                    Dim ech As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("cmdecho")
                    Dim otm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("orthomode")

                    '' Set system variable to new value
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("orthomode", 1)

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

                    '' Create a line that starts at 5,5 and ends at 12,3
                    acLine = New Line(ptStart, ptEnd)
                    acLine.Layer = "0"
                    acLine.Linetype = "Hidden"
                    Dim FLineAng As Double = acLine.Angle
                    Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI

                    Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE,LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "S-FND-BTEND")}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectFence(pts, acSelFtr)

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each SSObj As SelectedObject In SS

                                Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                                If typ = "AcDbPolyline" OrElse typ = "AcDbLine" Then

                                    Dim acEnt As Object = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                    LineList.Add(acEnt)
                                    Dim ID As String = acEnt.objectid.ToString
                                    Dim CableColor As String = acEnt.ColorIndex
                                    Dim IDpart As String() = ID.Split("(")
                                    Dim IDpart1 As String = IDpart(1)
                                    Dim IDpart2 As String() = IDpart1.Split(")")
                                    Dim CableID As String = IDpart2(0)
                                    Dim Cablelen As Double = acEnt.Length
                                    Dim Cablelen2 As Double = acEnt.Length / 2
                                    Dim CableStart As Point3d = acEnt.StartPoint
                                    Dim CableEnd As Point3d = acEnt.EndPoint
                                    CableStartList.Add(CableStart)
                                    CableEndList.Add(CableEnd)
                                    CablePointsList.Add(CableStart)
                                    CablePointsList.Add(CableEnd)

                                    If FLineAngD = 0 Then

                                        SortedCablePointsList = (From Xpnt In CablePointsList Order By Xpnt.X Select Xpnt).ToList

                                        For Each xpt As Point3d In SortedCablePointsList

                                            Dim XPOINT As Double = Math.Round(xpt.X, 10)

                                            If Not SortedCableStartPoints.Contains(XPOINT) Then

                                                SortedCableStartPoints.Add(XPOINT)
                                                PtAndID.Add(XPOINT, CableID)

                                            End If

                                        Next

                                    ElseIf FLineAngD = 90 OrElse FLineAngD = -270 Then

                                        SortedCablePointsList = (From Ypnt In CablePointsList Order By Ypnt.Y Select Ypnt).ToList

                                        For Each ypt As Point3d In SortedCablePointsList

                                            Dim YPOINT As Double = Math.Round(ypt.Y, 10)

                                            If Not SortedCableStartPoints.Contains(YPOINT) Then

                                                SortedCableStartPoints.Add(YPOINT)
                                                PtAndID.Add(YPOINT, CableID)

                                            End If

                                        Next

                                    ElseIf FLineAngD = 180 OrElse FLineAngD = -180 Then

                                        SortedCablePointsList = (From Xpnt In CablePointsList Order By Xpnt.X Select Xpnt).ToList

                                        For Each xpt As Point3d In SortedCablePointsList

                                            Dim XPOINT As Double = Math.Round(xpt.X, 10)

                                            If Not SortedCableStartPoints.Contains(XPOINT) Then

                                                SortedCableStartPoints.Add(XPOINT)
                                                PtAndID.Add(XPOINT, CableID)

                                            End If

                                        Next

                                    ElseIf FLineAngD = 270 OrElse FLineAngD = -90 Then

                                        SortedCablePointsList = (From Ypnt In CablePointsList Order By Ypnt.Y Select Ypnt).ToList

                                        For Each ypt As Point3d In SortedCablePointsList

                                            Dim YPOINT As Double = Math.Round(ypt.Y, 10)

                                            If Not SortedCableStartPoints.Contains(YPOINT) Then

                                                SortedCableStartPoints.Add(YPOINT)
                                                PtAndID.Add(YPOINT, CableID)

                                            End If

                                        Next


                                    End If

                                End If

                            Next


                            SortedCableXPointsList = (From Xpnt In CablePointsList Order By Xpnt.X Select Xpnt).ToList
                            For Each pt As Point3d In SortedCableXPointsList
                                'aced.WriteMessage(vbLf & "pt = " & pt.ToString())
                            Next

                            Dim CableXPT As Point3d = SortedCableXPointsList(0)
                            Dim CableX As Double = Math.Round(CableXPT.X, 10)
                            Dim CableXXPT As Point3d = SortedCableXPointsList(SortedCableXPointsList.Count - 1)

                            SortedCableYPointsList = (From Ypnt In CablePointsList Order By Ypnt.Y Select Ypnt).ToList
                            For Each pt As Point3d In SortedCableYPointsList
                                'aced.WriteMessage(vbLf & "pt = " & pt.ToString())
                            Next

                            Dim CableYPT As Point3d = SortedCableYPointsList(0)
                            Dim CableYYPT As Point3d = SortedCableYPointsList(SortedCableYPointsList.Count - 1)
                            Dim CableY As Double = Math.Round(CableYPT.Y, 10)

                            SortedCableStartPoints.Sort()

                            For Each pt As Double In SortedCableStartPoints
                                'aced.WriteMessage(vbLf & pt.ToString())
                            Next

                            Dim counter As Integer = 0

                            While counter < SortedCableStartPoints.Count - 1

                                Dim CableStart1 As Double = SortedCableStartPoints(counter)
                                Dim CableStart2 As Double = SortedCableStartPoints(counter + 1)
                                Dim CID As String = "0"
                                Dim pair As KeyValuePair(Of Double, String)

                                For Each pair In PtAndID

                                    Dim PID As Double = pair.Key

                                    If PID = CableStart1 Then

                                        CID = pair.Value

                                    End If

                                Next

                                Dim Beamdist As Double = CableStart2 - CableStart1

                                If Beamdist > 60 Then

                                    Dim tendonspaces As Integer = Math.Floor(Beamdist / 60) + 1
                                    Dim tendonoffset As Double = Beamdist / tendonspaces
                                    Dim LineStart As Point3d
                                    Dim LineEnd As Point3d
                                    Dim LineLength As Double

                                    For Each BeamEnt As Entity In LineList

                                        Dim ln As Line = TryCast(BeamEnt, Line)
                                        LineStart = ln.StartPoint
                                        LineEnd = ln.EndPoint
                                        LineLength = ln.Length
                                        Dim ID As String = BeamEnt.ObjectId.ToString
                                        Dim CableColor As String = BeamEnt.ColorIndex
                                        Dim LineIDpart As String() = ID.Split("(")
                                        Dim LineIDpart1 As String = LineIDpart(1)
                                        Dim LineIDpart2 As String() = LineIDpart1.Split(")")
                                        Dim LineID As String = LineIDpart2(0)
                                        Dim LineAng As Double = ln.Angle
                                        Dim LineAngD As Integer = LineAng * 180.0 / Math.PI
                                        Dim StartOff As Point3d
                                        Dim EndOff As Point3d

                                        If LineAngD = 90 OrElse LineAngD = -270 Then

                                            LineStart = New Point3d(LineStart.X, CableYPT.Y, 0)
                                            LineEnd = New Point3d(LineEnd.X, CableYYPT.Y, 0)

                                        ElseIf LineAngD = 270 OrElse LineAngD = -90 Then


                                            LineStart = New Point3d(LineStart.X, CableYYPT.Y, 0)
                                            LineEnd = New Point3d(LineEnd.X, CableYPT.Y, 0)


                                        ElseIf LineAngD = 0 OrElse LineAngD = -180 OrElse LineAngD = 360 Then

                                            LineStart = New Point3d(CableXPT.X, LineStart.Y, 0)
                                            LineEnd = New Point3d(CableXXPT.X, LineEnd.Y, 0)

                                        ElseIf LineAngD = 180 OrElse LineAngD = -360 Then

                                            LineStart = New Point3d(CableXXPT.X, LineStart.Y, 0)
                                            LineEnd = New Point3d(CableXPT.X, LineEnd.Y, 0)

                                        End If

                                        Dim EXTLineStart As Point3d = New Point3d(LineStart.X - 36 * Math.Cos(LineAng), LineStart.Y - 36 * Math.Sin(LineAng), LineStart.Z)
                                        Dim EXTLineEnd As Point3d = New Point3d(LineEnd.X + 36 * Math.Cos(LineAng), LineEnd.Y + 36 * Math.Sin(LineAng), LineEnd.Z)


                                        If CID = LineID Then

                                            For i As Integer = 1 To tendonspaces - 1

                                                If LineAngD = 90 OrElse LineAngD = 270 OrElse LineAngD = -90 OrElse LineAngD = -270 Then

                                                    StartOff = New Point3d(EXTLineStart.X + (tendonoffset * i) * Math.Cos(0.0), EXTLineStart.Y + (tendonoffset * i) * Math.Sin(0.0), EXTLineStart.Z)
                                                    EndOff = New Point3d(EXTLineEnd.X + (tendonoffset * i) * Math.Cos(0.0), EXTLineEnd.Y + (tendonoffset * i) * Math.Sin(0.0), EXTLineEnd.Z)

                                                ElseIf LineAngD = 0 OrElse LineAngD = 180 OrElse LineAngD = -180 OrElse LineAngD = 360 Then

                                                    StartOff = New Point3d(EXTLineStart.X + (tendonoffset * i) * Math.Cos(1.570796), EXTLineStart.Y + (tendonoffset * i) * Math.Sin(1.570796), EXTLineStart.Z)
                                                    EndOff = New Point3d(EXTLineEnd.X + (tendonoffset * i) * Math.Cos(1.570796), EXTLineEnd.Y + (tendonoffset * i) * Math.Sin(1.570796), EXTLineEnd.Z)

                                                End If


                                                Dim TendoLine As Line
                                                TendoLine = New Line(StartOff, EndOff)
                                                TendoLine.Layer = "S-FND-STEND"
                                                Dim SlabIptCol As New Point3dCollection()
                                                SlabIptCol.Clear()
                                                Dim SlabTendonIptList As New List(Of Point3d)
                                                SlabTendonIptList.Clear()
                                                Dim SortedSlabTendonPts As New List(Of Point3d)
                                                SortedSlabTendonPts.Clear()
                                                Dim Fencepts As New Point3dCollection()
                                                Fencepts.Add(StartOff)
                                                Fencepts.Add(EndOff)

                                                Dim acTypValArR As TypedValue() = New TypedValue() {New TypedValue(DxfCode.Start, "LINE,*POLYLINE"), New TypedValue(DxfCode.LayerName, "S-FND-SLABDP")}
                                                Dim acSelFtrR As SelectionFilter = New SelectionFilter(acTypValArR)
                                                Dim prSelResS As PromptSelectionResult = acDoc.Editor.SelectFence(Fencepts, acSelFtrR)

                                                If prSelResS.Status = PromptStatus.OK Then

                                                    Dim SSS As SelectionSet = prSelResS.Value

                                                    For Each SSSObj As SelectedObject In SSS

                                                        Dim typ As String = SSSObj.ObjectId.ObjectClass.Name()

                                                        If typ = "AcDbPolyline" OrElse typ = "AcDbLine" Then

                                                            Dim SlabEnt As Object = acTrans.GetObject(SSSObj.ObjectId, OpenMode.ForRead, False, True)
                                                            'SlabEnt.IntersectWith(TendoLine, Intersect.OnBothOperands, SlabIptCol, IntPtr.Zero, IntPtr.Zero)
                                                            SlabEnt.IntersectWith(TendoLine, Intersect.ExtendBoth, SlabIptCol, IntPtr.Zero, IntPtr.Zero)
                                                            'TendoLine.IntersectWith(SlabEnt, Intersect.ExtendBoth, SlabIptCol, IntPtr.Zero, IntPtr.Zero)

                                                        End If
                                                    Next

                                                End If

                                                If SlabIptCol.Count <> 0 Then


                                                    For Each STpt As Point3d In SlabIptCol

                                                        If Not SlabTendonIptList.Contains(STpt) Then

                                                            SlabTendonIptList.Add(STpt)

                                                        End If

                                                    Next

                                                End If

                                                For Each ppt As Point3d In SlabTendonIptList

                                                    'aced.WriteMessage(vbLf & "ppt = " & ppt.ToString)

                                                Next

                                                If LineAngD = 0 OrElse LineAngD = 360 Then

                                                    SortedSlabTendonPts = (From Xpnt In SlabTendonIptList Order By Xpnt.X Select Xpnt).ToList
                                                    StartOff = SortedSlabTendonPts(0)
                                                    EndOff = SortedSlabTendonPts(1)

                                                ElseIf LineAngD = 90 OrElse LineAngD = -270 Then

                                                    SortedSlabTendonPts = (From Ypnt In SlabTendonIptList Order By Ypnt.Y Select Ypnt).ToList
                                                    StartOff = SortedSlabTendonPts(0)
                                                    EndOff = SortedSlabTendonPts(1)

                                                ElseIf LineAngD = 180 OrElse LineAngD = -180 Then

                                                    SortedSlabTendonPts = (From Xpnt In SlabTendonIptList Order By Xpnt.X Select Xpnt).ToList
                                                    StartOff = SortedSlabTendonPts(1)
                                                    EndOff = SortedSlabTendonPts(0)

                                                ElseIf LineAngD = 270 OrElse LineAngD = -90 Then

                                                    SortedSlabTendonPts = (From Ypnt In SlabTendonIptList Order By Ypnt.Y Select Ypnt).ToList
                                                    StartOff = SortedSlabTendonPts(1)
                                                    EndOff = SortedSlabTendonPts(0)

                                                Else

                                                    SortedSlabTendonPts = (From pnt In SlabTendonIptList Order By pnt.X, pnt.Y, pnt.Z Select pnt).ToList
                                                    StartOff = SortedSlabTendonPts(0)
                                                    EndOff = SortedSlabTendonPts(1)

                                                End If

                                                'TendoLine = New Line(StartOff, EndOff)
                                                TendoLine.StartPoint = StartOff
                                                TendoLine.EndPoint = EndOff
                                                acBlkTblRec.AppendEntity(TendoLine)
                                                acTrans.AddNewlyCreatedDBObject(TendoLine, True)

                                            Next

                                        End If

                                    Next

                                End If

                                counter = counter + 1

                            End While

                        End If

                    Else
                        acDoc.Editor.WriteMessage(vbLf & "No Beam Selected!")

                    End If

                    '' Save the changes and dispose of the transaction
                    acTrans.Commit()

                End Using

            End Using

        End Sub

    End Class

End Namespace