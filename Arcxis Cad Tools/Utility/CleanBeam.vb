' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.IO

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.CleanBeam))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class CleanBeam

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
        <CommandMethod("clb", CommandFlags.Modal)>
        Public Sub CleanBeam() ' This method can have any name
            ' Put your command code here

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Get the window coordinates

            Dim ppo As New PromptPointOptions(vbLf & "Specify first corner:")
            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

            If ppr.Status <> PromptStatus.OK Then

                Return
            End If

            Dim PP1 As Point3d = ppr.Value
            Dim pco As New PromptCornerOptions(vbLf & "Specify opposite corner: ", ppr.Value)

            pco.UseDashedLine = True
            ppr = ed.GetCorner(pco)

            If ppr.Status <> PromptStatus.OK Then

                Return
            End If

            Dim PP2 As Point3d = ppr.Value

            IntersectLines(PP1, PP2)
            Eraselines(PP1, PP2)

        End Sub

        Private Sub Eraselines(PP1 As Point3d, PP2 As Point3d)

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim Ent As Entity
            Dim Ent2 As Entity
            Dim Ent3 As Entity
            Dim lay2 As String
            Dim type2 As String
            Dim line2 As Polyline
            Dim pts As New Point3dCollection()

            Using myTrans As Transaction = doc.TransactionManager.StartTransaction

                Dim tvv As TypedValue() = New TypedValue() {New TypedValue(0, "*LINE"), New TypedValue(DxfCode.LayerName, "FDPSLDR,S-FND-SLABDP")}
                Dim sff As SelectionFilter = New SelectionFilter(tvv)
                Dim result2 As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.SelectCrossingWindow(PP1, PP2, sff)

                If result2.Status = PromptStatus.OK Then

                    For Each obj2 As ObjectId In result2.Value.GetObjectIds

                        Dim rescount As Integer = result2.Value.Count

                        Ent2 = obj2.GetObject(OpenMode.ForRead)
                        lay2 = Ent2.Layer
                        type2 = obj2.ObjectClass.Name()

                        line2 = TryCast(Ent2, Polyline)

                        Dim vn As Integer = line2.NumberOfVertices

                        For i As Integer = 0 To vn - 1

                            Dim pt As Point3d = line2.GetPoint3dAt(i)

                            pts.Add(pt)

                        Next

                        Dim clpt As Point3d = pts(0)

                        pts.Add(clpt)

                        Dim tv As TypedValue() = New TypedValue() {New TypedValue(0, "*LINE"), New TypedValue(DxfCode.LayerName, "FDSHORT,S-FND-BEAM")}
                        Dim sf As SelectionFilter = New SelectionFilter(tv)
                        Dim result As PromptSelectionResult = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.SelectFence(pts, sf)
                        Dim result3 As PromptSelectionResult = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.SelectWindowPolygon(pts, sf)

                        If (result.Status = PromptStatus.OK) Then

                            For Each Obj As ObjectId In result.Value.GetObjectIds

                                Ent = Obj.GetObject(OpenMode.ForWrite)

                                Ent.Erase()

                            Next

                        End If

                        If (result3.Status = PromptStatus.OK) Then

                            For Each Obj3 As ObjectId In result3.Value.GetObjectIds

                                Ent3 = Obj3.GetObject(OpenMode.ForWrite)
                                Ent3.Linetype = "bylayer"

                            Next

                        End If

                    Next

                End If

                myTrans.Commit()
            End Using


        End Sub

        Private Sub ExtendLines2(PP1 As Point3d, PP2 As Point3d)

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim Ent As Entity
            Dim lay As String
            Dim type As String
            Dim Ent2 As Entity
            Dim lay2 As String
            Dim type2 As String
            Dim line As Line
            Dim line2 As Polyline
            Dim pts As New Point3dCollection()

            Dim tv As TypedValue() = New TypedValue() {New TypedValue(0, "*LINE"), New TypedValue(DxfCode.LayerName, "FDSHORT,S-FND-BEAM")}
            Dim sf As SelectionFilter = New SelectionFilter(tv)
            Dim psr As PromptSelectionResult = doc.Editor.SelectCrossingWindow(PP1, PP2, sf)

            If psr.Status = PromptStatus.OK Then

                Using acLckDoc As DocumentLock = doc.LockDocument()

                    Using myTrans As Transaction = doc.TransactionManager.StartTransaction
                        For Each Obj As ObjectId In psr.Value.GetObjectIds

                            Ent = Obj.GetObject(OpenMode.ForWrite)

                            lay = Ent.Layer
                            type = Obj.ObjectClass.Name()
                            line = TryCast(Ent, Line)

                            Dim tvv As TypedValue() = New TypedValue() {New TypedValue(0, "*LINE"), New TypedValue(DxfCode.LayerName, "FDPSLDR, S-FND-SLABDP")}
                            Dim sff As SelectionFilter = New SelectionFilter(tvv)
                            Dim result2 As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.SelectCrossingWindow(PP1, PP2, sff)

                            If result2.Status = PromptStatus.OK Then

                                For Each obj2 As ObjectId In result2.Value.GetObjectIds

                                    Ent2 = obj2.GetObject(OpenMode.ForRead)
                                    lay2 = Ent2.Layer
                                    type2 = obj2.ObjectClass.Name()
                                    line2 = TryCast(Ent2, Polyline)

                                Next
                                ''''''''''''''''''''''''''''''
                            End If

                        Next

                        myTrans.Commit()
                    End Using

                End Using
            End If

        End Sub


        Private Sub IntersectLines(PP1 As Point3d, PP2 As Point3d)

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database

            Dim Ent As Entity
            Dim Entt As Entity
            Dim lay As String
            Dim type As String

            Dim tv As TypedValue() = New TypedValue() {New TypedValue(0, "*LINE"), New TypedValue(DxfCode.LayerName, "FDSHORT, S-FND-BEAM")}

            Dim sf As SelectionFilter = New SelectionFilter(tv)
            Dim psr As PromptSelectionResult = doc.Editor.SelectCrossingWindow(PP1, PP2, sf)

            If psr.Status = PromptStatus.OK Then

                Using acLckDoc As DocumentLock = doc.LockDocument()

                    Using myTrans As Transaction = doc.TransactionManager.StartTransaction

                        Dim objs As ObjectId() = psr.Value.GetObjectIds()
                        Dim rids As New List(Of ObjectId)()
                        Dim btr As BlockTableRecord = TryCast(myTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                        For Each Obj As ObjectId In psr.Value.GetObjectIds

                            Entt = Obj.GetObject(OpenMode.ForWrite)
                            lay = Entt.Layer
                            type = Obj.ObjectClass.Name()
                            rids.Add(Obj)

                        Next

                        For Each subid As SelectedObject In psr.Value
                            Ent = TryCast(myTrans.GetObject(subid.ObjectId, OpenMode.ForRead, False), Entity)

                            lay = Ent.Layer
                            Dim selobjid As ObjectId = subid.ObjectId
                            type = selobjid.ObjectClass.Name()

                            Dim ln1 As Line = TryCast(Ent, Line)
                            Dim points As New List(Of Point3d)()
                            Dim pts As New Point3dCollection()
                            For Each id As ObjectId In rids

                                If id = subid.ObjectId Then
                                    Continue For
                                End If

                                Dim [next] As Entity = TryCast(myTrans.GetObject(id, OpenMode.ForRead, False), Entity)
                                Dim ln2 As Line = TryCast([next], Line)
                                ln1.IntersectWith(ln2, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero)

                            Next

                            For Each p As Point3d In pts
                                points.Add(p)
                            Next
                            Dim sp As Point3d = ln1.StartPoint
                            Dim ep As Point3d = ln1.EndPoint
                            If Not points.Contains(sp) Then
                                points.Add(sp)
                            End If
                            If Not points.Contains(ep) Then
                                points.Add(ep)
                            End If

                            points.Sort(Function(a As Point3d, b As Point3d) Convert.ToInt32(Convert.ToDouble(sp.DistanceTo(a).CompareTo(Convert.ToDouble(sp.DistanceTo(b))))))

                            For n As Integer = 0 To points.Count - 2
                                Dim lin As New Line(points(n), points(n + 1))

                                'lin.Layer = "FDSHORT"
                                'lin.Linetype = "By Layer"

                                Dim linlen As Integer = lin.Length

                                If linlen <> 12 AndAlso pts.Count <> 0 Then

                                    btr.AppendEntity(lin)
                                    myTrans.AddNewlyCreatedDBObject(lin, True)

                                End If

                            Next

                        Next

                        For Each id As ObjectId In rids
                            Ent = TryCast(myTrans.GetObject(id, OpenMode.ForRead), Entity)
                            If Not Ent.IsWriteEnabled Then
                                Ent.UpgradeOpen()
                            End If
                            Ent.[Erase]()
                        Next

                        myTrans.Commit()
                    End Using

                End Using
            End If

        End Sub

    End Class

End Namespace