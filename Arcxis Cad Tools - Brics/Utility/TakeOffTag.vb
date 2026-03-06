' (C) Copyright 2011 by  
'
Imports System
Imports System.Linq
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools_Brics.TakeOffTag))>
Namespace Arcxis_Cad_Tools_Brics

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class TakeOffTag

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

        Shared prevLabelType As String = "None"
        Shared prevTagType As String = "None"
        Shared prevRafterSlope As String = "None"
        Shared prevRafterTagLocation As String = "None"
        Shared prevGrade As String = "None"
        Shared prevMemberSizeButton As String = "None"
        Shared prevIJoistSize As String = "None"
        Shared prevIJoistSeries As String = "None"
        Shared prevJoistSpace As String = "None"
        Shared prevRafterSize As String = "None"
        Shared prevFenceTag As String = "None"
        Shared prevStorage As String = "None"
        Shared prevRoundUpRafter As String = "None"

        <CommandMethod("TOT")>
        Public Sub TakeOffTag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim frm As New Form_TakeOffTag

            If prevLabelType = "Tags" Then

                frm.RadioButton32.Checked = True

            ElseIf prevLabelType = "TakeOff" Then

                frm.RadioButton33.Checked = True

            End If

            If prevTagType = "Rafter" Then

                frm.RadioButton1.Checked = True

            ElseIf prevTagType = "CJoist" Then

                frm.RadioButton2.Checked = True

            ElseIf prevTagType = "TJI" Then

                frm.RadioButton3.Checked = True

            ElseIf prevTagType = "BCI" Then

                frm.RadioButton4.Checked = True

            ElseIf prevTagType = "Purlin" Then

                frm.RadioButton5.Checked = True

            ElseIf prevTagType = "Clip" Then

                frm.RadioButton6.Checked = True

            End If

            If prevRafterTagLocation = "Center" Then

                frm.RadioButton8.Checked = True

            ElseIf prevRafterTagLocation = "Left" Then

                frm.RadioButton7.Checked = True

            ElseIf prevRafterTagLocation = "Right" Then

                frm.RadioButton9.Checked = True

            End If

            If prevGrade = "#2" Then

                frm.RadioButton10.Checked = True

            ElseIf prevGrade = "#3" Then

                frm.RadioButton11.Checked = True

            End If

            If prevStorage = "No Storage" Then

                frm.RadioButton12.Checked = True

            ElseIf prevStorage = "With Storage" Then

                frm.RadioButton13.Checked = True

            End If

            If prevMemberSizeButton = "2x6" Then

                frm.RadioButton14.Checked = True

            ElseIf prevMemberSizeButton = "2x8" Then

                frm.RadioButton15.Checked = True

            ElseIf prevMemberSizeButton = "2x10" Then

                frm.RadioButton16.Checked = True

            ElseIf prevMemberSizeButton = "2x12" Then

                frm.RadioButton17.Checked = True

            ElseIf prevMemberSizeButton = "Auto" Then

                frm.RadioButton18.Checked = True

            End If

            If prevRafterSlope = "None" Then

                frm.ComboBox1.SelectedIndex = frm.ComboBox1.FindStringExact("8:12")

            Else

                frm.ComboBox1.SelectedIndex = frm.ComboBox1.FindStringExact(prevRafterSlope)

            End If

            If prevJoistSpace = "None" Then

                frm.ComboBox2.SelectedIndex = frm.ComboBox2.FindStringExact("24")

            Else

                frm.ComboBox2.SelectedIndex = frm.ComboBox2.FindStringExact(prevJoistSpace)

            End If

            If prevIJoistSeries = "110" Then

                frm.RadioButton19.Checked = True

            ElseIf prevIJoistSeries = "210" Then

                frm.RadioButton20.Checked = True

            ElseIf prevIJoistSeries = "230" Then

                frm.RadioButton21.Checked = True

            ElseIf prevIJoistSeries = "360" Then

                frm.RadioButton22.Checked = True


            ElseIf prevIJoistSeries = "560" Then

                frm.RadioButton23.Checked = True

            ElseIf prevIJoistSeries = "4500" Then

                frm.RadioButton24.Checked = True

            ElseIf prevIJoistSeries = "5000" Then

                frm.RadioButton25.Checked = True

            ElseIf prevIJoistSeries = "6000" Then

                frm.RadioButton26.Checked = True

            ElseIf prevIJoistSeries = "6500" Then

                frm.RadioButton27.Checked = True

            End If

            If prevIJoistSize = "11 1/4" Then

                frm.RadioButton28.Checked = True

            ElseIf prevIJoistSize = "11 7/8" Then

                frm.RadioButton29.Checked = True

            ElseIf prevIJoistSize = "14" Then

                frm.RadioButton30.Checked = True

            ElseIf prevIJoistSize = "16" Then

                frm.RadioButton31.Checked = True

            End If

            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            frm.ShowDialog()

            If TagCancel = "Cancel" Then

                Exit Sub

            End If

            SetTextStyle()
            AddLayer()

            If TagType = "Rafter" Then

                RafterTag()

            ElseIf TagType = "CJoist" Then

                CJoistTag()

            ElseIf TagType = "Purlin" Then

                PurlinTag()

            ElseIf TagType = "TJI" Then

                TJITag()

            ElseIf TagType = "BCI" Then

                BCITag()

            ElseIf TagType = "Clip" Then

                ClipTag()

            End If

            prevLabelType = LabelType
            prevTagType = TagType
            prevRafterSlope = RafterSlope
            prevRafterTagLocation = RafterTagLocation
            prevGrade = Grade
            prevMemberSizeButton = MemberSizeButton
            prevIJoistSeries = IJoistSeries
            prevIJoistSize = IJoistSize
            prevJoistSpace = JoistSpace
            prevRafterSize = RafterSize
            prevFenceTag = FenceTag
            prevStorage = Storage
            prevRoundUpRafter = RoundUpRafter

        End Sub

        Private Sub RafterTag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    'aced.WriteMessage(vbLf & "FenceTag = " & FenceTag.ToString)

                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select Rafter with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-RAFTER,DPIS-RAFTERS")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Rafter(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-RAFTER,DPIS-RAFTERS")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If


                    If (prSelRes.Status = PromptStatus.OK) Then

                        ' There are selected entities
                        ' Put your command using pickfirst set code here
                        Dim acSSet As SelectionSet = prSelRes.Value


                        '' Step through the objects in the selection set
                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                Dim GrpDesc As String = "Group_" & LineID
                                Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                GroupDict.UpgradeOpen()
                                GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)


                                Dim Slopepart As String() = RafterSlope.Split(":")
                                Dim Slopepart1 As Integer = Slopepart(0)
                                Dim Slopepart2 As Integer = Slopepart(1)
                                Dim Slope As Double = Slopepart1 / Slopepart2

                                Dim spts As New Point3dCollection()

                                Dim Rafterlen As Double = acEnt.Length

                                Dim Rafterlen2 As Double = acEnt.Length / 2
                                Dim rise As Double = Rafterlen * Slope
                                Dim hypt As Double = Math.Sqrt((rise * rise) + (Rafterlen * Rafterlen))
                                Dim HypotLen As Double = hypt + (Slopepart1 * 0.5)

                                Dim hyp3 As Double
                                Dim hyp4 As Double
                                Dim hyp5 As Double
                                Dim hyp6 As Double
                                Dim hyp7 As Double
                                Dim hyp8 As Double
                                Dim hyp9 As Double


                                If HypotLen >= 96 Then

                                    hyp3 = HypotLen / 12
                                    hyp4 = hyp3 + 0.99999999
                                    hyp5 = Math.Floor(hyp4)
                                    hyp6 = Math.Floor(hyp5 + 1)
                                    hyp7 = hyp6 * 0.5
                                    hyp8 = Math.Floor(hyp7)
                                    hyp9 = hyp8 * 2

                                Else

                                    hyp3 = HypotLen / 12
                                    hyp4 = hyp3
                                    hyp5 = Math.Floor(hyp4)
                                    hyp6 = Math.Floor(hyp5 + 1)
                                    'hyp7 = hyp6 * 0.5
                                    'hyp8 = Math.Floor(hyp7)
                                    hyp9 = hyp6

                                End If


                                Dim RafterSP As Point3d = acEnt.StartPoint
                                Dim RafterEP As Point3d = acEnt.EndPoint
                                Dim RafterAng As Double = acEnt.Angle
                                Dim RafterAngD As Double = Math.Round((RafterAng * 180.0) / Math.PI, 0)
                                Dim RafterIptCol As New Point3dCollection()
                                RafterIptCol.Clear()
                                Dim RafterMPT As Point3d

                                Dim RafterPointsList As New List(Of Point3d)
                                Dim SortedRafterPointsList As New List(Of Point3d)
                                Dim RafterStartPointsList As New List(Of Point3d)
                                Dim SortedRafterPoints As New List(Of Point3d)
                                Dim SortedRafterXPointsList As New List(Of Point3d)
                                Dim SortedRafterYPointsList As New List(Of Point3d)
                                RafterPointsList.Add(RafterSP)
                                RafterPointsList.Add(RafterEP)

                                If RafterAngD < 0 Then

                                    RafterAngD = 360 + RafterAngD

                                End If

                                Dim RafterAngRad As Double = RafterAngD * Math.PI / 180
                                Dim TagAng As Double
                                Dim TagOffset As Point3d
                                Dim TagJus As AttachmentPoint = AttachmentPoint.MiddleMid
                                Dim TagHoriMode As TextHorizontalMode = TextHorizontalMode.TextMid
                                Dim TagVertMode As TextVerticalMode = TextVerticalMode.TextVerticalMid


                                If FenceTag = "Yes" Then

                                    acEnt.IntersectWith(acLine, Intersect.OnBothOperands, RafterIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If RafterIptCol.Count <> 0 Then

                                        RafterMPT = RafterIptCol.Item(0)
                                        'aced.WriteMessage(vbLf & "WallIptCol.Count <> 0 = " & WallIptCol.Count.ToString)
                                        If RafterAngD = 90 OrElse RafterAngD = 270 OrElse RafterAngD = -90 OrElse RafterAngD = -270 Then

                                            TagOffset = New Point3d(RafterMPT.X + 4 * Math.Cos(3.14159265), RafterMPT.Y + 4 * Math.Sin(3.14159265), RafterMPT.Z)
                                            TagAng = 1.57079633

                                        ElseIf RafterAngD = 0 OrElse RafterAngD = 180 OrElse RafterAngD = -180 OrElse RafterAngD = 360 Then

                                            TagOffset = New Point3d(RafterMPT.X + 4 * Math.Cos(1.57079633), RafterMPT.Y + 4 * Math.Sin(1.57079633), RafterMPT.Z)
                                            TagAng = 0.0

                                        End If

                                    End If

                                Else

                                    If RafterTagLocation = "Center" Then

                                        RafterMPT = New Point3d(RafterSP.X + Rafterlen2 * Math.Cos(RafterAng), RafterSP.Y + Rafterlen2 * Math.Sin(RafterAng), RafterSP.Z)
                                        If RafterAngD = 90 OrElse RafterAngD = 270 OrElse RafterAngD = -90 OrElse RafterAngD = -270 Then

                                            TagOffset = New Point3d(RafterMPT.X + 4.875 * Math.Cos(3.14159265), RafterMPT.Y + 4.875 * Math.Sin(3.14159265), RafterMPT.Z)
                                            TagAng = 1.57079633

                                        ElseIf RafterAngD = 0 OrElse RafterAngD = 180 OrElse RafterAngD = -180 OrElse RafterAngD = 360 Then

                                            TagOffset = New Point3d(RafterMPT.X + 4.875 * Math.Cos(1.57079633), RafterMPT.Y + 4.875 * Math.Sin(1.57079633), RafterMPT.Z)
                                            TagAng = 0.0

                                        End If

                                    ElseIf RafterTagLocation = "Left" Then

                                        If RafterAngD = 90 OrElse RafterAngD = 270 OrElse RafterAngD = -90 OrElse RafterAngD = -270 Then

                                            SortedRafterYPointsList = (From Ypnt In RafterPointsList Order By Ypnt.Y Select Ypnt).ToList
                                            RafterMPT = SortedRafterYPointsList(0)
                                            TagOffset = New Point3d(RafterMPT.X + 3.5 * Math.Cos(4.71238898), RafterMPT.Y + 3.5 * Math.Sin(4.71238898), RafterMPT.Z)
                                            TagAng = 1.57079633
                                            TagJus = AttachmentPoint.MiddleRight
                                            TagHoriMode = TextHorizontalMode.TextRight
                                            TagVertMode = TextVerticalMode.TextVerticalMid


                                        ElseIf RafterAngD = 0 OrElse RafterAngD = 180 OrElse RafterAngD = -180 OrElse RafterAngD = 360 Then

                                            SortedRafterXPointsList = (From Xpnt In RafterPointsList Order By Xpnt.X Select Xpnt).ToList
                                            RafterMPT = SortedRafterXPointsList(0)
                                            TagOffset = New Point3d(RafterMPT.X + 3.5 * Math.Cos(3.14159265), RafterMPT.Y + 3.5 * Math.Sin(3.14159265), RafterMPT.Z)
                                            TagAng = 0.0
                                            TagJus = AttachmentPoint.MiddleRight
                                            TagHoriMode = TextHorizontalMode.TextRight
                                            TagVertMode = TextVerticalMode.TextVerticalMid

                                        End If

                                    ElseIf RafterTagLocation = "Right" Then

                                        If RafterAngD = 90 OrElse RafterAngD = 270 OrElse RafterAngD = -90 OrElse RafterAngD = -270 Then

                                            SortedRafterYPointsList = (From Ypnt In RafterPointsList Order By Ypnt.Y Select Ypnt).ToList
                                            RafterMPT = SortedRafterYPointsList(SortedRafterYPointsList.Count - 1)
                                            TagOffset = New Point3d(RafterMPT.X + 3.5 * Math.Cos(1.57079633), RafterMPT.Y + 3.5 * Math.Sin(1.57079633), RafterMPT.Z)
                                            TagAng = 1.57079633
                                            TagJus = AttachmentPoint.MiddleLeft
                                            TagHoriMode = TextHorizontalMode.TextLeft
                                            TagVertMode = TextVerticalMode.TextVerticalMid


                                        ElseIf RafterAngD = 0 OrElse RafterAngD = 180 OrElse RafterAngD = -180 OrElse RafterAngD = 360 Then

                                            SortedRafterXPointsList = (From Xpnt In RafterPointsList Order By Xpnt.X Select Xpnt).ToList
                                            RafterMPT = SortedRafterXPointsList(SortedRafterXPointsList.Count - 1)
                                            TagOffset = New Point3d(RafterMPT.X + 3.5 * Math.Cos(0), RafterMPT.Y + 3.5 * Math.Sin(0), RafterMPT.Z)
                                            TagAng = 0.0
                                            TagJus = AttachmentPoint.MiddleLeft
                                            TagHoriMode = TextHorizontalMode.TextLeft
                                            TagVertMode = TextVerticalMode.TextVerticalMid


                                        End If



                                    End If


                                End If


                                Dim RafterTag As String
                                Dim RafterTakeOffTag As String


                                If MemberSize = "2x6" Then

                                    RafterTag = hyp9 & "'"

                                Else

                                    RafterTag = MemberSize & "x" & hyp9 & "'"

                                End If

                                RafterTakeOffTag = Grade & "-" & MemberSize & "x" & hyp9 & "'"


                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)



                                '' Create a single-line text object
                                Dim acText As DBText = New DBText()
                                acText.SetDatabaseDefaults()
                                'acText.TextStyleName
                                acText.Height = 6
                                acText.TextString = RafterTag
                                acText.Layer = "S-FRM-TAGS"
                                acText.Rotation = TagAng
                                acText.WidthFactor = 0.85
                                acText.Oblique = 0.26179939
                                'acText.Justify = AttachmentPoint.MiddleMid
                                acText.Justify = TagJus
                                'acText.ColorIndex = 2
                                acText.IsMirroredInX = False
                                acText.IsMirroredInY = False
                                'acText.HorizontalMode = TextHorizontalMode.TextMid
                                acText.HorizontalMode = TagHoriMode
                                'acText.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText.VerticalMode = TagVertMode
                                acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)



                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = RafterTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = TagJus
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                'acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.HorizontalMode = TagHoriMode
                                'acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.VerticalMode = TagVertMode
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                'acBlkTblRec.AppendEntity(acText2)
                                ' acTrans.AddNewlyCreatedDBObject(acText2, True)
                                'dtids.Add(acText2.ObjectId)
                                'TakeoffGroup.Append(dtids)

                                If LabelType = "Tags" Then

                                    acBlkTblRec.AppendEntity(acText)
                                    acTrans.AddNewlyCreatedDBObject(acText, True)

                                Else

                                    acBlkTblRec.AppendEntity(acText)
                                    acTrans.AddNewlyCreatedDBObject(acText, True)
                                    dtids.Add(acText.ObjectId)
                                    acBlkTblRec.AppendEntity(acText2)
                                    acTrans.AddNewlyCreatedDBObject(acText2, True)
                                    dtids.Add(acText2.ObjectId)
                                    TakeoffGroup.Append(dtids)

                                End If


                            End If






                        Next

                        aced.WriteMessage(vbLf & "Selected Rafters have been tagged.")

                    Else

                        aced.WriteMessage(vbLf & "**Error** No Rafters have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub CJoistTag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line


            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select CJoist with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Ceiling Joist(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If

                    If (prSelRes.Status = PromptStatus.OK) Then

                        Dim acSSet As SelectionSet = prSelRes.Value

                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                Dim GrpDesc As String = "Group_" & LineID
                                Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                GroupDict.UpgradeOpen()
                                GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)

                                Dim spts As New Point3dCollection()
                                Dim Joistlen As Double = acEnt.Length
                                Dim Joistlen2 As Double
                                Dim Jlen3 As Double
                                Dim Jlen4 As Double
                                Dim Jlen5 As Double
                                Dim Jlen6 As Double
                                Dim Jlen7 As Double
                                Dim Jlen8 As Double
                                Dim Jlen9 As Double



                                If Joistlen >= 96 Then

                                    Joistlen2 = acEnt.Length / 2
                                    Jlen3 = Joistlen / 12
                                    Jlen4 = Jlen3 + 0.99999999
                                    Jlen5 = Math.Floor(Jlen4)
                                    Jlen6 = Math.Floor(Jlen5 + 1)
                                    Jlen7 = Jlen6 * 0.5
                                    Jlen8 = Math.Floor(Jlen7)
                                    Jlen9 = Jlen8 * 2

                                Else

                                    Joistlen2 = acEnt.Length / 2
                                    Jlen3 = Joistlen / 12
                                    Jlen4 = Jlen3
                                    Jlen5 = Math.Floor(Jlen4)
                                    Jlen6 = Math.Floor(Jlen5 + 1)
                                    'Jlen7 = Jlen6 * 0.5
                                    'Jlen8 = Math.Floor(Jlen7)
                                    Jlen9 = Jlen6


                                End If


                                Dim JoistSCLen As Double = Joistlen - 7
                                Dim JoistSP As Point3d = acEnt.StartPoint
                                Dim JoistEP As Point3d = acEnt.EndPoint
                                Dim JoistAng As Double = acEnt.Angle
                                Dim JoistAngD As Double = Math.Round((JoistAng * 180.0) / Math.PI, 0)
                                Dim JoistIptCol As New Point3dCollection()
                                JoistIptCol.Clear()
                                Dim JoistMPT As Point3d

                                If JoistAngD < 0 Then

                                    JoistAngD = 360 + JoistAngD

                                End If

                                Dim JoistAngRad As Double = JoistAngD * Math.PI / 180

                                If FenceTag = "Yes" Then

                                    acEnt.IntersectWith(acLine, Intersect.OnBothOperands, JoistIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If JoistIptCol.Count <> 0 Then

                                        JoistMPT = JoistIptCol.Item(0)
                                        'aced.WriteMessage(vbLf & "WallIptCol.Count <> 0 = " & WallIptCol.Count.ToString)

                                    End If

                                Else

                                    JoistMPT = New Point3d(JoistSP.X + Joistlen2 * Math.Cos(JoistAng), JoistSP.Y + Joistlen2 * Math.Sin(JoistAng), JoistSP.Z)

                                End If

                                Dim TagAng As Double
                                Dim TagOffset As Point3d

                                If JoistAngD = 90 OrElse JoistAngD = 270 OrElse JoistAngD = -90 OrElse JoistAngD = -270 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(3.14159265), JoistMPT.Y + 4.875 * Math.Sin(3.14159265), JoistMPT.Z)
                                    TagAng = 1.57079633

                                ElseIf JoistAngD = 0 OrElse JoistAngD = 180 OrElse JoistAngD = -180 OrElse JoistAngD = 360 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(1.57079633), JoistMPT.Y + 4.875 * Math.Sin(1.57079633), JoistMPT.Z)
                                    TagAng = 0.0

                                End If

                                Dim JoistTag As String
                                Dim JoistTakeOffTag As String

                                If MemberSizeButton = "Auto" Then

                                    If Grade = "#2" AndAlso Storage = "No Storage" Then

                                        GetMemberSize2NS(JoistSCLen)

                                    ElseIf Grade = "#3" AndAlso Storage = "No Storage" Then

                                        GetMemberSize3NS(JoistSCLen)

                                    ElseIf Grade = "#2" AndAlso Storage = "With Storage" Then

                                        GetMemberSize2WS(JoistSCLen)

                                    ElseIf Grade = "#3" AndAlso Storage = "With Storage" Then

                                        GetMemberSize3WS(JoistSCLen)

                                    End If

                                End If

                                If MemberSize = "2x6" Then

                                    JoistTag = Jlen9 & "'"
                                    JoistTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"

                                ElseIf MemberSize = "2x8" OrElse MemberSize = "2x10" OrElse MemberSize = "2x12" Then

                                    JoistTag = MemberSize & "x" & Jlen9 & "'"
                                    JoistTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"

                                Else

                                    JoistTag = "CHANGE SPACING"

                                End If

                                'If JoistTag <> "CHANGE SPACING" Then

                                JoistTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                '
                                'If

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                                '' Create a single-line text object
                                Dim acText As DBText = New DBText()
                                acText.SetDatabaseDefaults()
                                acText.Height = 6
                                acText.TextString = JoistTag
                                acText.Layer = "S-FRM-TAGS"
                                acText.Rotation = TagAng
                                acText.WidthFactor = 0.85
                                acText.Oblique = 0.26179939
                                acText.Justify = AttachmentPoint.MiddleMid
                                'acText.ColorIndex = 2
                                acText.IsMirroredInX = False
                                acText.IsMirroredInY = False
                                acText.HorizontalMode = TextHorizontalMode.TextMid
                                acText.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                ' acBlkTblRec.AppendEntity(acText)
                                'acTrans.AddNewlyCreatedDBObject(acText, True)
                                'dtids.Add(acText.ObjectId)

                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = JoistTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = AttachmentPoint.MiddleMid
                                'acText2.ColorIndex = 2
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                'acBlkTblRec.AppendEntity(acText2)
                                'acTrans.AddNewlyCreatedDBObject(acText2, True)

                                'dtids.Add(acText2.ObjectId)
                                'TakeoffGroup.Append(dtids)

                                If LabelType = "Tags" Then

                                    acBlkTblRec.AppendEntity(acText)
                                    acTrans.AddNewlyCreatedDBObject(acText, True)

                                Else

                                    acBlkTblRec.AppendEntity(acText)
                                    acTrans.AddNewlyCreatedDBObject(acText, True)


                                    If JoistTag <> "CHANGE SPACING" Then

                                        dtids.Add(acText.ObjectId)
                                        acBlkTblRec.AppendEntity(acText2)
                                        acTrans.AddNewlyCreatedDBObject(acText2, True)
                                        dtids.Add(acText2.ObjectId)
                                        TakeoffGroup.Append(dtids)

                                    End If



                                End If

                            End If

                        Next

                        aced.WriteMessage(vbLf & "Selected Ceiling Joists have been tagged.")

                    Else

                        aced.WriteMessage(vbLf & "**Error** No Ceiling Joists have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub PurlinTag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select Purlin with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-PURLIN")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Ceiling Purlin(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-PURLIN")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If

                    If (prSelRes.Status = PromptStatus.OK) Then

                        Dim acSSet As SelectionSet = prSelRes.Value

                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                'Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                'dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                'Dim GrpDesc As String = "Group_" & LineID
                                'Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                'Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                ' GroupDict.UpgradeOpen()
                                'GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                'acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)

                                Dim spts As New Point3dCollection()
                                Dim Purlinlen As Double = acEnt.Length
                                Dim Purlinlen2 As Double = acEnt.Length / 2
                                Dim Jlen3 As Double = Purlinlen / 12
                                Dim Jlen4 As Double = Jlen3 + 0.99999999
                                Dim Jlen5 As Double = Math.Floor(Jlen4)
                                Dim Jlen6 As Double = Math.Floor(Jlen5 + 1)
                                Dim Jlen7 As Double = Jlen6 * 0.5
                                Dim Jlen8 As Double = Math.Floor(Jlen7)
                                Dim Jlen9 As Double = Jlen8 * 2
                                Dim PurlinSCLen As Double = Purlinlen - 7
                                Dim PurlinSP As Point3d = acEnt.StartPoint
                                Dim PurlinEP As Point3d = acEnt.EndPoint
                                Dim PurlinAng As Double = acEnt.Angle
                                Dim PurlinAngD As Double = Math.Round((PurlinAng * 180.0) / Math.PI, 0)
                                Dim PurlinIptCol As New Point3dCollection()
                                PurlinIptCol.Clear()
                                Dim PurlinMPT As Point3d

                                If PurlinAngD < 0 Then

                                    PurlinAngD = 360 + PurlinAngD

                                End If

                                Dim PurlinAngRad As Double = PurlinAngD * Math.PI / 180

                                If FenceTag = "Yes" Then

                                    acEnt.IntersectWith(acLine, Intersect.OnBothOperands, PurlinIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If PurlinIptCol.Count <> 0 Then

                                        PurlinMPT = PurlinIptCol.Item(0)

                                    End If

                                Else

                                    PurlinMPT = New Point3d(PurlinSP.X + Purlinlen2 * Math.Cos(PurlinAng), PurlinSP.Y + Purlinlen2 * Math.Sin(PurlinAng), PurlinSP.Z)

                                End If

                                Dim TagAng As Double
                                Dim TagOffset As Point3d

                                If PurlinAngD = 90 OrElse PurlinAngD = 270 OrElse PurlinAngD = -90 OrElse PurlinAngD = -270 Then

                                    TagOffset = New Point3d(PurlinMPT.X + 4.875 * Math.Cos(3.14159265), PurlinMPT.Y + 4.875 * Math.Sin(3.14159265), PurlinMPT.Z)
                                    TagAng = 1.57079633

                                ElseIf PurlinAngD = 0 OrElse PurlinAngD = 180 OrElse PurlinAngD = -180 OrElse PurlinAngD = 360 Then

                                    TagOffset = New Point3d(PurlinMPT.X + 4.875 * Math.Cos(1.57079633), PurlinMPT.Y + 4.875 * Math.Sin(1.57079633), PurlinMPT.Z)
                                    TagAng = 0.0

                                End If

                                Dim PurlinTag As String
                                Dim PurlinTakeOffTag As String

                                If MemberSize = "2x6" Then

                                    PurlinTag = Jlen9 & "'"
                                    'PurlinTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                    'PurlinTakeOffTag = MemberSize & "x" & Jlen9 & "'" & " PURLIN"

                                Else

                                    PurlinTag = MemberSize & "x" & Jlen9 & "'"
                                    'PurlinTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                    'PurlinTakeOffTag = MemberSize & "x" & Jlen9 & "'" & " PURLIN"


                                End If

                                'PurlinTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                PurlinTakeOffTag = MemberSize & "x" & Jlen9 & "'" & " PURLIN"

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                                '' Create a single-line text object
                                'Dim acText As DBText = New DBText()
                                'acText.SetDatabaseDefaults()
                                'acText.Height = 6
                                'acText.TextString = PurlinTag
                                'acText.Layer = "S-FRM-TAGS"
                                'acText.Rotation = TagAng
                                'acText.WidthFactor = 0.85
                                'acText.Oblique = 0.26179939
                                'acText.Justify = AttachmentPoint.MiddleMid
                                'acText.IsMirroredInX = False
                                'acText.IsMirroredInY = False
                                'acText.HorizontalMode = TextHorizontalMode.TextMid
                                'acText.VerticalMode = TextVerticalMode.TextVerticalMid
                                'acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                'acBlkTblRec.AppendEntity(acText)
                                'acTrans.AddNewlyCreatedDBObject(acText, True)
                                'dtids.Add(acText.ObjectId)


                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = PurlinTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = AttachmentPoint.MiddleMid
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                acBlkTblRec.AppendEntity(acText2)
                                acTrans.AddNewlyCreatedDBObject(acText2, True)
                                'dtids.Add(acText2.ObjectId)


                                'TakeoffGroup.Append(dtids)

                            End If



                        Next

                        aced.WriteMessage(vbLf & "Selected Purlins have been tagged.")

                    Else


                        aced.WriteMessage(vbLf & "**Error** No Purlins have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub TJITag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line


            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select CJoist with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Ceiling Joist(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If

                    If (prSelRes.Status = PromptStatus.OK) Then

                        Dim acSSet As SelectionSet = prSelRes.Value

                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                'Dim GrpDesc As String = "Group_" & LineID
                                'Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                'Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                'GroupDict.UpgradeOpen()
                                'GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                'acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)

                                Dim spts As New Point3dCollection()
                                Dim Joistlen As Double = acEnt.Length
                                Dim Joistlen2 As Double
                                Dim Jlen3 As Double
                                Dim Jlen4 As Double
                                Dim Jlen5 As Double
                                Dim Jlen6 As Double
                                Dim Jlen7 As Double
                                Dim Jlen8 As Double
                                Dim Jlen9 As Double



                                'If Joistlen >= 96 Then

                                'Joistlen2 = acEnt.Length / 2
                                'Jlen3 = Joistlen / 12
                                'Jlen4 = Jlen3 + 0.99999999
                                'Jlen5 = Math.Floor(Jlen4)
                                'Jlen6 = Math.Floor(Jlen5 + 1)
                                'Jlen7 = Jlen6 * 0.5
                                'Jlen8 = Math.Floor(Jlen7)
                                'Jlen9 = Jlen8 * 2

                                'Else

                                Joistlen2 = acEnt.Length / 2
                                Jlen3 = Joistlen / 12
                                Jlen4 = Jlen3
                                Jlen5 = Math.Floor(Jlen4)
                                Jlen6 = Math.Floor(Jlen5 + 1)
                                'Jlen7 = Jlen6 * 0.5
                                'Jlen8 = Math.Floor(Jlen7)
                                Jlen9 = Jlen6


                                'End If


                                Dim JoistSCLen As Double = Joistlen - 7
                                Dim JoistSP As Point3d = acEnt.StartPoint
                                Dim JoistEP As Point3d = acEnt.EndPoint
                                Dim JoistAng As Double = acEnt.Angle
                                Dim JoistAngD As Double = Math.Round((JoistAng * 180.0) / Math.PI, 0)
                                Dim JoistIptCol As New Point3dCollection()
                                JoistIptCol.Clear()
                                Dim JoistMPT As Point3d

                                If JoistAngD < 0 Then

                                    JoistAngD = 360 + JoistAngD

                                End If

                                Dim JoistAngRad As Double = JoistAngD * Math.PI / 180

                                If FenceTag = "Yes" Then

                                    acEnt.IntersectWith(acLine, Intersect.OnBothOperands, JoistIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If JoistIptCol.Count <> 0 Then

                                        JoistMPT = JoistIptCol.Item(0)
                                        'aced.WriteMessage(vbLf & "WallIptCol.Count <> 0 = " & WallIptCol.Count.ToString)

                                    End If

                                Else

                                    JoistMPT = New Point3d(JoistSP.X + Joistlen2 * Math.Cos(JoistAng), JoistSP.Y + Joistlen2 * Math.Sin(JoistAng), JoistSP.Z)

                                End If

                                Dim TagAng As Double
                                Dim TagOffset As Point3d

                                If JoistAngD = 90 OrElse JoistAngD = 270 OrElse JoistAngD = -90 OrElse JoistAngD = -270 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(3.14159265), JoistMPT.Y + 4.875 * Math.Sin(3.14159265), JoistMPT.Z)
                                    TagAng = 1.57079633

                                ElseIf JoistAngD = 0 OrElse JoistAngD = 180 OrElse JoistAngD = -180 OrElse JoistAngD = 360 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(1.57079633), JoistMPT.Y + 4.875 * Math.Sin(1.57079633), JoistMPT.Z)
                                    TagAng = 0.0

                                End If

                                'Dim JoistTag As String
                                Dim JoistTakeOffTag As String

                                'JoistTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                JoistTakeOffTag = Jlen9 & "'-" & IJoistSize & " TJI " & IJoistSeries & "s"

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                                '' Create a single-line text object
                                ' Dim acText As DBText = New DBText()
                                'acText.SetDatabaseDefaults()
                                'acText.Height = 6
                                'acText.TextString = JoistTag
                                'acText.Layer = "S-FRM-TAGS"
                                'acText.Rotation = TagAng
                                'acText.WidthFactor = 0.85
                                'acText.Oblique = 0.26179939
                                'acText.Justify = AttachmentPoint.MiddleMid
                                'acText.IsMirroredInX = False
                                'acText.IsMirroredInY = False
                                'acText.HorizontalMode = TextHorizontalMode.TextMid
                                'acText.VerticalMode = TextVerticalMode.TextVerticalMid
                                'acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                'acBlkTblRec.AppendEntity(acText)
                                'acTrans.AddNewlyCreatedDBObject(acText, True)

                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = JoistTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = AttachmentPoint.MiddleMid
                                'acText2.ColorIndex = 2
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                acBlkTblRec.AppendEntity(acText2)
                                acTrans.AddNewlyCreatedDBObject(acText2, True)


                            End If

                        Next

                        aced.WriteMessage(vbLf & "Selected Ceiling Joists have been tagged.")

                    Else

                        aced.WriteMessage(vbLf & "**Error** No Ceiling Joists have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub BCITag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select CJoist with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Ceiling Joist(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-CJOIST")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If

                    If (prSelRes.Status = PromptStatus.OK) Then

                        Dim acSSet As SelectionSet = prSelRes.Value

                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                'Dim GrpDesc As String = "Group_" & LineID
                                'Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                'Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                'GroupDict.UpgradeOpen()
                                'GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                'acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)

                                Dim spts As New Point3dCollection()
                                Dim Joistlen As Double = acEnt.Length
                                Dim Joistlen2 As Double
                                Dim Jlen3 As Double
                                Dim Jlen4 As Double
                                Dim Jlen5 As Double
                                Dim Jlen6 As Double
                                Dim Jlen7 As Double
                                Dim Jlen8 As Double
                                Dim Jlen9 As Double

                                'If Joistlen >= 96 Then

                                'Joistlen2 = acEnt.Length / 2
                                'Jlen3 = Joistlen / 12
                                'Jlen4 = Jlen3 + 0.99999999
                                'Jlen5 = Math.Floor(Jlen4)
                                'Jlen6 = Math.Floor(Jlen5 + 1)
                                'Jlen7 = Jlen6 * 0.5
                                'Jlen8 = Math.Floor(Jlen7)
                                'Jlen9 = Jlen8 * 2

                                'Else

                                Joistlen2 = acEnt.Length / 2
                                Jlen3 = Joistlen / 12
                                Jlen4 = Jlen3
                                Jlen5 = Math.Floor(Jlen4)
                                Jlen6 = Math.Floor(Jlen5 + 1)
                                'Jlen7 = Jlen6 * 0.5
                                'Jlen8 = Math.Floor(Jlen7)
                                Jlen9 = Jlen6

                                'End If

                                Dim JoistSCLen As Double = Joistlen - 7
                                Dim JoistSP As Point3d = acEnt.StartPoint
                                Dim JoistEP As Point3d = acEnt.EndPoint
                                Dim JoistAng As Double = acEnt.Angle
                                Dim JoistAngD As Double = Math.Round((JoistAng * 180.0) / Math.PI, 0)
                                Dim JoistIptCol As New Point3dCollection()
                                JoistIptCol.Clear()
                                Dim JoistMPT As Point3d

                                If JoistAngD < 0 Then

                                    JoistAngD = 360 + JoistAngD

                                End If

                                Dim JoistAngRad As Double = JoistAngD * Math.PI / 180

                                If FenceTag = "Yes" Then

                                    acEnt.IntersectWith(acLine, Intersect.OnBothOperands, JoistIptCol, IntPtr.Zero, IntPtr.Zero)

                                    If JoistIptCol.Count <> 0 Then

                                        JoistMPT = JoistIptCol.Item(0)
                                        'aced.WriteMessage(vbLf & "WallIptCol.Count <> 0 = " & WallIptCol.Count.ToString)

                                    End If

                                Else

                                    JoistMPT = New Point3d(JoistSP.X + Joistlen2 * Math.Cos(JoistAng), JoistSP.Y + Joistlen2 * Math.Sin(JoistAng), JoistSP.Z)

                                End If

                                Dim TagAng As Double
                                Dim TagOffset As Point3d

                                If JoistAngD = 90 OrElse JoistAngD = 270 OrElse JoistAngD = -90 OrElse JoistAngD = -270 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(3.14159265), JoistMPT.Y + 4.875 * Math.Sin(3.14159265), JoistMPT.Z)
                                    TagAng = 1.57079633

                                ElseIf JoistAngD = 0 OrElse JoistAngD = 180 OrElse JoistAngD = -180 OrElse JoistAngD = 360 Then

                                    TagOffset = New Point3d(JoistMPT.X + 4.875 * Math.Cos(1.57079633), JoistMPT.Y + 4.875 * Math.Sin(1.57079633), JoistMPT.Z)
                                    TagAng = 0.0

                                End If

                                'Dim JoistTag As String
                                Dim JoistTakeOffTag As String

                                'JoistTakeOffTag = Grade & "-" & MemberSize & "x" & Jlen9 & "'"
                                JoistTakeOffTag = Jlen9 & "'-" & IJoistSize & " BCI " & IJoistSeries & "s"

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                                '' Create a single-line text object
                                ' Dim acText As DBText = New DBText()
                                'acText.SetDatabaseDefaults()
                                'acText.Height = 6
                                'acText.TextString = JoistTag
                                'acText.Layer = "S-FRM-TAGS"
                                'acText.Rotation = TagAng
                                'acText.WidthFactor = 0.85
                                'acText.Oblique = 0.26179939
                                'acText.Justify = AttachmentPoint.MiddleMid
                                'acText.IsMirroredInX = False
                                'acText.IsMirroredInY = False
                                'acText.HorizontalMode = TextHorizontalMode.TextMid
                                'acText.VerticalMode = TextVerticalMode.TextVerticalMid
                                'acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                'acBlkTblRec.AppendEntity(acText)
                                'acTrans.AddNewlyCreatedDBObject(acText, True)

                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = JoistTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = AttachmentPoint.MiddleMid
                                'acText2.ColorIndex = 2
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                acBlkTblRec.AppendEntity(acText2)
                                acTrans.AddNewlyCreatedDBObject(acText2, True)


                            End If

                        Next

                        aced.WriteMessage(vbLf & "Selected Ceiling Joists have been tagged.")

                    Else

                        aced.WriteMessage(vbLf & "**Error** No Ceiling Joists have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub


        Private Sub ClipTag()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim ech As Integer = Application.GetSystemVariable("cmdecho")
            'Dim opts As New PromptSelectionOptions()
            Dim prSelRes As PromptSelectionResult
            Dim acLine As Line

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    'aced.WriteMessage(vbLf & "FenceTag = " & FenceTag.ToString)

                    If FenceTag = "Yes" Then

                        Dim pts As New Point3dCollection()
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        aced.WriteMessage(vbLf & "Select Rafter with a crossing fenceline...")

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
                        'acLine.Linetype = "Hidden"
                        Dim FLineAng As Double = acLine.Angle
                        Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI
                        Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(0, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-RAFTER")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.SelectFence(pts, acSelFtr)

                    Else

                        Dim opts As New PromptSelectionOptions()
                        opts.MessageForAdding = vbLf & "Select Rafter(s) to tag: ..."
                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LINE"), New TypedValue(DxfCode.LayerName, "S-FRM-RAFTER")}
                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                        prSelRes = acDoc.Editor.GetSelection(opts, acSelFtr)

                    End If


                    If (prSelRes.Status = PromptStatus.OK) Then

                        ' There are selected entities
                        ' Put your command using pickfirst set code here
                        Dim acSSet As SelectionSet = prSelRes.Value


                        '' Step through the objects in the selection set
                        For Each SSObj As SelectedObject In acSSet

                            Dim typ As String = SSObj.ObjectId.ObjectClass.Name()

                            If typ = "AcDbLine" Then

                                Dim dtids As ObjectIdCollection = New ObjectIdCollection()
                                dtids.Clear()

                                Dim acEnt As Line = acTrans.GetObject(SSObj.ObjectId, OpenMode.ForRead, False, True)
                                Dim lay As String = acEnt.Layer
                                Dim col As String = acEnt.ColorIndex
                                Dim ID As String = acEnt.ObjectId.ToString
                                Dim LineID As String = ID
                                Dim GrpDesc As String = "Group_" & LineID
                                Dim GroupDict As DBDictionary = CType(acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForRead), DBDictionary)
                                Dim TakeoffGroup As Group = New Group(GrpDesc, True)
                                GroupDict.UpgradeOpen()
                                GroupDict.SetAt(TakeoffGroup.Description, TakeoffGroup)
                                acTrans.AddNewlyCreatedDBObject(TakeoffGroup, True)


                                Dim Slopepart As String() = RafterSlope.Split(":")
                                Dim Slopepart1 As Integer = Slopepart(0)
                                Dim Slopepart2 As Integer = Slopepart(1)
                                Dim Slope As Double = Slopepart1 / Slopepart2

                                Dim spts As New Point3dCollection()

                                Dim Rafterlen As Double = acEnt.Length

                                Dim Rafterlen2 As Double = acEnt.Length / 2
                                Dim rise As Double = Rafterlen * Slope
                                Dim hypt As Double = Math.Sqrt((rise * rise) + (Rafterlen * Rafterlen))
                                Dim HypotLen As Double = hypt + (Slopepart1 * 0.5)

                                Dim hyp3 As Double
                                Dim hyp4 As Double
                                Dim hyp5 As Double
                                Dim hyp6 As Double
                                Dim hyp7 As Double
                                Dim hyp8 As Double
                                Dim hyp9 As Double


                                If HypotLen >= 96 Then

                                    hyp3 = HypotLen / 12
                                    hyp4 = hyp3 + 0.99999999
                                    hyp5 = Math.Floor(hyp4)
                                    hyp6 = Math.Floor(hyp5 + 1)
                                    hyp7 = hyp6 * 0.5
                                    hyp8 = Math.Floor(hyp7)
                                    hyp9 = hyp8 * 2

                                Else

                                    hyp3 = HypotLen / 12
                                    hyp4 = hyp3
                                    hyp5 = Math.Floor(hyp4)
                                    hyp6 = Math.Floor(hyp5 + 1)
                                    'hyp7 = hyp6 * 0.5
                                    'hyp8 = Math.Floor(hyp7)
                                    hyp9 = hyp6

                                End If

                                Dim RafterSP As Point3d = acEnt.StartPoint
                                Dim RafterEP As Point3d = acEnt.EndPoint
                                Dim RafterAng As Double = acEnt.Angle
                                Dim RafterAngD As Double = Math.Round((RafterAng * 180.0) / Math.PI, 0)
                                Dim RafterIptCol As New Point3dCollection()
                                RafterIptCol.Clear()
                                Dim RafterMPT As Point3d

                                Dim RafterPointsList As New List(Of Point3d)
                                Dim SortedRafterPointsList As New List(Of Point3d)
                                Dim RafterStartPointsList As New List(Of Point3d)
                                Dim SortedRafterPoints As New List(Of Point3d)
                                Dim SortedRafterXPointsList As New List(Of Point3d)
                                Dim SortedRafterYPointsList As New List(Of Point3d)
                                RafterPointsList.Add(RafterSP)
                                RafterPointsList.Add(RafterEP)

                                If RafterAngD < 0 Then

                                    RafterAngD = 360 + RafterAngD

                                End If

                                Dim RafterAngRad As Double = RafterAngD * Math.PI / 180
                                Dim TagAng As Double
                                Dim TagOffset As Point3d
                                Dim TagJus As AttachmentPoint = AttachmentPoint.MiddleMid
                                Dim TagHoriMode As TextHorizontalMode = TextHorizontalMode.TextMid
                                Dim TagVertMode As TextVerticalMode = TextVerticalMode.TextVerticalMid

                                RafterMPT = New Point3d(RafterSP.X + Rafterlen2 * Math.Cos(RafterAng), RafterSP.Y + Rafterlen2 * Math.Sin(RafterAng), RafterSP.Z)
                                If RafterAngD = 90 OrElse RafterAngD = 270 OrElse RafterAngD = -90 OrElse RafterAngD = -270 Then

                                    TagOffset = New Point3d(RafterMPT.X + 4.875 * Math.Cos(3.14159265), RafterMPT.Y + 4.875 * Math.Sin(3.14159265), RafterMPT.Z)
                                    TagAng = 1.57079633

                                ElseIf RafterAngD = 0 OrElse RafterAngD = 180 OrElse RafterAngD = -180 OrElse RafterAngD = 360 Then

                                    TagOffset = New Point3d(RafterMPT.X + 4.875 * Math.Cos(1.57079633), RafterMPT.Y + 4.875 * Math.Sin(1.57079633), RafterMPT.Z)
                                    TagAng = 0.0

                                End If

                                Dim RafterTakeOffTag As String

                                RafterTakeOffTag = "H2.5 CLIPS"

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                                '' Create a single-line text object
                                'Dim acText As DBText = New DBText()
                                'acText.SetDatabaseDefaults()
                                'acText.Height = 6
                                'acText.TextString = RafterTag
                                'acText.Layer = "S-FRM-TAGS"
                                'acText.Rotation = TagAng
                                'acText.WidthFactor = 0.85
                                'acText.Oblique = 0.26179939
                                'acText.Justify = TagJus
                                'acText.IsMirroredInX = False
                                'acText.IsMirroredInY = False
                                'acText.HorizontalMode = TagHoriMode
                                'acText.VerticalMode = TagVertMode
                                'acText.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)
                                'acBlkTblRec.AppendEntity(acText)
                                'acTrans.AddNewlyCreatedDBObject(acText, True)


                                '' Create a single-line text object
                                Dim acText2 As DBText = New DBText()
                                acText2.SetDatabaseDefaults()
                                acText2.Height = 6
                                acText2.TextString = RafterTakeOffTag
                                acText2.Layer = "S-FRM-TAKEOFFS"
                                acText2.Rotation = TagAng
                                acText2.WidthFactor = 0.85
                                acText2.Oblique = 0.26179939
                                acText2.Justify = TagJus
                                acText2.IsMirroredInX = False
                                acText2.IsMirroredInY = False
                                'acText2.HorizontalMode = TextHorizontalMode.TextMid
                                acText2.HorizontalMode = TagHoriMode
                                'acText2.VerticalMode = TextVerticalMode.TextVerticalMid
                                acText2.VerticalMode = TagVertMode
                                acText2.AlignmentPoint = New Point3d(TagOffset.X, TagOffset.Y, 0)

                                acBlkTblRec.AppendEntity(acText2)
                                acTrans.AddNewlyCreatedDBObject(acText2, True)

                            End If

                        Next

                        aced.WriteMessage(vbLf & "Selected Rafters have been tagged.")

                    Else

                        aced.WriteMessage(vbLf & "**Error** No Rafters have been selected! Please try again.")

                    End If

                    Application.SetSystemVariable("cmdecho", ech)

                    ' Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using



        End Sub


        Private Sub AddLayer()

            Dim Layerprop(,) As String = New String(,) {{"S-FRM-TAGS", "1", "Continuous"},
                                                        {"S-FRM-TAKEOFFS", "251", "Continuous"}}

            Dim Rowcount As Integer = Layerprop.GetUpperBound(0)

            '' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Open the Layer table for read
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                For row = 0 To Rowcount

                    Dim sLayerName As String = Layerprop(row, 0)
                    Dim LayColor As Integer = Layerprop(row, 1)
                    Dim LayLinetype As String = Layerprop(row, 2)

                    If acLyrTbl.Has(sLayerName) = False Then

                        Using acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                            '' Assign the layer a name
                            acLyrTblRec.Name = sLayerName

                            '' Assign the layer the ACI color 1 and a name
                            acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, LayColor)

                            '' Upgrade the Layer table for write
                            acLyrTbl.UpgradeOpen()

                            '' Append the new layer to the Layer table and the transaction
                            acLyrTbl.Add(acLyrTblRec)
                            acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                            '' Open the Layer table for read
                            Dim acLinTbl As LinetypeTable
                            acLinTbl = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)

                            If acLinTbl.Has(LayLinetype) = True Then
                                '' Upgrade the Layer Table Record for write
                                acLyrTblRec.UpgradeOpen()

                                '' Set the linetype for the layer
                                acLyrTblRec.LinetypeObjectId = acLinTbl(LayLinetype)

                            Else

                                '' Load the Center Linetype
                                acCurDb.LoadLineTypeFile(LayLinetype, "acad.lin")

                                '' Upgrade the Layer Table Record for write
                                acLyrTblRec.UpgradeOpen()

                                '' Set the linetype for the layer
                                acLyrTblRec.LinetypeObjectId = acLinTbl(LayLinetype)

                            End If
                        End Using

                    End If

                Next

                '' Save the changes and dispose of the transaction
                acTrans.Commit()
            End Using
        End Sub

        Private Sub SetTextStyle()

            '' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    '' Open the Block table for read
                    Dim acBlkTbl As BlockTable
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                    '' Open the Block table record Model space for write
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    Dim tsTbl As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)

                    Dim textstylename As String = "Lengths"

                    If tsTbl.Has(textstylename) Then

                        Dim TxtStyleVar As Object
                        'Get the TextStylye SYSTEM VARIABLE
                        TxtStyleVar = Application.GetSystemVariable("TextStyle")
                        'SET TextStyle SYSTEM VARIABLE
                        Application.SetSystemVariable("TextStyle", "Lengths")

                    Else

                        Dim st As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                        Dim strr As TextStyleTableRecord = New TextStyleTableRecord()
                        strr.Name = "Lengths"
                        st.Add(strr)
                        strr.FileName = "simplex.shx"
                        strr.ObliquingAngle = 0.26179939
                        strr.XScale = 0.85
                        strr.TextSize = 6.0
                        strr.IsVertical = False
                        strr.IsShapeFile = False
                        acTrans.AddNewlyCreatedDBObject(strr, True)

                        'make as current
                        Dim TxtStyleVar As Object
                        'Get the TextStylye SYSTEM VARIABLE
                        TxtStyleVar = Application.GetSystemVariable("TextStyle")
                        'SET TextStyle SYSTEM VARIABLE
                        Application.SetSystemVariable("TextStyle", "Lengths")

                    End If


                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub GetMemberSize2NS(JoistSCLen As Double)

            If JoistSpace = "24" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 251 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 251 AndAlso JoistSCLen > 211 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 211 AndAlso JoistSCLen > 167 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 167 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "19.2" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 281 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 281 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 187 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 187 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "16" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12.8" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "9.6" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "8" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If



            End If

        End Sub

        Private Sub GetMemberSize3NS(JoistSCLen As Double)

            If JoistSpace = "24" Then

                If JoistSCLen <= 229 AndAlso JoistSCLen > 193 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 193 AndAlso JoistSCLen > 159 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 159 AndAlso JoistSCLen > 126 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 126 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "19.2" Then

                If JoistSCLen <= 256 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 178 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 178 AndAlso JoistSCLen > 141 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 141 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "16" Then

                If JoistSCLen <= 280 AndAlso JoistSCLen > 237 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 237 AndAlso JoistSCLen > 195 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 195 AndAlso JoistSCLen > 155 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 155 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12.8" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 265 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 265 AndAlso JoistSCLen > 218 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 218 AndAlso JoistSCLen > 173 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 173 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 273 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 273 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 179 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 179 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "9.6" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "8" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            End If

        End Sub

        Private Sub GetMemberSize2WS(JoistSCLen As Double)

            If JoistSpace = "24" Then

                If JoistSCLen <= 209 AndAlso JoistSCLen > 177 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 177 AndAlso JoistSCLen > 150 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 150 AndAlso JoistSCLen > 118 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 118 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "19.2" Then

                If JoistSCLen <= 234 AndAlso JoistSCLen > 198 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 198 AndAlso JoistSCLen > 167 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 167 AndAlso JoistSCLen > 132 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 132 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "16" Then

                If JoistSCLen <= 256 AndAlso JoistSCLen > 217 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 217 AndAlso JoistSCLen > 183 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 183 AndAlso JoistSCLen > 144 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 144 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12.8" Then

                If JoistSCLen <= 286 AndAlso JoistSCLen > 243 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 243 AndAlso JoistSCLen > 205 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 205 AndAlso JoistSCLen > 162 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 162 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 251 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 251 AndAlso JoistSCLen > 211 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 211 AndAlso JoistSCLen > 167 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 167 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "9.6" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 281 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 281 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 187 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 187 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "8" Then

                If JoistSCLen <= 288 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 192 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 192 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            End If

        End Sub

        Private Sub GetMemberSize3WS(JoistSCLen As Double)

            If JoistSpace = "24" Then

                If JoistSCLen <= 162 AndAlso JoistSCLen > 137 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 137 AndAlso JoistSCLen > 113 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 113 AndAlso JoistSCLen > 89 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 89 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "19.2" Then

                If JoistSCLen <= 181 AndAlso JoistSCLen > 153 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 153 AndAlso JoistSCLen > 126 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 126 AndAlso JoistSCLen > 100 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 100 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "16" Then

                If JoistSCLen <= 198 AndAlso JoistSCLen > 167 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 167 AndAlso JoistSCLen > 138 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 138 AndAlso JoistSCLen > 110 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 110 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12.8" Then

                If JoistSCLen <= 222 AndAlso JoistSCLen > 187 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 187 AndAlso JoistSCLen > 154 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 154 AndAlso JoistSCLen > 122 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 122 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "12" Then

                If JoistSCLen <= 229 AndAlso JoistSCLen > 193 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 193 AndAlso JoistSCLen > 159 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 159 AndAlso JoistSCLen > 126 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 126 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "9.6" Then

                If JoistSCLen <= 256 AndAlso JoistSCLen > 216 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 216 AndAlso JoistSCLen > 178 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 178 AndAlso JoistSCLen > 141 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 141 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            ElseIf JoistSpace = "8" Then

                If JoistSCLen <= 280 AndAlso JoistSCLen > 237 Then

                    Module_TakeOffTag.MemberSize = "2x12"

                ElseIf JoistSCLen <= 237 AndAlso JoistSCLen > 195 Then

                    Module_TakeOffTag.MemberSize = "2x10"

                ElseIf JoistSCLen <= 195 AndAlso JoistSCLen > 155 Then

                    Module_TakeOffTag.MemberSize = "2x8"

                ElseIf JoistSCLen <= 155 Then

                    Module_TakeOffTag.MemberSize = "2x6"

                Else

                    Module_TakeOffTag.MemberSize = "CHANGE SPACING"

                End If

            End If

        End Sub


        Private Sub ConvertToFraction(Rafterlen As Double, Slope As Double, Slopepart1 As Double)

            Dim RafterLenFTIN As Integer = Rafterlen / 12
            Dim RafterLenFT As Integer = Math.Floor(RafterLenFTIN)
            Dim RafterLenINft As Integer = RafterLenFTIN - RafterLenFT
            Dim RafterLenINin As Integer = RafterLenINft * 12
            Dim RafterLenIN As Integer = Math.Floor(RafterLenINin)
            Dim RafterLenFracIN As Integer = (RafterLenINin - RafterLenIN) * 16

            Dim RunLen As String = "0"

            Dim RunLenFTIN As Double = Rafterlen / 12
            Dim RunLenFT As Integer = Math.Floor(RunLenFTIN)
            Dim RunLenINft As Double = RunLenFTIN - RunLenFT
            Dim RunLenINin As Double = RunLenINft * 12
            Dim RunLenIN As Integer = Math.Floor(RunLenINin)
            Dim RunLenFracIN As Double = (RunLenINin - RunLenIN) * 16
            Dim RunLenFracDC As Integer = Math.Round(RunLenFracIN)
            Dim prc As Integer = 16

            If RunLenFracDC = 1 Then

                RunLenIN = RunLenIN + 1
                RunLenFracDC = 0

                RunLen = RunLenFT & "'-" & RunLenIN & """"

            End If

            If RunLenFracDC <> 0 Then

                While RunLenFracDC Mod 2 = 0

                    Dim bfracmod = RunLenFracDC Mod 2
                    Dim RunLENMOD = RunLenFracDC Mod 2
                    RunLenFracDC = RunLenFracDC / 2
                    prc = prc / 2

                End While

                Dim RunLenfrac As String = RunLenFracDC & "/" & prc
                RunLen = RunLenFT & "'-" & RunLenIN & " " & RunLenfrac & """"


            End If





            Dim rise As Double = Rafterlen * Slope
            Dim hypt As Double = Math.Sqrt((rise * rise) + (Rafterlen * Rafterlen))
            Dim HypotLen As Double = hypt + (Slopepart1 * 0.5)

            Dim hyp3 As Double = HypotLen / 12
            Dim hyp4 As Double = hyp3 + 0.99999999
            Dim hyp5 As Double = Math.Floor(hyp4)
            Dim hyp6 As Double = Math.Floor(hyp5 + 1)
            Dim hyp7 As Double = hyp6 * 0.5
            Dim hyp8 As Double = Math.Floor(hyp7)
            Dim hyp9 As Double = hyp8 * 2

            Dim HypLen As String = "0"
            Dim HypLenFTIN As Double = HypotLen / 12
            Dim HypLenFT As Integer = Math.Floor(HypLenFTIN)
            Dim HypLenINft As Double = HypLenFTIN - HypLenFT
            Dim HypLenINin As Double = HypLenINft * 12
            Dim HypLenIN As Integer = Math.Floor(HypLenINin)
            Dim HypLenFracIN As Double = (HypLenINin - HypLenIN) * 16
            Dim HypLenFracDC As Integer = Math.Round(HypLenFracIN)

            If HypLenFracDC <> 0 Then

                While HypLenFracDC Mod 2 = 0

                    Dim bfracmod = HypLenFracDC Mod 2
                    Dim HypLENMOD = HypLenFracDC Mod 2

                    HypLenFracDC = HypLenFracDC / 2

                    prc = prc / 2

                End While

                Dim HypLenfrac As String = HypLenFracDC & "/" & prc

                HypLen = HypLenFT & "'-" & HypLenIN & " " & HypLenfrac

            End If

            'aced.WriteMessage(vbLf & "** Run = " & RunLen & ", Pitch = " & RafterSlope & ", Actual rafter length needed = " & HypLen & " **")

        End Sub

    End Class

End Namespace