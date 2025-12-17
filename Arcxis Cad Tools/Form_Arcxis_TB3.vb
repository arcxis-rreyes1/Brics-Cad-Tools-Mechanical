Imports System
Imports System.Drawing.Printing
Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices

Public Class Form_Arcxis_TB3

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor

        If Trim(TextBox1.Text).Length = 0 Then

            TextBox1.BackColor = System.Drawing.Color.Red

        ElseIf CheckBox1.Checked = True AndAlso Trim(TextBox2.Text).Length = 0 Then

            TextBox2.BackColor = System.Drawing.Color.Red

        Else

            Dim ElevNum As Integer = TextBox1.Text
            Dim OptNum As Integer = TextBox2.Text
            Dim NewLayoutName As String = ComboBox1.Text
            Dim panover As Integer
            Dim pandown As Integer
            Dim ExistlayoutName As String = ATB_FirLayoutName
            Dim ExistlayoutName2 As String = ATB_SecLayoutName
            Dim NextlayoutName As String
            Dim frm As New Form_Arcxis_TB1
            Dim pgltr As Char
            Dim elnt As String
            Dim ExistOptlayList As New List(Of String)
            ExistOptlayList.Clear()
            Dim OptlayList As New List(Of String)
            OptlayList.Clear()
            Dim ExistOptLayListCount As Integer
            Dim OptLayListCount As Integer

            If ExistlayoutName <> NewLayoutName Then

                RenameFramingLayout(ExistlayoutName, NewLayoutName)

            End If

            Dim i As Integer = Asc("A")
            Dim x As Char = Chr(i + 48)

            If NewLayoutName = "A (R)" Then

                elnt = "RIGHT"

                For pagenum As Integer = 0 To ElevNum - 1

                    pandown = 0
                    pgltr = Chr(65 + pagenum)
                    panover = 1800 * pagenum
                    NextlayoutName = pgltr & " (R)"

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                    If CheckBox1.Checked = True Then

                        NextlayoutName = pgltr & " opts. (R)"

                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        pandown = 1800
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                    End If

                Next

            ElseIf NewLayoutName = "F-1A (R)" Then

                elnt = "RIGHT"

                For pagenum As Integer = 0 To ElevNum - 1

                    pandown = 0
                    pgltr = Chr(65 + pagenum)
                    panover = 1800 * pagenum
                    NextlayoutName = "F-" & (pagenum + 1) & pgltr & " (R)"

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                    If CheckBox1.Checked = True Then

                        NextlayoutName = "F-" & (pagenum + 1) & pgltr & " opts. (R)"
                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        pandown = 1800
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                    End If

                Next

            ElseIf NewLayoutName = "A1(R)" Then

                elnt = "RIGHT"

                For pagenum As Integer = 0 To ElevNum - 1

                    pandown = 0
                    pgltr = Chr(65 + pagenum)
                    panover = 1800 * pagenum
                    NextlayoutName = pgltr & "1" & "(R)"
                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                    If CheckBox1.Checked = True Then

                        For optpagenum As Integer = 1 To OptNum

                            NextlayoutName = pgltr & "1" & "(R)" & "OP" & optpagenum
                            CreateFramingLayout(NewLayoutName, NextlayoutName)
                            pandown = 1800 * optpagenum
                            ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                        Next

                    End If

                Next

            Else

                If ATB_LayoutList.Count > 1 Then

                    For Each lay In ATB_LayoutList

                        If lay.Contains("OPT") Then

                            ExistOptlayList.Add(lay)

                        End If

                    Next

                    For Each lay In ExistOptlayList

                        'aced.WriteMessage(vbLf & "lay = " & lay.ToString())

                    Next

                    ExistOptLayListCount = ExistOptlayList.Count

                    If ExistOptLayListCount = 1 AndAlso OptNum = 1 Then

                        OptlayList.Add(ExistOptlayList(0))

                    ElseIf ExistOptLayListCount = 1 AndAlso OptNum > 1 Then

                        Dim optnam As String = ExistOptlayList(0)

                        If optnam.Contains("OPT 1") Then

                            OptlayList.Add(ExistOptlayList(0))

                            For opn As Integer = 2 To OptNum

                                elnt = "RIGHT"
                                pandown = 1800 * (opn - 1)
                                NextlayoutName = ATB_LayoutList(0) & " " & "OPT" & opn
                                CreateFramingLayout(NewLayoutName, NextlayoutName)
                                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                            Next

                        Else

                            Dim newoptnam As String = ExistOptlayList(0) & " 1"
                            OptlayList.Add(newoptnam)
                            ExistlayoutName = ExistOptlayList(0)
                            NewLayoutName = newoptnam
                            RenameFramingLayout(ExistlayoutName, NewLayoutName)
                            elnt = "RIGHT"

                            For opn As Integer = 2 To OptNum

                                pandown = 1800 * (opn - 1)
                                NextlayoutName = ExistOptlayList(0) & " " & opn
                                CreateFramingLayout(NewLayoutName, NextlayoutName)
                                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                            Next

                        End If

                    ElseIf ExistOptLayListCount = OptNum Then

                        For Each lay In ExistOptlayList

                            OptlayList.Add(lay)

                        Next

                    End If

                End If

                OptLayListCount = OptlayList.Count

                For Each lay In OptlayList

                    'aced.WriteMessage(vbLf & "lay =" & lay.ToString())

                Next

                Dim optlaynam As String = OptlayList(0)
                Dim optlaynamsplit As String()
                optlaynamsplit = optlaynam.Split(")")
                Dim olsp As String = optlaynamsplit(0)
                Dim olsp2 As String = optlaynamsplit(1)

                Dim laynamsplit As String()
                Dim FirstLayoutName As String = ComboBox1.Text

                If FirstLayoutName.Contains("(") Then

                    laynamsplit = FirstLayoutName.Split("(")
                    Dim elsp As String = laynamsplit(0)
                    Dim elsp2 As String = laynamsplit(1)
                    Dim LL As Integer = Asc(elsp)
                    Dim NexLL As Char = Chr(LL)

                    elnt = "RIGHT"

                    For pagenum As Integer = 1 To ElevNum - 1

                        Dim LLT As Integer = Asc(NexLL)
                        NexLL = Chr(LLT + 1)
                        pandown = 0
                        NewLayoutName = ATB_LayoutList(0)

                        panover = 1800 * pagenum
                        NextlayoutName = NexLL & " (" & elsp2
                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                        If CheckBox1.Checked = True Then

                            For optpagenum As Integer = 1 To OptNum

                                If OptLayListCount = 1 AndAlso OptNum = 1 Then

                                    pandown = 0
                                    NextlayoutName = NexLL & " (R)" & olsp2
                                    NewLayoutName = OptlayList(0)

                                ElseIf OptLayListCount = 1 AndAlso OptNum > 1 Then

                                    pandown = 1800 * (optpagenum - 1)
                                    NewLayoutName = OptlayList(0)
                                    NextlayoutName = NexLL & " (R) OPT " & optpagenum


                                ElseIf OptLayListCount = OptNum Then

                                    pandown = 0
                                    NextlayoutName = NexLL & " (R) OPT " & optpagenum
                                    NewLayoutName = OptlayList(optpagenum - 1)

                                End If

                                CreateFramingLayout(NewLayoutName, NextlayoutName)

                                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt)

                            Next

                        End If

                    Next

                End If

            End If

            SetPSLTScale()

            Me.Close()

        End If

    End Sub

    Private Sub CreateFramingLayout(NewLayoutName As String, NextlayoutName As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Get the current document and database
        Dim acDocex As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDbex As Database = acDoc.Database
        Dim aced As Editor = acDocex.Editor

        ' Create a transaction for the external drawing
        Using acTransEx As Transaction = acCurDbex.TransactionManager.StartTransaction()

            Dim layoutsEx As DBDictionary = TryCast(acTransEx.GetObject(acCurDbex.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

            ' Check to see if the layout exists in the external drawing
            If layoutsEx.Contains(NextlayoutName) = False Then

                Dim Layid As ObjectId = layoutsEx.GetAt(NewLayoutName)
                Dim layex As Layout = TryCast(acTransEx.GetObject(Layid, OpenMode.ForRead), Layout)
                Dim blkBlkRecEx As BlockTableRecord = acTransEx.GetObject(layex.BlockTableRecordId, OpenMode.ForRead)

                ' Get the objects from the block associated with the layout
                Dim idCol As ObjectIdCollection = New ObjectIdCollection()
                For Each id As ObjectId In blkBlkRecEx

                    idCol.Add(id)

                Next

                ' Create a transaction for the current drawing
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    ' Get the block table and create a new block then copy the objects between drawings
                    Dim blkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)

                    Using blkBlkRec As New BlockTableRecord
                        blkBlkRec.Name = "*Paper_Space" & CStr(layoutsEx.Count() - 1)

                        blkTbl.Add(blkBlkRec)
                        acTrans.AddNewlyCreatedDBObject(blkBlkRec, True)
                        Dim idmap As New IdMapping
                        acCurDb.DeepCloneObjects(idCol, blkBlkRec.ObjectId, idmap, False)

                        ' Create a new layout and then copy properties between drawings
                        Dim layouts As DBDictionary = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForWrite)

                        Using lay As New Layout
                            lay.LayoutName = NextlayoutName
                            lay.AddToLayoutDictionary(acCurDb, blkBlkRec.ObjectId)
                            acTrans.AddNewlyCreatedDBObject(lay, True)
                            lay.CopyFrom(layex)

                        End Using


                    End Using

                    ' Save the changes made
                    acTrans.Commit()
                End Using
            Else
                ' Display a message if the layout could not be found in the specified drawing
                acDoc.Editor.WriteMessage(vbLf & "Layout '" & NewLayoutName & "' could not be imported '" & "'.")
            End If

            ' Discard the changes made to the external drawing file
            acTransEx.Commit()
        End Using


    End Sub

    Private Sub ModifyViewPortCenter(NextlayoutName As String, panover As Integer, pandown As Integer, elnt As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed = acDoc.Editor

        ' Create a transaction for the external drawing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

            ' Check to see if the layout exists in the external drawing
            If layouts.Contains(NextlayoutName) = True Then

                Dim Layid As ObjectId = layouts.GetAt(NextlayoutName)
                Dim lay As Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForRead), Layout)

                '' Open the Block table for read
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Paper space for write
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                Dim blkBlkRec As BlockTableRecord = acTrans.GetObject(lay.BlockTableRecordId, OpenMode.ForRead)
                Dim vpIds As ObjectIdCollection = New ObjectIdCollection()

                For Each objID As ObjectId In blkBlkRec

                    If (objID.ObjectClass.DxfName.ToUpper = "VIEWPORT") Then

                        vpIds.Add(objID)

                    ElseIf (objID.ObjectClass.DxfName.ToUpper = "INSERT") Then

                        ' Open the block reference
                        Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(objID, OpenMode.ForRead), BlockReference)

                        Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                        Dim RevvblockName As String = RevTblRec.Name

                        If RevvblockName.Contains("Arcxis Title Block") Then

                            ' Iterate the attribute collection
                            For Each attId As ObjectId In RevBlockRef.AttributeCollection

                                ' Open the attribute reference
                                Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)

                                'attref.SetAttributeFromBlock(attDef, blkRef.BlockTransform)

                                Dim tagvalue As String = attref.Tag

                                Dim textvalue As String = attref.TextString

                                If tagvalue.Contains("ELEVATION") Then

                                    attref.TextString = elnt

                                End If

                            Next

                        End If

                    End If


                Next

                For Each vpId As ObjectId In vpIds

                    Dim layoutviewport = TryCast(acTrans.GetObject(vpId, OpenMode.ForWrite), Viewport)

                    If layoutviewport IsNot Nothing Then

                        Dim vcp As Point2d = layoutviewport.ViewCenter
                        Dim vtp As Point3d = layoutviewport.ViewTarget
                        Dim cp As Point3d = layoutviewport.CenterPoint
                        Dim wvp As Double = layoutviewport.Width
                        Dim hvp As Double = layoutviewport.Height

                        layoutviewport.ViewCenter = New Point2d((vcp.X + panover), (vcp.Y - pandown))
                        layoutviewport.CenterPoint = New Point3d(cp.X, cp.Y, cp.Z)
                        layoutviewport.Width = wvp
                        layoutviewport.Height = hvp

                    End If

                Next

            Else
                ' Display a message if the layout could not be found in the specified drawing
                acDoc.Editor.WriteMessage(vbLf & "Layout '" & NextlayoutName & "' could not be imported '" & "'.")
            End If

            ' Discard the changes made to the external drawing file
            acTrans.Commit()
        End Using

    End Sub

    Private Sub RenameFramingLayout(ExistlayoutName As String, NewlayoutName As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Create a transaction for the external drawing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

            ' Check to see if the layout exists in the external drawing
            If layouts.Contains(NewlayoutName) = False Then


                Dim Layid As ObjectId = layouts.GetAt(ExistlayoutName)
                Dim lay As Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForWrite), Layout)

                lay.LayoutName = NewlayoutName


            End If

            ' Discard the changes made to the external drawing file
            acTrans.Commit()

        End Using

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then

            TextBox2.Enabled = True

        ElseIf CheckBox1.Checked = False Then

            TextBox2.Clear()
            TextBox2.Enabled = False

        End If

    End Sub

    Private Sub SetPSLTScale()

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acLayoutMgr As LayoutManager = LayoutManager.Current

            Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                layAndTab.Add(lay.TabOrder, lay.LayoutName)
            Next

            Dim curtab As String = acLayoutMgr.CurrentLayout

            For Each layStr In layAndTab.Values

                If layStr <> "Model" Then

                    acLayoutMgr.CurrentLayout = layStr
                    Application.SetSystemVariable("Psltscale", 0)

                    Dim bt = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)

                    If bt.Has("Arcxis Title Block") Then

                        Dim blkname As String = "Arcxis Title Block"
                        GetTB(blkname)

                    End If

                    If bt.Has("Arcxis Title Block 24x36") Then

                        Dim blkname As String = "Arcxis Title Block 24x36"
                        GetTB(blkname)

                    End If

                End If

            Next

            acLayoutMgr.CurrentLayout = curtab

            ' Abort the changes to the database
            acTrans.Commit()

        End Using

    End Sub


    Private Sub GetTB(blkname As String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, blkname)}
            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
            Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)
            Dim Rev1List As New ArrayList

            If prSelRes.Status = PromptStatus.OK Then

                Dim SS As SelectionSet = prSelRes.Value
                Dim sscount As Integer = SS.Count

                If SS IsNot Nothing Then

                    For Each brId As ObjectId In prSelRes.Value.GetObjectIds()

                        Dim strutid As String = brId.ToString

                        ' Open the block reference
                        Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)
                        Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                        Dim TBName As String = TBTblRec.Name
                        Dim layidd As String = TBTblRec.LayoutId.ToString

                        If TBName = "Arcxis Title Block" OrElse TBName = "Arcxis Title Block 24x36" Then

                            Dim BLKEntity As Entity = acTrans.GetObject(brId, OpenMode.ForWrite)
                            ZoomObjects(BLKEntity)

                        End If

                    Next

                End If

            End If

            ' Abort the changes to the database
            acTrans.Commit()

        End Using

    End Sub

    Private Sub ZoomObjects(BLKEntity As Entity)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Using view As ViewTableRecord = ed.GetCurrentView()
                Dim WCS2DCS As Matrix3d = Matrix3d.PlaneToWorld(view.ViewDirection)
                WCS2DCS = Matrix3d.Displacement(view.Target - Point3d.Origin) * WCS2DCS
                WCS2DCS = Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * WCS2DCS
                WCS2DCS = WCS2DCS.Inverse()
                Dim ext As Extents3d = BLKEntity.GeometricExtents
                Dim tmp As Extents3d = BLKEntity.GeometricExtents
                ext.AddExtents(tmp)
                ext.TransformBy(WCS2DCS)
                view.Width = ext.MaxPoint.X - ext.MinPoint.X
                view.Height = ext.MaxPoint.Y - ext.MinPoint.Y
                view.CenterPoint = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2.0, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0)
                ed.SetCurrentView(view)
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.Text = "A (R)" Then

            ComboBox2.Text = "A opts. (R)"

        ElseIf ComboBox1.Text = "F-1A (R)" Then

            ComboBox2.Text = "F-1A opts. (R)"

        ElseIf ComboBox1.Text = "A1(R)" Then

            ComboBox2.Text = "A1(R)OP1"

        End If

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        TextBox1.BackColor = System.Drawing.Color.White

        If (Not Char.IsControl(e.KeyChar) _
                     AndAlso (Not Char.IsDigit(e.KeyChar) _
                     AndAlso (e.KeyChar <> Microsoft.VisualBasic.ChrW(46)))) Then
            e.Handled = True

        End If

        Button1.Enabled = True

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress

        TextBox2.BackColor = System.Drawing.Color.White

        If (Not Char.IsControl(e.KeyChar) _
                     AndAlso (Not Char.IsDigit(e.KeyChar) _
                     AndAlso (e.KeyChar <> Microsoft.VisualBasic.ChrW(46)))) Then
            e.Handled = True

        End If

        'Button1.Enabled = True

    End Sub

    Private Sub GetExistLayout()

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor
        Dim laynum As Integer
        Dim laynam As String
        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                layAndTab.Add(lay.TabOrder, lay.LayoutName)

                If lay.LayoutName <> "model" Then

                    If lay.LayoutName.Contains("OPT") Then

                        ComboBox2.Text = lay.LayoutName

                    Else

                        ComboBox1.Text = lay.LayoutName


                    End If


                End If


            Next

            For Each layStr In layAndTab.Values


                If layStr <> "Model" Then

                    'frm.ListBox1.Items.Add(layStr)

                End If

            Next

        End Using


    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        If CheckBox2.Checked = True Then

            ComboBox1.Text = ATB_FirLayoutName
            ComboBox2.Text = ATB_SecLayoutName

        ElseIf CheckBox2.Checked = False Then

            ComboBox1.Text = "A (R)"
            ComboBox2.Text = "A opts. (R)"


        End If


    End Sub

End Class