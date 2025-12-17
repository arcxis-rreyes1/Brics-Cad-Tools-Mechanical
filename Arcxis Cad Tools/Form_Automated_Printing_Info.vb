Imports System
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Windows.Controls
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports System.Drawing

Public Class Form_Automated_Printing_Info

    Private lockOthers As Boolean = False
    Private lockOthers1 As Boolean = False
    Private lockOthers2 As Boolean = False
    Private Sub Form_Automated_Printing_Info_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim PlanType As New List(Of String)
        Dim Swings As New List(Of String)
        Dim Elevations As New List(Of String)
        Dim Options As New List(Of String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Dim frm As New Form_Automated_Printing_Info

        Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
            ' Open the Block table for read
            Dim blkTable As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)

            ' Open the BlockTableRecord (ModelSpace) for read
            Dim blkTableRec As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            ' Iterate through the ModelSpace block table record
            For Each objId As ObjectId In blkTableRec

                Dim entity As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                ' Check if the entity is a block reference
                If TypeOf entity Is BlockReference Then

                    Dim blkRef As BlockReference = CType(entity, BlockReference)

                    ' Get the block table record for the block reference
                    Dim blkDef As BlockTableRecord = TryCast(acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                    ' Check if the block name is "PAGE"
                    If blkDef.Name.ToUpper() = "PAGE" Then

                        Dim blockobjecthandle As String = blkRef.Handle.Value.ToString()


                        ' Check if the block reference has attributes
                        If blkRef.AttributeCollection.Count > 0 Then
                            ' Iterate through the attributes
                            For Each attId As ObjectId In blkRef.AttributeCollection
                                Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForRead), AttributeReference)

                                ' Check the attribute tag and add the value to the appropriate list
                                If attRef IsNot Nothing Then
                                    Select Case attRef.Tag.ToUpper()

                                        Case "MOD"

                                            If Not PlanType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                PlanType.Add(attRef.TextString)
                                            End If

                                        Case "ELEV"

                                            If attRef.TextString.Contains(",") Then
                                                Dim values() As String = attRef.TextString.Split(","c) ' Split the string into an array of values

                                                For Each value As String In values

                                                    If Not Elevations.Contains(value) Then

                                                        Elevations.Add(value)

                                                    End If ' Add each value to the TempElev list, trimming any extra spaces

                                                Next

                                            Else

                                                If Not Elevations.Contains(attRef.TextString) Then

                                                    Elevations.Add(attRef.TextString)

                                                End If

                                            End If

                                        Case "SW"

                                            If Swings.Contains(attRef.TextString) Then

                                                Swings.Add(attRef.TextString)

                                            End If

                                        Case "OPTIONS"


                                            If Not Options.Contains(attRef.TextString) Then

                                                Options.Add(attRef.TextString)

                                            End If

                                        Case "SIZE"

                                    End Select
                                End If
                            Next
                        End If
                    End If
                    'End If
                End If
            Next


            Elevations.Sort()
            Swings.Sort()
            PlanType.Sort()

            For Each item In Elevations
                CheckedListBox5.Items.Add(item)
            Next

        End Using


        CheckedListBox1.SetItemChecked(0, True)
        CheckedListBox4.SetItemChecked(0, True)
        CheckedListBox5.SetItemChecked(0, True)
        Dim fm As New Arcxis_Cad_Tools.FileManipulation()

        Dim PlanStamps = fm.GetCustomDwgPropReliable("STAMPS")
        If PlanStamps <> "" Then
            If PlanStamps.Contains("TML-TX") Then
                SealsList.SetItemChecked(0, True)
            End If
        End If

        lockOthers = True
        lockOthers1 = True
        lockOthers2 = True

        ResizeCheckListBoxAndGroupBox()

    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) _
    Handles CheckedListBox1.ItemCheck

        Me.BeginInvoke(Sub()
                           Dim clb = DirectCast(sender, CheckedListBox)
                           If clb.Items.Count = 0 Then Return

                           ' Build post-click states (what the checks will be after this click)
                           Dim states As Boolean() = Enumerable.Range(0, clb.Items.Count) _
            .Select(Function(i As Integer)
                        If i = e.Index Then
                            Return e.NewValue = CheckState.Checked
                        Else
                            Return clb.GetItemChecked(i)
                        End If
                    End Function) _
            .ToArray()

                           ' 1) If user is CHECKING item(0), make it the only checked and select it.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Checked Then
                               For i As Integer = 0 To clb.Items.Count - 1
                                   clb.SetItemChecked(i, i = 0)
                               Next
                               clb.SelectedIndex = 0
                               Return
                           End If

                           ' 2) If user is UNCHECKING item(0), clear its highlight.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Unchecked Then
                               clb.ClearSelected()
                               ' continue—no other action needed
                           End If

                           ' 3) If a non-zero item is being checked while item(0) is (now) checked, uncheck item(0).
                           If e.Index <> 0 AndAlso e.NewValue = CheckState.Checked AndAlso (states(0) OrElse clb.GetItemChecked(0)) Then
                               clb.SetItemChecked(0, False)
                               clb.SelectedIndex = e.Index  ' highlight the newly-checked item
                               states(0) = False             ' keep local states consistent
                           End If

                           ' 4) If all non-zero items are checked, collapse to only item(0).
                           If clb.Items.Count > 1 Then
                               Dim allNonZeroChecked As Boolean =
                Enumerable.Range(1, clb.Items.Count - 1) _
                          .All(Function(i As Integer)
                                   ' use states (post-click) for accuracy
                                   Return states(i)
                               End Function)

                               If allNonZeroChecked Then
                                   For i As Integer = 0 To clb.Items.Count - 1
                                       clb.SetItemChecked(i, i = 0)
                                   Next
                                   clb.SelectedIndex = 0
                               End If
                           End If
                       End Sub)
    End Sub

    Private Sub CheckedListBOX5_ItemCheck(sender As Object, e As ItemCheckEventArgs) _
    Handles CheckedListBox5.ItemCheck

        Me.BeginInvoke(Sub()
                           Dim clb = DirectCast(sender, CheckedListBox)
                           If clb.Items.Count = 0 Then Return

                           ' Build post-click states (what the checks will be after this click)
                           Dim states As Boolean() = Enumerable.Range(0, clb.Items.Count) _
            .Select(Function(i As Integer)
                        If i = e.Index Then
                            Return e.NewValue = CheckState.Checked
                        Else
                            Return clb.GetItemChecked(i)
                        End If
                    End Function) _
            .ToArray()

                           ' 1) If user is CHECKING item(0), make it the only checked and select it.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Checked Then
                               For i As Integer = 0 To clb.Items.Count - 1
                                   clb.SetItemChecked(i, i = 0)
                               Next
                               clb.SelectedIndex = 0
                               Return
                           End If

                           ' 2) If user is UNCHECKING item(0), clear its highlight.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Unchecked Then
                               clb.ClearSelected()
                               ' continue—no other action needed
                           End If

                           ' 3) If a non-zero item is being checked while item(0) is (now) checked, uncheck item(0).
                           If e.Index <> 0 AndAlso e.NewValue = CheckState.Checked AndAlso (states(0) OrElse clb.GetItemChecked(0)) Then
                               clb.SetItemChecked(0, False)
                               clb.SelectedIndex = e.Index  ' highlight the newly-checked item
                               states(0) = False             ' keep local states consistent
                           End If

                           ' 4) If all non-zero items are checked, collapse to only item(0).
                           If clb.Items.Count > 1 Then
                               Dim allNonZeroChecked As Boolean =
                Enumerable.Range(1, clb.Items.Count - 1) _
                          .All(Function(i As Integer)
                                   ' use states (post-click) for accuracy
                                   Return states(i)
                               End Function)

                               If allNonZeroChecked Then
                                   For i As Integer = 0 To clb.Items.Count - 1
                                       clb.SetItemChecked(i, i = 0)
                                   Next
                                   clb.SelectedIndex = 0
                               End If
                           End If
                       End Sub)
    End Sub

    Private Sub CheckedListBOX4_ItemCheck(sender As Object, e As ItemCheckEventArgs) _
    Handles CheckedListBox4.ItemCheck

        Me.BeginInvoke(Sub()
                           Dim clb = DirectCast(sender, CheckedListBox)
                           If clb.Items.Count = 0 Then Return

                           ' Build post-click states (what the checks will be after this click)
                           Dim states As Boolean() = Enumerable.Range(0, clb.Items.Count) _
            .Select(Function(i As Integer)
                        If i = e.Index Then
                            Return e.NewValue = CheckState.Checked
                        Else
                            Return clb.GetItemChecked(i)
                        End If
                    End Function) _
            .ToArray()

                           ' 1) If user is CHECKING item(0), make it the only checked and select it.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Checked Then
                               For i As Integer = 0 To clb.Items.Count - 1
                                   clb.SetItemChecked(i, i = 0)
                               Next
                               clb.SelectedIndex = 0
                               Return
                           End If

                           ' 2) If user is UNCHECKING item(0), clear its highlight.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Unchecked Then
                               clb.ClearSelected()
                               ' continue—no other action needed
                           End If

                           ' 3) If a non-zero item is being checked while item(0) is (now) checked, uncheck item(0).
                           If e.Index <> 0 AndAlso e.NewValue = CheckState.Checked AndAlso (states(0) OrElse clb.GetItemChecked(0)) Then
                               clb.SetItemChecked(0, False)
                               clb.SelectedIndex = e.Index  ' highlight the newly-checked item
                               states(0) = False             ' keep local states consistent
                           End If

                           ' 4) If all non-zero items are checked, collapse to only item(0).
                           If clb.Items.Count > 1 Then
                               Dim allNonZeroChecked As Boolean =
                Enumerable.Range(1, clb.Items.Count - 1) _
                          .All(Function(i As Integer)
                                   ' use states (post-click) for accuracy
                                   Return states(i)
                               End Function)

                               If allNonZeroChecked Then
                                   For i As Integer = 0 To clb.Items.Count - 1
                                       clb.SetItemChecked(i, i = 0)
                                   Next
                                   clb.SelectedIndex = 0
                               End If
                           End If
                       End Sub)
    End Sub

    Private Sub ResizeCheckListBoxAndGroupBox()

        Me.SuspendLayout()
        Dim RealListCount As Integer = CheckedListBox5.Items.Count + 1
        ' 1) CheckedListBox5: cap to 8 visible rows
        Dim maxVisibleItems As Integer = 8
        Dim visibleCount As Integer = Math.Min(RealListCount, maxVisibleItems)
        CheckedListBox5.Height = (visibleCount * CheckedListBox5.ItemHeight) - 10

        ' 2) GroupBox5 wraps the checklist
        Dim paddingBelow As Integer = 6
        Dim titleBarHeight5 As Integer = GroupBox5.Height - GroupBox5.ClientSize.Height
        GroupBox5.Height = CheckedListBox5.Top + CheckedListBox5.Height + paddingBelow + titleBarHeight5


        Dim RealListCount1 As Integer = SealsList.Items.Count + 1
        ' 1) CheckedListBox5: cap to 8 visible rows
        Dim maxVisibleItems1 As Integer = 8
        Dim RealListCoun1t As Integer = SealsList.Items.Count + 1
        Dim visibleCount1 As Integer = Math.Min(RealListCount1, maxVisibleItems1)
        SealsList.Height = (visibleCount1 * SealsList.ItemHeight) - 10


        ' 2) GroupBox5 wraps the checklist
        Dim paddingBelow1 As Integer = 6
        Dim titleBarHeightsEALS As Integer = GroupBox7.Height - GroupBox7.ClientSize.Height
        GroupBox7.Height = SealsList.Top + SealsList.Height + paddingBelow + titleBarHeightsEALS
        GroupBox7.Top = GroupBox5.Bottom + 6

        ' 3) Move buttons just below GroupBox5
        Dim spacingBelowGroupBox As Integer = 8
        Button1.Top = GroupBox7.Bottom + spacingBelowGroupBox
        Button2.Top = GroupBox7.Bottom + spacingBelowGroupBox

        ' 4) GroupBox1 wraps GroupBox5 + buttons
        Dim titleBarHeight1 As Integer = GroupBox1.Height - GroupBox1.ClientSize.Height
        Dim bottomMost As Integer = Math.Max(Button1.Bottom, Button2.Bottom)
        Dim paddingBottomGB1 As Integer = 20
        GroupBox1.Height = (bottomMost + paddingBottomGB1 + titleBarHeight1) - GroupBox1.Top

        ' 5) Form grows to contain GroupBox1
        Dim paddingBottomForm As Integer = 20 ' extra breathing room
        Dim neededClientHeight As Integer = GroupBox1.Top + GroupBox1.Height + paddingBottomForm

        ' Keep current width; adjust height only
        Dim newClientSize As New System.Drawing.Size(Me.ClientSize.Width, neededClientHeight)

        ' Optional: cap to screen working area (avoids going off-screen)
        Dim screenMaxHeight As Integer = Screen.FromControl(Me).WorkingArea.Height
        If newClientSize.Height > screenMaxHeight Then
            newClientSize.Height = screenMaxHeight
            Me.AutoScroll = True ' enable scroll if needed
        End If

        ' Apply client size
        Me.ClientSize = newClientSize

        ' Optional: enforce a minimum height
        Me.MinimumSize = New System.Drawing.Size(Me.MinimumSize.Width, Me.Height)

        Me.ResumeLayout()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub SheetLabels_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SheetLabels.SelectedIndexChanged
        SheetLabels.BackColor = System.Drawing.Color.White
        Button1.Visible = True
    End Sub
End Class