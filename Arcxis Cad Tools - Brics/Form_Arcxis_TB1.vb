Imports System
Imports System.Drawing.Printing
Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports PlotType = Teigha.DatabaseServices.PlotType
Imports Exception = Teigha.Runtime.Exception
Imports Teigha.GraphicsInterface
Imports Viewport = Teigha.DatabaseServices.Viewport
Imports OpenMode = Teigha.DatabaseServices.OpenMode
Imports System.Windows.Media
Imports Color = Teigha.Colors.Color

Public Class Form_Arcxis_TB1


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then

            For i As Integer = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, True)

            Next i

        ElseIf CheckBox1.Checked = False Then

            For i As Integer = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, False)
            Next i

        End If

        Me.Update()
    End Sub

    'Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    '    ' Get the current document and database
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim aced As Editor = acDoc.Editor
    '    Dim selectList As New ArrayList
    '    Dim Listbox1List As New ArrayList
    '    Dim slc As Integer = ListBox1.SelectedItems.Count


    '    If slc <> 0 Then
    '        For Each selectitem As String In ListBox1.SelectedItems

    '            selectList.Add(selectitem)

    '        Next

    '        Dim curItem As String

    '        Dim selectcount As Integer = ListBox1.SelectedItems.Count
    '        Dim listcount As Integer = Listbox1List.Count

    '        If selectcount > listcount Then

    '            For Each item As String In ListBox1.SelectedItems

    '                If Not Listbox1List.Contains(item) Then
    '                    Listbox1List.Add(item)

    '                End If

    '            Next

    '        Else

    '            For Each item As String In Listbox1List

    '                If Not selectList.Contains(item) Then
    '                    Listbox1List.Remove(item)
    '                    Exit For

    '                End If

    '            Next

    '        End If

    '        curItem = Listbox1List(Listbox1List.Count - 1)

    '        '' Lock the new document
    '        Using acLckDoc As DocumentLock = acDoc.LockDocument()

    '            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

    '                Dim blkname As String = "Arcxis Title Block"
    '                Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, curItem)}
    '                Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
    '                Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)
    '                Dim Rev1List As New ArrayList
    '                'Dim Rev1List As List(Of String) = New List(Of String)

    '                If prSelRes.Status = PromptStatus.OK Then

    '                    Dim SS As SelectionSet = prSelRes.Value
    '                    Dim sscount As Integer = SS.Count

    '                    If SS IsNot Nothing Then

    '                        For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
    '                            If brId.ObjectClass.Name = "AcDbBlockReference" Then
    '                                Dim strutid As String = brId.ToString

    '                                ' Open the block reference
    '                                Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

    '                                Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
    '                                Dim TBName As String = TBTblRec.Name


    '                                ' Iterate the attribute collection
    '                                For Each attId As ObjectId In TBBlockRef.AttributeCollection

    '                                    ' Open the attribute reference
    '                                    Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
    '                                    Dim tagvalue As String = attref.Tag


    '                                    If tagvalue = "PLAN " Then

    '                                        TextBox1.Text = attref.TextString

    '                                    ElseIf tagvalue = "ELEVATION" Then

    '                                        ComboBox1.Text = attref.TextString

    '                                    ElseIf tagvalue.Contains("OPTIONAL") Then

    '                                        ComboBox2.Text = attref.TextString

    '                                    ElseIf tagvalue = "SUBDIVISION" Then

    '                                        TextBox2.Text = attref.TextString

    '                                    ElseIf tagvalue = "ADDRESS" Then

    '                                        TextBox3.Text = attref.TextString

    '                                    ElseIf tagvalue = "CUSTOMER'S NAME" Then

    '                                        ComboBox3.Text = attref.TextString

    '                                    ElseIf tagvalue = "PROJECT#" Then

    '                                        TextBox4.Text = attref.TextString

    '                                    ElseIf tagvalue = "DES" Then

    '                                        'TextBox5.Text = attref.TextString

    '                                    ElseIf tagvalue = "CHK" Then

    '                                        'TextBox6.Text = attref.TextString

    '                                    ElseIf tagvalue = "1/8"" = 1'-0""" Then

    '                                        ComboBox4.Text = attref.TextString

    '                                    ElseIf tagvalue = "PLANDATE" Then

    '                                        TextBox8.Text = attref.TextString

    '                                    ElseIf tagvalue = "FR-1" Then

    '                                        TextBox7.Text = attref.TextString

    '                                    ElseIf tagvalue = "1" Then

    '                                        TextBox10.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE1" Then

    '                                        TextBox11.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION1" Then

    '                                        TextBox12.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL1" Then

    '                                        TextBox13.Text = attref.TextString

    '                                    ElseIf tagvalue = "2" Then

    '                                        TextBox14.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE2" Then

    '                                        TextBox15.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION2" Then

    '                                        TextBox16.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL2" Then

    '                                        TextBox17.Text = attref.TextString

    '                                    ElseIf tagvalue = "3" Then

    '                                        TextBox18.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE3" Then

    '                                        TextBox19.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION3" Then

    '                                        TextBox20.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL3" Then

    '                                        TextBox21.Text = attref.TextString

    '                                    ElseIf tagvalue = "4" Then

    '                                        TextBox22.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE4" Then

    '                                        TextBox23.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION4" Then

    '                                        TextBox24.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL4" Then

    '                                        TextBox25.Text = attref.TextString

    '                                    ElseIf tagvalue = "5" Then

    '                                        TextBox26.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE5" Then

    '                                        TextBox27.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION5" Then

    '                                        TextBox28.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL5" Then

    '                                        TextBox29.Text = attref.TextString

    '                                    ElseIf tagvalue = "6" Then

    '                                        TextBox30.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE6" Then

    '                                        TextBox31.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION6" Then

    '                                        TextBox32.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL6" Then

    '                                        TextBox33.Text = attref.TextString

    '                                    ElseIf tagvalue = "7" Then

    '                                        TextBox34.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDATE7" Then

    '                                        TextBox35.Text = attref.TextString

    '                                    ElseIf tagvalue = "REVDESCRIPTION7" Then

    '                                        TextBox36.Text = attref.TextString

    '                                    ElseIf tagvalue = "INITIAL7" Then

    '                                        TextBox37.Text = attref.TextString

    '                                    ElseIf tagvalue = "LOT" Then

    '                                        TextBox38.Text = attref.TextString

    '                                    ElseIf tagvalue = "BLK" Then

    '                                        TextBox42.Text = attref.TextString

    '                                    ElseIf tagvalue = "SEC" Then

    '                                        TextBox43.Text = attref.TextString


    '                                    ElseIf tagvalue = "BUILDER" Then

    '                                        TextBox5.Text = attref.TextString

    '                                    ElseIf tagvalue = "BUILDER_ADDRESS" Then

    '                                        TextBox6.Text = attref.TextString

    '                                    ElseIf tagvalue = "CITY_STATE_ZIP" Then

    '                                        TextBox44.Text = attref.TextString

    '                                    ElseIf tagvalue = "OFFICE_PHONE" Then

    '                                        TextBox45.Text = attref.TextString

    '                                    ElseIf tagvalue = "ARCHITECTUAL_DATE" Then

    '                                        TextBox46.Text = attref.TextString

    '                                    ElseIf tagvalue = "STRUCTURAL_DATE" Then

    '                                        TextBox47.Text = attref.TextString

    '                                    End If

    '                                Next

    '                            End If

    '                        Next

    '                    End If

    '                End If

    '                acTrans.Commit()

    '            End Using

    '        End Using

    '        Button3.Enabled = False
    '        Button3.BackColor = System.Drawing.SystemColors.Control
    '        Button1.Enabled = False
    '        Button1.BackColor = System.Drawing.SystemColors.Control
    '        Button4.Enabled = True

    '        Module_Arcxis_TB.ATB_CurLayoutName = curItem

    '    End If

    'End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor
        Dim selectList As New ArrayList
        Dim Listbox1List As New ArrayList
        Dim slc As Integer = ListBox1.SelectedItems.Count

        If ListBox1.SelectedItems.Count = 1 Then
            TextBox1.Enabled = True
            TextBox7.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
        Else
            TextBox1.Enabled = False
            TextBox7.Enabled = False
            ComboBox1.Enabled = False
            ComboBox1.Enabled = False
        End If

        If slc <> 0 Then
            For Each selectitem As String In ListBox1.SelectedItems

                selectList.Add(selectitem)

            Next

            Dim curItem As String

            Dim selectcount As Integer = ListBox1.SelectedItems.Count
            Dim listcount As Integer = Listbox1List.Count

            If selectcount > listcount Then

                For Each item As String In ListBox1.SelectedItems

                    If Not Listbox1List.Contains(item) Then
                        Listbox1List.Add(item)

                    End If

                Next

            Else

                For Each item As String In Listbox1List

                    If Not selectList.Contains(item) Then
                        Listbox1List.Remove(item)
                        Exit For

                    End If

                Next

            End If

            curItem = Listbox1List(Listbox1List.Count - 1)

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim blkname As String = "Arcxis Title Block"
                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, curItem)}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)
                    Dim Rev1List As New ArrayList
                    'Dim Rev1List As List(Of String) = New List(Of String)

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value
                        Dim sscount As Integer = SS.Count

                        If SS IsNot Nothing Then

                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                                If brId.ObjectClass.Name = "AcDbBlockReference" Then
                                    Dim strutid As String = brId.ToString

                                    ' Open the block reference
                                    Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                                    Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                    Dim TBName As String = TBTblRec.Name


                                    ' Iterate the attribute collection
                                    For Each attId As ObjectId In TBBlockRef.AttributeCollection

                                        ' Open the attribute reference
                                        Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                        Dim tagvalue As String = attref.Tag


                                        If tagvalue = "PLAN " Then

                                            TextBox1.Text = attref.TextString

                                        ElseIf tagvalue = "ELEVATION" Then

                                            ComboBox1.Text = attref.TextString

                                        ElseIf tagvalue.Contains("OPTIONAL") Then

                                            ComboBox2.Text = attref.TextString

                                        ElseIf tagvalue = "SUBDIVISION" Then

                                            TextBox2.Text = attref.TextString

                                        ElseIf tagvalue = "ADDRESS" Then

                                            TextBox3.Text = attref.TextString

                                        ElseIf tagvalue = "CUSTOMER'S NAME" Then

                                            ComboBox3.Text = attref.TextString

                                        ElseIf tagvalue = "PROJECT#" Then

                                            TextBox4.Text = attref.TextString

                                        ElseIf tagvalue = "DES" Then

                                            'TextBox5.Text = attref.TextString

                                        ElseIf tagvalue = "CHK" Then

                                            'TextBox6.Text = attref.TextString

                                        ElseIf tagvalue = "1/8"" = 1'-0""" Then

                                            ComboBox4.Text = attref.TextString

                                        ElseIf tagvalue = "PLANDATE" Then

                                            TextBox8.Text = attref.TextString

                                        ElseIf tagvalue = "FR-1" Then

                                            TextBox7.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE1" Then

                                            TextBox10.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION1" Then

                                            TextBox12.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL1" Then

                                            TextBox13.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE2" Then

                                            TextBox14.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION2" Then

                                            TextBox16.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL2" Then

                                            TextBox17.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE3" Then

                                            TextBox18.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION3" Then

                                            TextBox20.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL3" Then

                                            TextBox21.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE4" Then

                                            TextBox22.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION4" Then

                                            TextBox24.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL4" Then

                                            TextBox25.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE5" Then

                                            TextBox26.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION5" Then

                                            TextBox28.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL5" Then

                                            TextBox29.Text = attref.TextString

                                        ElseIf tagvalue = "REVDATE6" Then

                                            TextBox30.Text = attref.TextString

                                        ElseIf tagvalue = "REVDESCRIPTION6" Then

                                            TextBox32.Text = attref.TextString

                                        ElseIf tagvalue = "INITIAL6" Then

                                            TextBox33.Text = attref.TextString

                                        ElseIf tagvalue = "LOT" Then

                                            TextBox38.Text = attref.TextString

                                        ElseIf tagvalue = "BLK" Then

                                            TextBox42.Text = attref.TextString

                                        ElseIf tagvalue = "SEC" Then

                                            TextBox43.Text = attref.TextString


                                        ElseIf tagvalue = "BUILDER" Then

                                            TextBox5.Text = attref.TextString

                                        ElseIf tagvalue = "BUILDER_ADDRESS" Then

                                            TextBox6.Text = attref.TextString

                                        ElseIf tagvalue = "CITY_STATE_ZIP" Then

                                            TextBox44.Text = attref.TextString

                                        ElseIf tagvalue = "OFFICE_PHONE" Then

                                            TextBox45.Text = attref.TextString

                                        ElseIf tagvalue = "ARCHITECTUAL_DATE" Then

                                            TextBox46.Text = attref.TextString

                                        ElseIf tagvalue = "STRUCTURAL_DATE" Then

                                            TextBox47.Text = attref.TextString

                                        End If

                                    Next

                                End If

                            Next

                        End If

                    End If

                    acTrans.Commit()

                End Using

            End Using

            Button3.Enabled = False
            Button3.BackColor = System.Drawing.SystemColors.Control
            Button1.Enabled = False
            Button1.BackColor = System.Drawing.SystemColors.Control
            Button4.Enabled = True
            Button8.Enabled = True

            Module_Arcxis_TB.ATB_CurLayoutName = curItem
        End If

        If ListBox1.SelectedItems.Count > 0 Then

            PlotPDF11x17.Enabled = True
            PDF24x36.Enabled = True

        End If
    End Sub

    Private Sub ReadForm(curitem)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim blkname As String = "Arcxis Title Block"
                Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, curitem)}
                Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)
                Dim Rev1List As New ArrayList
                'Dim Rev1List As List(Of String) = New List(Of String)

                If prSelRes.Status = PromptStatus.OK Then

                    Dim SS As SelectionSet = prSelRes.Value
                    Dim sscount As Integer = SS.Count

                    If SS IsNot Nothing Then

                        For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                            If brId.ObjectClass.Name = "AcDbBlockReference" Then
                                Dim strutid As String = brId.ToString

                                ' Open the block reference
                                Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                                Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                Dim TBName As String = TBTblRec.Name


                                ' Iterate the attribute collection
                                For Each attId As ObjectId In TBBlockRef.AttributeCollection

                                    ' Open the attribute reference
                                    Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                    Dim tagvalue As String = attref.Tag


                                    If tagvalue = "PLAN " Then

                                        TextBox1.Text = attref.TextString

                                    ElseIf tagvalue = "ELEVATION" Then

                                        ComboBox1.Text = attref.TextString

                                    ElseIf tagvalue.Contains("OPTIONAL") Then

                                        ComboBox2.Text = attref.TextString

                                    ElseIf tagvalue = "SUBDIVISION" Then

                                        TextBox2.Text = attref.TextString

                                    ElseIf tagvalue = "ADDRESS" Then

                                        TextBox3.Text = attref.TextString

                                    ElseIf tagvalue = "CUSTOMER'S NAME" Then

                                        ComboBox3.Text = attref.TextString

                                    ElseIf tagvalue = "PROJECT#" Then

                                        TextBox4.Text = attref.TextString

                                    ElseIf tagvalue = "DES" Then

                                        'TextBox5.Text = attref.TextString

                                    ElseIf tagvalue = "CHK" Then

                                        'TextBox6.Text = attref.TextString

                                    ElseIf tagvalue = "1/8"" = 1'-0""" Then

                                        ComboBox4.Text = attref.TextString

                                    ElseIf tagvalue = "PLANDATE" Then

                                        'TextBox8.Text = attref.TextString

                                    ElseIf tagvalue = "FR-1" Then

                                        TextBox7.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE1" Then

                                        TextBox10.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION1" Then

                                        TextBox12.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL1" Then

                                        TextBox13.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE2" Then

                                        TextBox14.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION2" Then

                                        TextBox16.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL2" Then

                                        TextBox17.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE3" Then

                                        TextBox18.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION3" Then

                                        TextBox20.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL3" Then

                                        TextBox21.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE4" Then

                                        TextBox22.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION4" Then

                                        TextBox24.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL4" Then

                                        TextBox25.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE5" Then

                                        TextBox26.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION5" Then

                                        TextBox28.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL5" Then

                                        TextBox29.Text = attref.TextString

                                    ElseIf tagvalue = "REVDATE6" Then

                                        TextBox30.Text = attref.TextString

                                    ElseIf tagvalue = "REVDESCRIPTION6" Then

                                        TextBox32.Text = attref.TextString

                                    ElseIf tagvalue = "INITIAL6" Then

                                        TextBox33.Text = attref.TextString

                                    ElseIf tagvalue = "LOT" Then

                                        TextBox38.Text = attref.TextString

                                    ElseIf tagvalue = "BLK" Then

                                        TextBox42.Text = attref.TextString

                                    ElseIf tagvalue = "SEC" Then

                                        TextBox43.Text = attref.TextString


                                    ElseIf tagvalue = "BUILDER" Then

                                        TextBox5.Text = attref.TextString

                                    ElseIf tagvalue = "BUILDER_ADDRESS" Then

                                        TextBox6.Text = attref.TextString

                                    ElseIf tagvalue = "CITY_STATE_ZIP" Then

                                        TextBox44.Text = attref.TextString

                                    ElseIf tagvalue = "OFFICE_PHONE" Then

                                        TextBox45.Text = attref.TextString

                                    ElseIf tagvalue = "ARCHITECTUAL_DATE" Then

                                        TextBox46.Text = attref.TextString

                                    ElseIf tagvalue = "STRUCTURAL_DATE" Then

                                        TextBox47.Text = attref.TextString

                                    End If

                                Next

                            End If

                        Next

                    End If

                End If

                acTrans.Commit()

            End Using

        End Using
    End Sub



    Private Sub TextBox9_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

        Button1.Enabled = True
        Button1.BackColor = System.Drawing.Color.Red

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim frm As New Form_Arcxis_TB1
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim selectedtabs As New ArrayList
        For Each Layout1 In ListBox1.SelectedItems
            selectedtabs.Add(Layout1)

        Next
        '' Lock the new document
        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                For Each item As Object In ListBox1.SelectedItems

                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, item)}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                                If brId.ObjectClass.Name = "AcDbBlockReference" Then
                                    Dim strutid As String = brId.ToString


                                    ' Open the block reference
                                    Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)
                                    Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                    Dim RevvblockName As String = RevTblRec.Name

                                    ' Iterate the attribute collection
                                    For Each attId As ObjectId In RevBlockRef.AttributeCollection



                                        ' Open the attribute reference
                                        Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                        Dim tagvalue As String = attref.Tag
                                        Dim textvalue As String = attref.TextString
                                        If tagvalue.Contains("REVDESCRIPTION") AndAlso textvalue = "-" Then

                                            Dim RevDescTagValue As String = tagvalue
                                            Dim RevDescNum As String = RevDescTagValue.Substring(14)
                                            Dim RevNumTagValue As String = RevDescNum
                                            Dim RevInitTag As String = "INITIAL" & RevDescNum
                                            Dim RevDateTag As String = "REVDATE" & RevDescNum

                                            For Each attIdd As ObjectId In RevBlockRef.AttributeCollection
                                                Dim attreff As AttributeReference = DirectCast(acTrans.GetObject(attIdd, OpenMode.ForWrite), AttributeReference)
                                                Dim taggvalue As String = attreff.Tag

                                                If taggvalue = RevNumTagValue Then

                                                    attreff.TextString = RevDescNum

                                                ElseIf taggvalue = RevDateTag Then

                                                    attreff.TextString = DateTimePicker1.Text

                                                ElseIf taggvalue = RevDescTagValue Then

                                                    attreff.TextString = ComboBox5.Text

                                                ElseIf taggvalue = RevInitTag Then

                                                    attreff.TextString = TextBox9.Text

                                                End If

                                            Next

                                            Exit For

                                        End If


                                    Next
                                End If
                            Next

                        End If

                    End If

                Next

                acTrans.Commit()
            End Using

        End Using
        ListBox1.Items.Clear()
        Button1.BackColor = System.Drawing.SystemColors.Control
        Button1.Enabled = False
        frm.Update()
        frm.Refresh()
        GetLayoutList()
        For Each thing1 In selectedtabs
            ListBox1.SelectedItems.Add(thing1)
        Next

        If TextBox33.Text <> "-" Then
            ComboBox5.Enabled = False
            DateTimePicker1.Enabled = False
            TextBox9.Enabled = False
        End If

    End Sub

    Private Sub LayoutList()

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                layAndTab.Add(lay.TabOrder, lay.LayoutName)
            Next

            For Each layStr In layAndTab.Values
                'ed.WriteMessage(v & vbCrLf)

                If layStr <> "Model" Then


                End If

            Next

            ' Abort the changes to the database
            acTrans.Abort()

        End Using

    End Sub

    Private Sub UpdateAttributesInDatabase(acCurDb As Database, attbName As String, attbvaule As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Get the IDs of the spaces we want to process
        ' and simply call a function to process each
        Dim msId As ObjectId, psId As ObjectId

        Dim blockname As String = "Arcxis Title Block"

        Dim tr As Transaction = acCurDb.TransactionManager.StartTransaction()

        Using tr

            Dim bt As BlockTable = DirectCast(tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
            msId = bt(BlockTableRecord.ModelSpace)
            psId = bt(BlockTableRecord.PaperSpace)

            ' Not needed, but quicker than aborting
            tr.Commit()
        End Using

        Dim msCount As Integer = UpdateAttributesInBlock(msId, blockname, attbName, attbvaule)
        Dim psCount As Integer = UpdateAttributesInBlock(psId, blockname, attbName, attbvaule)

        ed.Regen()

        ' Display the results
        ed.WriteMessage(vbLf & "Processing file: " + acCurDb.Filename)
        ed.WriteMessage(vbLf & "Updated {0} instance{1} of " + "attribute {2} in the modelspace.", msCount, If(msCount = 1, "", "s"), attbName)
        ed.WriteMessage(vbLf & "Updated {0} instance{1} of " + "attribute {2} in the default paperspace.", psCount, If(psCount = 1, "", "s"), attbName)

    End Sub


    Private Function UpdateAttributesInBlock(btrId As ObjectId, blockName As String, attbName As String, attbValue As String) As Integer

        ' Will return the number of attributes modified
        Dim changedCount As Integer = 0
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim tr As Transaction = doc.TransactionManager.StartTransaction()

        Using tr

            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)

            ' Test each entity in the container...
            For Each entId As ObjectId In btr

                Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)

                If ent IsNot Nothing Then

                    Dim br As BlockReference = TryCast(ent, BlockReference)

                    If br IsNot Nothing Then

                        Dim bd As BlockTableRecord = DirectCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                        ' ... to see whether it's a block with
                        ' the name we're after
                        If bd.Name.ToUpper() = blockName Then



                            ' Check each of the attributes...
                            For Each arId As ObjectId In br.AttributeCollection

                                Dim obj As DBObject = tr.GetObject(arId, OpenMode.ForRead)

                                Dim ar As AttributeReference = TryCast(obj, AttributeReference)

                                If ar IsNot Nothing Then
                                    ' ... to see whether it has
                                    ' the tag we're after
                                    If ar.Tag.ToUpper() = attbName Then

                                        ' Windows.MessageBox.Show(attbName & " attbName")
                                        ' If so, update the value
                                        ' and increment the counter

                                        ar.UpgradeOpen()
                                        ' Windows.MessageBox.Show(attbValue & " attbValue")
                                        ar.TextString = attbValue
                                        ar.DowngradeOpen()

                                        changedCount += 1

                                    End If

                                End If

                            Next
                        End If

                        ' Recurse for nested blocks
                        changedCount += UpdateAttributesInBlock(br.BlockTableRecord, blockName, attbName, attbValue)

                    End If

                End If
            Next

            tr.Commit()
        End Using

        Return changedCount

    End Function

    Private Sub ComboBox5_textChanged(sender As Object, e As EventArgs) Handles ComboBox5.TextChanged

        Button1.Enabled = True
        Button1.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Module_Arcxis_TB.ATB_UserSelect = ""
        Module_Arcxis_TB.ATB_UserSelect2 = ""
        DeleteWaterMark()
        Me.Close()

    End Sub

    Private Sub PDF24x36_Click(sender As Object, e As EventArgs) Handles PDF24x36.Click


        PageSetUp24x36()

        'Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            'Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                Dim plotTransPrev = Application.GetSystemVariable("PLOTTRANSPARENCYOVERRIDE")
                Application.SetSystemVariable("BackGroundPlot", 0)
                Application.SetSystemVariable("CMDDIA", 0)
                Application.SetSystemVariable("FILEDIA", 0)
                'acDoc.SendStringToExecute("-updatefields all 0 ", True, False, False)
                Try
                    ' 1) Ensure output folder exists
                    Dim outputDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\"
                    Dim DWGname As String = DirectCast(Application.GetSystemVariable("DWGNAME"), String)
                    DWGname = DWGname.Remove(DWGname.Length - 4)
                    Dim pdfFile As String = outputDir & DWGname & ".pdf"
                    Dim dsdFile As String = outputDir & DWGname & ".dsd"

                    Dim dwgprefix As String = Application.GetSystemVariable("dwgprefix")
                    Dim DWGnm As String = Application.GetSystemVariable("dwgName")
                    Dim dwgFile As String = dwgprefix & DWGnm

                    If File.Exists(dsdFile) Then
                        File.Delete(dsdFile)
                    End If

                    If File.Exists(pdfFile) Then
                        File.Delete(pdfFile)
                    End If

                    ' 4) Build DSD entries
                    Dim dsd As New DsdData()
                    Dim dsdEntries As New DsdEntryCollection()

                    For Each lay As Object In ListBox1.SelectedItems
                        Dim title As String = DWGnm.Remove(DWGnm.Length - 4) & "-" & lay

                        Dim de As New DsdEntry()
                        de.DwgName = dwgFile
                        de.Layout = lay           ' layout name
                        de.Title = title
                        de.Nps = "Arcxis24x36"            ' named page setup (ensure it exists)
                        de.NpsSourceDwg = dwgFile
                        dsdEntries.Add(de)
                    Next

                    dsd.SetDsdEntryCollection(dsdEntries)
                    dsd.SheetType = SheetType.MultiPdf
                    dsd.NoOfCopies = 1
                    dsd.IsHomogeneous = True
                    dsd.ProjectPath = outputDir
                    dsd.DestinationName = pdfFile   ' belt+braces
                    dsd.Dwf3dOptions.PublishWithMaterials = True
                    dsd.Dwf3dOptions.GroupByXrefHierarchy = True

                    ' Suppress prompts via API flags (some builds honor these)
                    dsd.SetUnrecognizedData("PromptForDwfName", "FALSE")
                    dsd.SetUnrecognizedData("PromptForName", "FALSE")

                    ' 5) Write DSD to disk, then hard-edit the text (covers all variants)
                    If File.Exists(dsdFile) Then File.Delete(dsdFile)
                    dsd.WriteDsd(dsdFile)

                    Dim text As String = File.ReadAllText(dsdFile)

                    ' Force no prompt + correct output + PDF type
                    Dim ensure As New List(Of String) From {
                "PromptForDwfName=False",
                "PromptForName=False",
                "PwdProtectPublishedDWF=False",
                "IncludeHyperlinks=TRUE",
                "IncludeLayer=TRUE",
                "Type=6",                             ' 6 = PDF in many DSDs
                "OutDir=" & outputDir.Replace("\", "\\"),
                "Dst=" & pdfFile.Replace("\", "\\")   ' some DSDs use Dst for final file
            }

                    ' Normalize common variants then inject ours
                    text = text.Replace("PromptForDwfName=True", "PromptForDwfName=False")
                    text = text.Replace("PromptForName=True", "PromptForName=False")
                    text = text.Replace("Type=3", "Type=6") ' DWF->PDF if needed

                    ' Guarantee we have OutDir and Dst lines (add if missing)
                    If Not text.Contains(vbCrLf & "OutDir=") Then text &= vbCrLf & "OutDir=" & outputDir
                    If Not text.Contains(vbCrLf & "Dst=") Then text &= vbCrLf & "Dst=" & pdfFile

                    ' Re-apply our ensure list to be certain
                    For Each line In ensure
                        Dim key = line.Split("="c)(0)
                        Dim idx = text.IndexOf(key & "=", StringComparison.OrdinalIgnoreCase)
                        If idx >= 0 Then
                            ' replace the whole row
                            Dim rowEnd = text.IndexOfAny({ControlChars.Cr, ControlChars.Lf}, idx)
                            If rowEnd < 0 Then rowEnd = text.Length
                            text = text.Remove(idx, rowEnd - idx).Insert(idx, line)
                        Else
                            text &= vbCrLf & line
                        End If
                    Next

                    File.WriteAllText(dsdFile, text)

                    ' Re-read into DsdData so Publisher uses our edits
                    dsd.ReadDsd(dsdFile)

                    ' 6) Pick PDF PC3 (or pass Nothing)
                    Dim pc As PlotConfig = Nothing
                    Try
                        pc = PlotConfigManager.SetCurrentConfig("AutoCAD PDF (General Documentation) - Brics.pc3")
                    Catch
                        ' ignore; Publisher can still use per-layout NPS
                    End Try

                    ' 7) Publish silently
                    Application.Publisher.PublishExecute(dsd, pc)

                    ' Cleanup
                    If File.Exists(dsdFile) Then File.Delete(dsdFile)

                    acTrans.Commit()

                Finally
                    ' Restore system vars
                    Application.SetSystemVariable("BackGroundPlot", bgPrev)
                    Application.SetSystemVariable("CMDDIA", cmdPrev)
                    Application.SetSystemVariable("FILEDIA", fileDiaPrev)
                End Try

                ' Save the changes made
                acTrans.Commit()

            End Using

        End Using

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

        TextBox39.Text = ""
        TextBox40.Text = ""
        TextBox41.Text = ""

        Dim objPrint As New System.Drawing.Printing.PrinterSettings
        Dim printerformat As System.Drawing.Printing.PaperSize

        For Each printer As String In PrinterSettings.InstalledPrinters

            If printer = ComboBox6.Text Then

                Dim PrinterObj As New System.Drawing.Printing.PrinterSettings()
                PrinterObj.PrinterName = printer
                For Each printerformat In PrinterObj.PaperSizes()

                    If printerformat.PaperName.Contains("Letter") OrElse printerformat.PaperName.Contains("8.5x11") Then

                        TextBox39.Text = printerformat.PaperName

                    End If

                    If printerformat.PaperName.Contains("11") OrElse printerformat.PaperName.Contains("11x17") OrElse printerformat.PaperName.Contains("11 x 17") OrElse printerformat.PaperName.Contains("11"" x 17""") OrElse printerformat.PaperName.Contains("11""x17""") OrElse printerformat.PaperName.Contains("Tabloid") Then
                        TextBox40.Text = printerformat.PaperName

                    End If

                    If printerformat.PaperName.Contains("24x36") OrElse printerformat.PaperName.Contains("24 x 36") OrElse printerformat.PaperName.Contains("24"" x 36""") OrElse printerformat.PaperName.Contains("24""x36""") Then
                        TextBox41.Text = printerformat.PaperName

                    End If


                Next

            End If

        Next

        If Trim(TextBox39.Text).Length = 0 Then

            TextBox39.Text = "Paper Size not Available"

        End If

        If Trim(TextBox40.Text).Length = 0 Then

            TextBox40.Text = "Paper Size not Available"

        End If

        If Trim(TextBox41.Text).Length = 0 Then

            TextBox41.Text = "Paper Size not Available"

        End If

    End Sub
    Private Sub PlotPDF11x17_Click(sender As Object, e As EventArgs) Handles PlotPDF11x17.Click

        PageSetUp11x17()

        'Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            'Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                Dim plotTransPrev = Application.GetSystemVariable("PLOTTRANSPARENCYOVERRIDE")
                Application.SetSystemVariable("BackGroundPlot", 0)
                Application.SetSystemVariable("CMDDIA", 0)
                Application.SetSystemVariable("FILEDIA", 0)
                'acDoc.SendStringToExecute("-updatefields all 0 ", True, False, False)
                Try
                    ' 1) Ensure output folder exists
                    Dim outputDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\"
                    Dim DWGname As String = DirectCast(Application.GetSystemVariable("DWGNAME"), String)
                    DWGname = DWGname.Remove(DWGname.Length - 4)
                    Dim pdfFile As String = outputDir & DWGname & ".pdf"
                    Dim dsdFile As String = outputDir & DWGname & ".dsd"

                    Dim dwgprefix As String = Application.GetSystemVariable("dwgprefix")
                    Dim DWGnm As String = Application.GetSystemVariable("dwgName")
                    Dim dwgFile As String = dwgprefix & DWGnm

                    If File.Exists(dsdFile) Then
                        File.Delete(dsdFile)
                    End If

                    If File.Exists(pdfFile) Then
                        File.Delete(pdfFile)
                    End If

                    ' 4) Build DSD entries
                    Dim dsd As New DsdData()
                    Dim dsdEntries As New DsdEntryCollection()

                    For Each lay As Object In ListBox1.SelectedItems
                        Dim title As String = DWGnm.Remove(DWGnm.Length - 4) & "-" & lay

                        Dim de As New DsdEntry()
                        de.DwgName = dwgFile
                        de.Layout = lay           ' layout name
                        de.Title = title
                        de.Nps = "Arcxis"            ' named page setup (ensure it exists)
                        de.NpsSourceDwg = dwgFile
                        dsdEntries.Add(de)
                    Next

                    dsd.SetDsdEntryCollection(dsdEntries)
                    dsd.SheetType = SheetType.MultiPdf
                    dsd.NoOfCopies = 1
                    dsd.IsHomogeneous = True
                    dsd.ProjectPath = outputDir
                    dsd.DestinationName = pdfFile   ' belt+braces
                    dsd.Dwf3dOptions.PublishWithMaterials = True
                    dsd.Dwf3dOptions.GroupByXrefHierarchy = True

                    ' Suppress prompts via API flags (some builds honor these)
                    dsd.SetUnrecognizedData("PromptForDwfName", "FALSE")
                    dsd.SetUnrecognizedData("PromptForName", "FALSE")

                    ' 5) Write DSD to disk, then hard-edit the text (covers all variants)
                    If File.Exists(dsdFile) Then File.Delete(dsdFile)
                    dsd.WriteDsd(dsdFile)

                    Dim text As String = File.ReadAllText(dsdFile)

                    ' Force no prompt + correct output + PDF type
                    Dim ensure As New List(Of String) From {
                "PromptForDwfName=False",
                "PromptForName=False",
                "PwdProtectPublishedDWF=False",
                "IncludeHyperlinks=TRUE",
                "IncludeLayer=TRUE",
                "Type=6",                             ' 6 = PDF in many DSDs
                "OutDir=" & outputDir.Replace("\", "\\"),
                "Dst=" & pdfFile.Replace("\", "\\")   ' some DSDs use Dst for final file
            }

                    ' Normalize common variants then inject ours
                    text = text.Replace("PromptForDwfName=True", "PromptForDwfName=False")
                    text = text.Replace("PromptForName=True", "PromptForName=False")
                    text = text.Replace("Type=3", "Type=6") ' DWF->PDF if needed

                    ' Guarantee we have OutDir and Dst lines (add if missing)
                    If Not text.Contains(vbCrLf & "OutDir=") Then text &= vbCrLf & "OutDir=" & outputDir
                    If Not text.Contains(vbCrLf & "Dst=") Then text &= vbCrLf & "Dst=" & pdfFile

                    ' Re-apply our ensure list to be certain
                    For Each line In ensure
                        Dim key = line.Split("="c)(0)
                        Dim idx = text.IndexOf(key & "=", StringComparison.OrdinalIgnoreCase)
                        If idx >= 0 Then
                            ' replace the whole row
                            Dim rowEnd = text.IndexOfAny({ControlChars.Cr, ControlChars.Lf}, idx)
                            If rowEnd < 0 Then rowEnd = text.Length
                            text = text.Remove(idx, rowEnd - idx).Insert(idx, line)
                        Else
                            text &= vbCrLf & line
                        End If
                    Next

                    File.WriteAllText(dsdFile, text)

                    ' Re-read into DsdData so Publisher uses our edits
                    dsd.ReadDsd(dsdFile)

                    ' 6) Pick PDF PC3 (or pass Nothing)
                    Dim pc As PlotConfig = Nothing
                    Try
                        pc = PlotConfigManager.SetCurrentConfig("ARCXIS - DWG To PDF - Brics.pc3")
                    Catch
                        ' ignore; Publisher can still use per-layout NPS
                    End Try

                    ' 7) Publish silently
                    Application.Publisher.PublishExecute(dsd, pc)

                    ' Cleanup
                    If File.Exists(dsdFile) Then File.Delete(dsdFile)

                    acTrans.Commit()

                Finally
                    ' Restore system vars
                    Application.SetSystemVariable("BackGroundPlot", bgPrev)
                    Application.SetSystemVariable("CMDDIA", cmdPrev)
                    Application.SetSystemVariable("FILEDIA", fileDiaPrev)
                End Try

                ' Save the changes made
                acTrans.Commit()

            End Using

        End Using

    End Sub

    Private Sub Plot_Click(sender As Object, e As EventArgs) Handles Plot.Click

        Module_Arcxis_TB.ATB_PlotterName = ComboBox6.Text

        If RadioButton1.Checked = True Then

            Module_Arcxis_TB.ATB_PlotSize = TextBox39.Text

        ElseIf RadioButton2.Checked = True Then


            If ATB_PlotterName.Contains("SAVIN") Then

                Module_Arcxis_TB.ATB_PlotSize = "11x17"

            Else

                Module_Arcxis_TB.ATB_PlotSize = TextBox40.Text

            End If



        ElseIf RadioButton3.Checked = True Then

            Module_Arcxis_TB.ATB_PlotSize = TextBox41.Text

        ElseIf RadioButton4.Checked = True Then


            Module_Arcxis_TB.ATB_PlotSize = ComboBox7.Text

        End If

        If RadioButton9.Checked = True Then

            Module_Arcxis_TB.ATB_PlotStyleName = "Arcxis.ctb"

        ElseIf RadioButton10.Checked = True Then

            Module_Arcxis_TB.ATB_PlotStyleName = "Arcxis.ctb"


        End If


        If RadioButton5.Checked = True Then

            Module_Arcxis_TB.ATB_WaterMarkText = RadioButton5.Text.ToUpper()

        ElseIf RadioButton6.Checked = True Then

            Module_Arcxis_TB.ATB_WaterMarkText = RadioButton6.Text.ToUpper()

        ElseIf RadioButton8.Checked = True Then

            Module_Arcxis_TB.ATB_WaterMarkText = RadioButton8.Text.ToUpper()


        End If


        If RadioButton7.Checked <> True Then

            Dim psize As String = "11x17"

            WaterMark(psize)

        End If

        'Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim CurVar As Object
        'Get the BACKGROUNDPLOT SYSTEM VARIABLE
        CurVar = Application.GetSystemVariable("BackGroundPlot")
        'SET BACKGROUNDPLOT SYSTEM VARIABLE TO ZERO
        Application.SetSystemVariable("BackGroundPlot", 0)

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            'Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim lays As DBDictionary = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead)

                'Step through layout list
                For Each item As DBDictionaryEntry In lays

                    Dim acLayout As Layout = CType(item.Value.GetObject(OpenMode.ForRead), Layout)
                    Dim acLayName As String = acLayout.LayoutName

                    For Each itemlist As Object In ListBox1.SelectedItems

                        'Check to see if Layout is PaperSpace		
                        If acLayName = itemlist Then
                            LayoutManager.Current.CurrentLayout = acLayName

                            'Get the PlotInfo from the layout
                            Dim acPlInfo As PlotInfo = New PlotInfo()
                            acPlInfo.Layout = acLayout.ObjectId

                            'Get a copy of the PlotSettings from the layout
                            Dim acPlSet As PlotSettings = New PlotSettings(acLayout.ModelType)
                            acPlSet.CopyFrom(acLayout)

                            'Update the PlotSettings object
                            Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                            '' Set the plot device to use
                            ' acPlSetVdr.SetPlotConfigurationName(acPlSet, "HP6700.pc3", "Letter")
                            acPlSetVdr.SetPlotConfigurationName(acPlSet, ATB_PlotterName, ATB_PlotSize)

                            ' Set the plot type
                            acPlSetVdr.SetPlotType(acPlSet, PlotType.Extents)

                            ' Set the plot Centered
                            acPlSetVdr.SetPlotCentered(acPlSet, True)

                            ' Set the plot scale
                            acPlSetVdr.SetUseStandardScale(acPlSet, True)
                            acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit)
                            acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches)
                            acPlSet.ScaleLineweights = False

                            ' Specify if plot styles should be displayed on the layout
                            acPlSet.ShowPlotStyles = True

                            ' Rebuild plotter, plot style, and canonical media lists 
                            ' (must be called before setting the plot style)
                            acPlSetVdr.RefreshLists(acPlSet)

                            ' Specify the shaded viewport options
                            acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed

                            acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                            ' Specify the plot options
                            acPlSet.PrintLineweights = True
                            acPlSet.PlotTransparency = False
                            acPlSet.PlotPlotStyles = True
                            acPlSet.DrawViewportsFirst = False


                            ' Specify the plot orientation
                            acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090)

                            ' Specify the plot orientation
                            acPlSetVdr.SetCurrentStyleSheet(acPlSet, PlotStyleName)

                            ' Zoom to show the whole paper
                            acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)

                            'Set the plot info as an override since it will not be saved back to the layout
                            acPlInfo.OverrideSettings = acPlSet

                            'Validate the plot info
                            Dim acPlInfoVdr As PlotInfoValidator = New PlotInfoValidator()
                            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                            acPlInfoVdr.Validate(acPlInfo)


                            'Check to see if a plot is already in progress
                            If PlotFactory.ProcessPlotState = ProcessPlotState.NotPlotting Then

                                Using acPlEng As PlotEngine = PlotFactory.CreatePublishEngine()

                                    'Track the plot progress with a Progress dialog
                                    Dim acPlProgDlg As PlotProgressDialog = New PlotProgressDialog(False, 1, True)

                                    Using (acPlProgDlg)
                                        'Define the status messages to display when plotting starts
                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.DialogTitle) = "Plot Progress"

                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelJobButtonMessage) = "Cancel Job"

                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage) = "Cancel Sheet"

                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetSetProgressCaption) = "Sheet Set Progress"

                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetProgressCaption) = "Sheet Progress"

                                        'Set the plot progress range
                                        acPlProgDlg.LowerPlotProgressRange = 0
                                        acPlProgDlg.UpperPlotProgressRange = 100
                                        acPlProgDlg.PlotProgressPos = 0

                                        'Display the Progress dialog
                                        acPlProgDlg.OnBeginPlot()
                                        acPlProgDlg.IsVisible = True

                                        'Start to plot the layout
                                        acPlEng.BeginPlot(acPlProgDlg, Nothing)

                                        'Define the plot output
                                        acPlEng.BeginDocument(acPlInfo, acDoc.Name, Nothing, 1, False, "c:\myplot")

                                        'Display information about the current plot
                                        acPlProgDlg.PlotMsgString(PlotMessageIndex.Status) = "Plotting: " & acDoc.Name & " - " & acLayout.LayoutName

                                        'Set the sheet progress range
                                        acPlProgDlg.OnBeginSheet()
                                        acPlProgDlg.LowerSheetProgressRange = 0
                                        acPlProgDlg.UpperSheetProgressRange = 100
                                        acPlProgDlg.SheetProgressPos = 0

                                        'Plot the first sheet/layout
                                        Dim acPlPageInfo As PlotPageInfo = New PlotPageInfo()
                                        acPlEng.BeginPage(acPlPageInfo, acPlInfo, True, Nothing)

                                        acPlEng.BeginGenerateGraphics(Nothing)
                                        acPlEng.EndGenerateGraphics(Nothing)

                                        'Finish plotting the sheet/layout
                                        acPlEng.EndPage(Nothing)
                                        acPlProgDlg.SheetProgressPos = 100
                                        acPlProgDlg.OnEndSheet()

                                        'Finish plotting the document
                                        acPlEng.EndDocument(Nothing)

                                        'Finish the plot
                                        acPlProgDlg.PlotProgressPos = 100
                                        acPlProgDlg.OnEndPlot()
                                        acPlEng.EndPlot(Nothing)

                                    End Using
                                End Using
                            End If

                            '''''''''''''''''''''''''''''''''''''''''''''''''

                        End If

                    Next
                Next

            End Using

        End Using

        'REVERT BACK TO THE ORIGINAL BACKGROUNDPLOT SYSTEM VARIABLE
        Application.SetSystemVariable("BackGroundPlot", CurVar)


    End Sub
    Private Sub PageSetUp()

        For Each layoutname As Object In ListBox1.SelectedItems
            ' Get the current document and database, and start a transaction
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim plSets As DBDictionary = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead)
                Dim vStyles As DBDictionary = acTrans.GetObject(acCurDb.VisualStyleDictionaryId, OpenMode.ForRead)

                Dim acPlSet As PlotSettings
                Dim createNew As Boolean = False

                ' Reference the Layout Manager
                Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                ' Get the current layout and output its name in the Command Line window
                ' Dim acLayout As Layout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead)

                Dim layoutManager__1 As LayoutManager = LayoutManager.Current
                Dim layoutId As ObjectId = layoutManager__1.GetLayoutId(layoutname)
                'Windows.MessageBox.Show(layoutId.ToString & "  llayoutId")
                Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                ' Check to see if the page setup exists
                If plSets.Contains("ARCXIS DWG TO PDF (11x17)") = False Then
                    createNew = True

                    ' Create a new PlotSettings object: 
                    '    True - model space, False - named layout
                    acPlSet = New PlotSettings(acLayout.ModelType)
                    acPlSet.CopyFrom(acLayout)

                    acPlSet.PlotSettingsName = "ARCXIS DWG TO PDF (11x17)"
                    acPlSet.AddToPlotSettingsDictionary(acCurDb)
                    acTrans.AddNewlyCreatedDBObject(acPlSet, True)
                Else
                    acPlSet = plSets.GetAt("ARCXIS DWG TO PDF (11x17)").GetObject(OpenMode.ForWrite)
                End If

                Try
                    Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                    ' Set the Plotter and page size
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, "ARCXIS - DWG To PDF - Brics.pc3", "ANSI_full_bleed_B_(17.00_x_11.00_Inches)")

                    ' Set to plot to the current display
                    'If accLayout.ModelType = False Then
                    'acPlSetVdr.SetPlotType(acPlSet, PlotType.Layout)
                    'Else
                    acPlSetVdr.SetPlotType(acPlSet, PlotType.Extents)

                    'acPlSetVdr.SetPlotCentered(acPlSet, True)
                    'End If

                    ' Set the plot offset
                    'acPlSetVdr.SetPlotOrigin(acPlSet, New Point2d(0, 0))
                    acPlSetVdr.SetPlotCentered(acPlSet, True)
                    ' Set the plot scale
                    acPlSetVdr.SetUseStandardScale(acPlSet, True)
                    acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.StdScale1To1)
                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches)
                    acPlSet.ScaleLineweights = True

                    ' Specify if plot styles should be displayed on the layout
                    acPlSet.ShowPlotStyles = True

                    ' Rebuild plotter, plot style, and canonical media lists 
                    ' (must be called before setting the plot style)
                    acPlSetVdr.RefreshLists(acPlSet)

                    ' Specify the shaded viewport options
                    acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed

                    acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                    ' Specify the plot options
                    acPlSet.PrintLineweights = True
                    acPlSet.PlotTransparency = False
                    acPlSet.PlotPlotStyles = True
                    acPlSet.DrawViewportsFirst = False

                    ' Use only on named layouts - Hide paperspace objects option
                    ' plSet.PlotHidden = True

                    ' Specify the plot orientation
                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000)

                    ' Set the plot style
                    'If acCurDb.PlotStyleMode = True Then
                    acPlSetVdr.SetCurrentStyleSheet(acPlSet, "ARCXIS.ctb")
                    'Else
                    'acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.stb")
                    'End If

                    ' Zoom to show the whole paper
                    acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)
                Catch es As Exception
                    MsgBox(es.Message)
                End Try

                ' Save the changes made
                acTrans.Commit()

                If createNew = True Then
                    acPlSet.Dispose()
                End If
            End Using

        Next

    End Sub

    Private Sub PageSetUp11x17()

        If RadioButton11.Checked = True Or RadioButton12.Checked = True Then
            Module_Arcxis_TB.ATB_PlotStyleName = "DPIS-11x17.ctb"
        Else
            Module_Arcxis_TB.ATB_PlotStyleName = "ARCXIS.ctb"
        End If

        Module_Arcxis_TB.ATB_PlotSetUpName = "Arcxis"

        For Each layoutname As Object In ListBox1.SelectedItems
            ' Get the current document and database, and start a transaction
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acLckDoc As DocumentLock = acDoc.LockDocument()
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim plSets As DBDictionary = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead)
                    Dim vStyles As DBDictionary = acTrans.GetObject(acCurDb.VisualStyleDictionaryId, OpenMode.ForRead)

                    Dim acPlSet As PlotSettings
                    Dim createNew As Boolean = False

                    ' Reference the Layout Manager
                    Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                    ' Get the current layout and output its name in the Command Line window
                    ' Dim acLayout As Layout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead)

                    Dim layoutManager__1 As LayoutManager = LayoutManager.Current
                    Dim layoutId As ObjectId = layoutManager__1.GetLayoutId(layoutname)
                    'Windows.MessageBox.Show(layoutId.ToString & "  llayoutId")
                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                    ' Check to see if the page setup exists
                    If plSets.Contains(ATB_PlotSetUpName) = False Then
                        createNew = True

                        ' Create a new PlotSettings object: 
                        '    True - model space, False - named layout
                        acPlSet = New PlotSettings(acLayout.ModelType)
                        acPlSet.CopyFrom(acLayout)

                        acPlSet.PlotSettingsName = ATB_PlotSetUpName
                        acPlSet.AddToPlotSettingsDictionary(acCurDb)
                        acTrans.AddNewlyCreatedDBObject(acPlSet, True)
                    Else
                        acPlSet = plSets.GetAt(ATB_PlotSetUpName).GetObject(OpenMode.ForWrite)
                    End If

                    Try
                        Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                        ' Set the Plotter and page size
                        acPlSetVdr.SetPlotConfigurationName(acPlSet, "ARCXIS - DWG To PDF - Brics.pc3", "ANSI_full_bleed_B_(17.00_x_11.00_Inches)")

                        ' Set to plot to the current display
                        'If accLayout.ModelType = False Then
                        'acPlSetVdr.SetPlotType(acPlSet, PlotType.Layout)
                        'Else
                        acPlSetVdr.SetPlotType(acPlSet, PlotType.Extents)

                        'acPlSetVdr.SetPlotCentered(acPlSet, True)
                        'End If

                        ' Set the plot offset
                        'acPlSetVdr.SetPlotOrigin(acPlSet, New Point2d(0, 0))
                        acPlSetVdr.SetPlotCentered(acPlSet, True)
                        ' Set the plot scale
                        acPlSetVdr.SetUseStandardScale(acPlSet, True)
                        acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.StdScale1To1)
                        acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches)
                        acPlSet.ScaleLineweights = True

                        ' Specify if plot styles should be displayed on the layout
                        acPlSet.ShowPlotStyles = True

                        ' Rebuild plotter, plot style, and canonical media lists 
                        ' (must be called before setting the plot style)
                        acPlSetVdr.RefreshLists(acPlSet)

                        ' Specify the shaded viewport options
                        acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed

                        acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                        ' Specify the plot options
                        acPlSet.PrintLineweights = True
                        acPlSet.PlotTransparency = False
                        acPlSet.PlotPlotStyles = True
                        acPlSet.DrawViewportsFirst = False

                        ' Use only on named layouts - Hide paperspace objects option
                        ' plSet.PlotHidden = True

                        ' Specify the plot orientation
                        acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000)

                        ' Set the plot style
                        'If acCurDb.PlotStyleMode = True Then
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, ATB_PlotStyleName)
                        'Else
                        'acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.stb")
                        'End If

                        ' Zoom to show the whole paper
                        acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)

                    Catch es As Exception
                        MsgBox(es.Message)
                    End Try

                    ' Save the changes made
                    acTrans.Commit()

                    If createNew = True Then
                        acPlSet.Dispose()
                    End If
                End Using

            End Using

        Next

    End Sub

    Private Sub PageSetUp24x36()

        If RadioButton11.Checked = True Or RadioButton12.Checked = True Then
            Module_Arcxis_TB.ATB_PlotStyleName = "DPIS-11x17.ctb"
        Else
            Module_Arcxis_TB.ATB_PlotStyleName = "ARCXIS.ctb"
        End If

        Module_Arcxis_TB.ATB_PlotSetUpName = "Arcxis24x36"

        For Each layoutname As Object In ListBox1.SelectedItems
            ' Get the current document and database, and start a transaction
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim plSets As DBDictionary = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead)
                    Dim vStyles As DBDictionary = acTrans.GetObject(acCurDb.VisualStyleDictionaryId, OpenMode.ForRead)

                    Dim acPlSet As PlotSettings
                    Dim createNew As Boolean = False

                    ' Reference the Layout Manager
                    Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                    ' Get the current layout and output its name in the Command Line window
                    ' Dim acLayout As Layout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead)

                    Dim layoutManager__1 As LayoutManager = LayoutManager.Current
                    Dim layoutId As ObjectId = layoutManager__1.GetLayoutId(layoutname)
                    'Windows.MessageBox.Show(layoutId.ToString & "  llayoutId")
                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                    ' Check to see if the page setup exists
                    If plSets.Contains(ATB_PlotSetUpName) = False Then
                        createNew = True

                        ' Create a new PlotSettings object: 
                        '    True - model space, False - named layout
                        acPlSet = New PlotSettings(acLayout.ModelType)
                        acPlSet.CopyFrom(acLayout)

                        acPlSet.PlotSettingsName = ATB_PlotSetUpName
                        acPlSet.AddToPlotSettingsDictionary(acCurDb)
                        acTrans.AddNewlyCreatedDBObject(acPlSet, True)
                    Else
                        acPlSet = plSets.GetAt(ATB_PlotSetUpName).GetObject(OpenMode.ForWrite)
                    End If

                    Try
                        Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current


                        'ARCXIS - DWG To PDF - Brics.pc3

                        ' Set the Plotter and page size
                        acPlSetVdr.SetPlotConfigurationName(acPlSet, "AutoCAD PDF (General Documentation) - Brics.pc3", "ANSI_full_bleed_D_(34.00_x_22.00_Inches)")

                        ' Set to plot to the current display
                        'If accLayout.ModelType = False Then
                        'acPlSetVdr.SetPlotType(acPlSet, PlotType.Layout)
                        'Else
                        acPlSetVdr.SetPlotType(acPlSet, PlotType.Extents)

                        'acPlSetVdr.SetPlotCentered(acPlSet, True)
                        'End If

                        ' Set the plot offset
                        'acPlSetVdr.SetPlotOrigin(acPlSet, New Point2d(0, 0))
                        acPlSetVdr.SetPlotCentered(acPlSet, True)
                        ' Set the plot scale
                        acPlSetVdr.SetUseStandardScale(acPlSet, True)
                        acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.StdScale1To1)
                        acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches)
                        acPlSet.ScaleLineweights = True

                        ' Specify if plot styles should be displayed on the layout
                        acPlSet.ShowPlotStyles = True

                        ' Rebuild plotter, plot style, and canonical media lists 
                        ' (must be called before setting the plot style)
                        acPlSetVdr.RefreshLists(acPlSet)

                        ' Specify the shaded viewport options
                        acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed

                        acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                        ' Specify the plot options
                        acPlSet.PrintLineweights = True
                        acPlSet.PlotTransparency = False
                        acPlSet.PlotPlotStyles = True
                        acPlSet.DrawViewportsFirst = False

                        ' Use only on named layouts - Hide paperspace objects option
                        ' plSet.PlotHidden = True

                        ' Specify the plot orientation
                        acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000)

                        ' Set the plot style
                        'If acCurDb.PlotStyleMode = True Then
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, ATB_PlotStyleName)
                        'Else
                        'acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.stb")
                        'End If

                        ' Zoom to show the whole paper
                        acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)
                    Catch es As Exception
                        MsgBox(es.Message)
                    End Try

                    ' Save the changes made
                    acTrans.Commit()

                    If createNew = True Then
                        acPlSet.Dispose()
                    End If
                End Using
            End Using

        Next

    End Sub


    Private Sub WaterMark(psize As String)

        If RadioButton5.Checked = True Then

            ATB_WaterMarkText = RadioButton5.Text.ToUpper()

        ElseIf RadioButton6.Checked = True Then

            ATB_WaterMarkText = RadioButton6.Text.ToUpper()

        ElseIf RadioButton8.Checked = True Then

            ATB_WaterMarkText = RadioButton8.Text.ToUpper()

        End If

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


                '' Open the Layer table for read
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                Dim sLayerName As String = "Arcxis-WaterMark"

                If acLyrTbl.Has(sLayerName) = False Then
                    Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                    '' Assign the layer the ACI color 1 and a name
                    acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 244)
                    acLyrTblRec.Name = sLayerName

                    '' Upgrade the Layer table for write
                    acLyrTbl.UpgradeOpen()

                    '' Append the new layer to the Layer table and the transaction
                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                    '' Set the layer Center current
                    acCurDb.Clayer = acLyrTbl(sLayerName)


                    'ElseIf acLyrTbl.Has(sLayerName) = True Then
                    '' Set the layer Center current
                    'acCurDb.Clayer = acLyrTbl(sLayerName)


                End If

                Dim tsTbl As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)

                'Dim str As TextStyleTableRecord = New TextStyleTableRecord()

                Dim textstylename As String = "Arcxis-WaterMark"

                If tsTbl.Has(textstylename) Then

                    Dim TxtStyleVar As Object
                    'Get the TextStylye SYSTEM VARIABLE
                    TxtStyleVar = Application.GetSystemVariable("TextStyle")
                    'SET TextStyle SYSTEM VARIABLE
                    Application.SetSystemVariable("TextStyle", "Arcxis-WaterMark")

                Else

                    Dim st As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)

                    Dim strr As TextStyleTableRecord = New TextStyleTableRecord()

                    strr.Name = "Arcxis-WaterMark"

                    st.Add(strr)

                    strr.Font = New FontDescriptor("Swis721 Ex BT", True, True, Nothing, Nothing)

                    acTrans.AddNewlyCreatedDBObject(strr, True)

                    'make as current

                    Dim TxtStyleVar As Object
                    'Get the TextStylye SYSTEM VARIABLE
                    TxtStyleVar = Application.GetSystemVariable("TextStyle")
                    'SET TextStyle SYSTEM VARIABLE
                    Application.SetSystemVariable("TextStyle", "Arcxis-WaterMark")


                End If

                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                For Each layoutname As Object In ListBox1.SelectedItems

                    'Windows.MessageBox.Show(layoutname & "  layoutname")

                    '' Open the Block table record Model space for write
                    'Dim acBlkTblRec As BlockTableRecord
                    'acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                    Dim layoutManager__1 As LayoutManager = LayoutManager.Current
                    Dim layoutId As ObjectId = layoutManager__1.GetLayoutId(layoutname)
                    'Windows.MessageBox.Show(layoutId.ToString & "  llayoutId")
                    Dim layout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)
                    Dim acBlkTblRec As BlockTableRecord = Nothing
                    acBlkTblRec = DirectCast(acTrans.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite), BlockTableRecord)

                    '' Create a single-line text object
                    Dim acText As DBText = New DBText()
                    acText.SetDatabaseDefaults()
                    Dim minPt As Point2d = acCurDb.Limmin
                    Dim maxPt As Point2d = acCurDb.Limmax
                    Dim midPt As Point2d = minPt.Add(minPt.GetVectorTo(maxPt) / 2)

                    'acText.AlignmentPoint = New Point3d(midPt.X, midPt.Y, 0)
                    'acText.AlignmentPoint = New Point3d(0.5, 0.5, 0)
                    Dim trpy As Transparency
                    trpy = New Transparency(CByte(80))
                    acText.Transparency = trpy

                    If ATB_WaterMarkText = "FOR REVIEW ONLY" OrElse ATB_WaterMarkText = "MASTER SET" Then

                        If psize = "11x17" Then
                            acText.Height = 1.125

                        ElseIf psize = "24x36" Then

                            acText.Height = 2.125

                        End If

                    Else

                        If psize = "11x17" Then
                            acText.Height = 1.0

                        ElseIf psize = "24x36" Then

                            acText.Height = 2.0

                        End If

                    End If

                    acText.TextString = ATB_WaterMarkText
                    acText.Layer = "Arcxis-WaterMark"
                    acText.Rotation = 0.523599
                    acText.Justify = AttachmentPoint.MiddleCenter
                    acText.ColorIndex = 254
                    'acText.Position = New Teigha.Geometry.Point3d(2, 2, 0)
                    acText.HorizontalMode = TextHorizontalMode.TextCenter
                    acText.VerticalMode = TextVerticalMode.TextVerticalMid

                    If psize = "11x17" Then

                        acText.AlignmentPoint = New Teigha.Geometry.Point3d(8.5, 5.5, 0)

                    ElseIf psize = "24x36" Then

                        acText.AlignmentPoint = New Teigha.Geometry.Point3d(18, 12, 0)

                    End If

                    acText.WidthFactor = 0.95

                    acBlkTblRec.AppendEntity(acText)
                    acTrans.AddNewlyCreatedDBObject(acText, True)

                    Dim drawOrder As DrawOrderTable = TryCast(acTrans.GetObject(acBlkTblRec.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)

                    Dim ids As New ObjectIdCollection()

                    ids.Add(acText.ObjectId)

                    'move the selected entity so that entity is 
                    'drawn in the beginning of the draw order.

                    drawOrder.MoveToBottom(ids)

                Next

                'set focus to the editor
                ' Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                '' Save the changes and dispose of the transaction
                acTrans.Commit()
            End Using

        End Using

    End Sub

    Private Sub DeleteWaterMark()

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


                '' Create a TypedValue array to define the filter criteria
                Dim acTypValAr(2) As TypedValue
                acTypValAr.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
                acTypValAr.SetValue(New TypedValue(DxfCode.LayerName, "Arcxis-WaterMark"), 1)
                acTypValAr.SetValue(New TypedValue(DxfCode.Operator, "or>"), 2)

                Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

                Dim result As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.SelectAll(acSelFtr)
                If (result.Status = PromptStatus.OK) Then
                    ' There are selected entities
                    ' Put your command using pickfirst set code here
                    Dim acSSet As SelectionSet = result.Value

                    '' Step through the objects in the selection set
                    For Each acSSObj As SelectedObject In acSSet
                        '' Check to make sure a valid SelectedObject object was returned
                        'If Not IsDBNull(acSSObj) Then

                        Dim typ As String = acSSObj.ObjectId.ObjectClass.Name()

                        If typ = "AcDbText" Then
                            Dim acEnt As DBText = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite, False, True)
                            'Dim textEntity As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)

                            Dim lay As String = acEnt.Layer

                            If lay = "Arcxis-WaterMark" Then

                                acEnt.Erase()

                            End If

                        End If

                    Next


                    ''''''''''''''''''''''''''''''
                End If
                acTrans.Commit()

            End Using
        End Using

    End Sub

    Private Sub UpdateDesignType(tbname As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' get the working Database
        Dim db = HostApplicationServices.WorkingDatabase

        Using acLckDoc As DocumentLock = doc.LockDocument()

            ' start a transaction
            Using tr = db.TransactionManager.StartTransaction()
                ' open the block table
                Dim bt = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                ' check if the block table contains the block to rename
                'If bt.Has(oldName) Then
                ' check if the block table already contains a block named as the new name
                If bt.Has(tbname) Then

                    ' open the block definition
                    Dim btr = DirectCast(tr.GetObject(bt(tbname), OpenMode.ForWrite), BlockTableRecord)

                    Dim TBBlockRef As BlockReference

                    For Each id In btr.GetBlockReferenceIds(False, False)

                        Dim oAcadEntity As Entity = tr.GetObject(id, OpenMode.ForWrite)

                        If TypeOf oAcadEntity Is BlockReference Then

                            TBBlockRef = CType(oAcadEntity, BlockReference)


                            For Each attId As ObjectId In TBBlockRef.AttributeCollection

                                ' Open the attribute reference
                                Dim attref As AttributeReference = DirectCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)

                                'attref.SetAttributeFromBlock(attDef, blkRef.BlockTransform)

                                Dim tagvalue As String = attref.Tag

                                If tagvalue.Contains("OPTIONAL") Then
                                    ' Windows.MessageBox.Show(tagvalue & " tagvalue at optional line")
                                    attref.TextString = ComboBox2.Text.ToUpper()
                                    ' Windows.MessageBox.Show(ComboBox2.Text & " ComboBox2.Text")

                                ElseIf tagvalue = "DT1" Then

                                    If ComboBox2.Text.ToUpper() = "DESIGN 1" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 1 & 2" Then

                                        attref.TextString = "X"

                                    Else

                                        attref.TextString = ""

                                    End If

                                ElseIf tagvalue = "DT2" Then

                                    If ComboBox2.Text.ToUpper() = "DESIGN 2" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" Then

                                        attref.TextString = "X"

                                    Else

                                        attref.TextString = ""

                                    End If

                                ElseIf tagvalue = "DT3" Then

                                    If ComboBox2.Text.ToUpper() = "DESIGN 3" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 3 & 4" Then

                                        attref.TextString = "X"

                                    Else

                                        attref.TextString = ""

                                    End If

                                End If

                            Next

                        End If

                    Next

                    ' Update existing block references
                    For Each objID As ObjectId In btr.GetBlockReferenceIds(False, True)
                        Dim acBlkRef As BlockReference = tr.GetObject(objID, OpenMode.ForWrite)
                        acBlkRef.RecordGraphicsModified(True)

                    Next


                End If

                tr.Commit()

            End Using

        End Using

    End Sub


    Private Sub RadioButton7_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

        DeleteWaterMark()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.SystemColors.Control
        Dim frm As New Form_Arcxis_TB1


        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim Selectedlayouts As New ArrayList

        For Each blicky In ListBox1.SelectedItems

            Selectedlayouts.Add(blicky)

        Next
        '' Lock the new document
        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                For Each item As Object In ListBox1.SelectedItems

                    Dim blkname As String = "Arcxis Title Block"
                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, item)}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            frm.Update()

                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                                If brId.ObjectClass.Name = "AcDbBlockReference" Then

                                    Dim strutid As String = brId.ToString

                                    ' Open the block reference
                                    Dim BlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                                    Dim TblRec As BlockTableRecord = TryCast(acTrans.GetObject(BlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                                    Dim RevvblockName As String = TblRec.Name

                                    ' Iterate the attribute collection
                                    For Each attId As ObjectId In BlockRef.AttributeCollection

                                        Dim attidnum As String = attId.ToString

                                        ' Open the attribute reference
                                        Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)

                                        'attref.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                        attref.UpgradeOpen()

                                        Dim tagvalue As String = attref.Tag
                                        Dim dtype As String
                                        If ListBox1.SelectedItems.Count < 2 Then
                                            If tagvalue = "PLAN " Then

                                                attref.TextString = TextBox1.Text

                                            ElseIf tagvalue = "ELEVATION" Then

                                                attref.TextString = ComboBox1.Text
                                                'Me.ComboBox1.Items.Add(attref.TextString)
                                            ElseIf tagvalue.Contains("OPTIONAL") Then

                                                attref.TextString = ComboBox2.Text.ToUpper()

                                            ElseIf tagvalue = "FR-1" Then

                                                attref.TextString = TextBox7.Text

                                                Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                                                Dim aclayoutmanager As LayoutManager = LayoutManager.Current
                                                If layDict.Contains(item) Then
                                                    aclayoutmanager.RenameLayout(item, TextBox7.Text)
                                                    Selectedlayouts.Remove(item)
                                                    Selectedlayouts.Add(TextBox7.Text)
                                                End If


                                            End If

                                        End If
                                        Dim blk As String = "BLOCK"
                                        If tagvalue = "PLAN " Then

                                            'attref.TextString = TextBox1.Text

                                        ElseIf tagvalue = "ELEVATION" Then

                                            'attref.TextString = ComboBox1.Text
                                            'Me.ComboBox1.Items.Add(attref.TextString)

                                        ElseIf tagvalue.Contains("OPTIONAL") Then

                                            'attref.TextString = ComboBox2.Text.ToUpper()

                                        ElseIf tagvalue = "SUBDIVISION" Then

                                            attref.TextString = TextBox2.Text

                                        ElseIf tagvalue = "ADDRESS" Then

                                            attref.TextString = TextBox3.Text

                                        ElseIf tagvalue = "CUSTOMER'S NAME" Then

                                            attref.TextString = ComboBox3.Text

                                        ElseIf tagvalue = "PROJECT#" Then

                                            attref.TextString = TextBox4.Text

                                        ElseIf tagvalue = "1/8"" = 1'-0""" Then

                                            attref.TextString = ComboBox4.Text

                                        ElseIf tagvalue = "PLANDATE" Then

                                            attref.TextString = TextBox8.Text

                                        ElseIf tagvalue = "LOT" Then

                                            attref.TextString = TextBox38.Text

                                        ElseIf tagvalue = "BLK" Then

                                            attref.TextString = TextBox42.Text

                                        ElseIf tagvalue = "BLOCK" Then

                                            attref.TextString = TextBox42.Text

                                        ElseIf tagvalue = "SEC" Then

                                            attref.TextString = TextBox43.Text

                                        ElseIf tagvalue = "SECTION" Then

                                            attref.TextString = TextBox43.Text

                                        ElseIf tagvalue = "BUILDER" Then

                                            attref.TextString = TextBox5.Text

                                        ElseIf tagvalue = "BUILDER_ADDRESS" Then

                                            attref.TextString = TextBox6.Text

                                        ElseIf tagvalue = "CITY_STATE_ZIP" Then

                                            attref.TextString = TextBox44.Text

                                        ElseIf tagvalue = "OFFICE_PHONE" Then

                                            attref.TextString = TextBox46.Text

                                        ElseIf tagvalue = "ARCHITECTUAL_DATE" Then

                                            attref.TextString = TextBox45.Text
                                        ElseIf tagvalue = "STRUCTURAL_DATE" Then

                                            attref.TextString = TextBox47.Text


                                        ElseIf tagvalue.Contains("00") Then

                                            Dim destyp As String = attref.TextString

                                            ' Step through each object in the block table record
                                            'For Each objID As ObjectId In btr

                                            For Each objID As ObjectId In TblRec

                                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)

                                                If TypeOf dbObj Is AttributeDefinition Then

                                                    Dim acAtt As AttributeDefinition = dbObj
                                                    acAtt.UpgradeOpen()

                                                    Dim tagstr As String = acAtt.Tag
                                                    Dim txstr As String = acAtt.TextString
                                                    Dim promstr As String = acAtt.Prompt

                                                    If promstr = "Lot Number" Then

                                                        acAtt.TextString = TextBox38.Text

                                                    ElseIf promstr = "Block Number" Then

                                                        acAtt.TextString = TextBox42.Text

                                                    ElseIf promstr = "Section Number" Then

                                                        acAtt.TextString = TextBox43.Text

                                                    End If

                                                    acAtt.DowngradeOpen()

                                                End If
                                            Next

                                            ' Update existing block references
                                            For Each objID As ObjectId In TblRec.GetBlockReferenceIds(False, True)
                                                Dim acBlkRef As BlockReference = acTrans.GetObject(objID, OpenMode.ForWrite)
                                                acBlkRef.RecordGraphicsModified(True)

                                            Next
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                        ElseIf tagvalue = "DT1" Then

                                            If ComboBox2.Text.ToUpper() = "DESIGN 1" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 1 & 2" Then

                                                attref.TextString = "X"

                                            Else

                                                attref.TextString = ""

                                            End If


                                        ElseIf tagvalue = "DT2" Then

                                            If ComboBox2.Text.ToUpper() = "DESIGN 2" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" Then

                                                attref.TextString = "X"

                                            Else


                                                attref.TextString = ""

                                            End If


                                        ElseIf tagvalue = "DT3" Then

                                            If ComboBox2.Text.ToUpper() = "DESIGN 3" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 3 & 4" Then

                                                attref.TextString = "X"

                                            Else

                                                attref.TextString = ""


                                            End If
                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                        ElseIf tagvalue.Contains("X") Then

                                            ' Step through each object in the block table record
                                            'For Each objID As ObjectId In btr

                                            For Each objID As ObjectId In TblRec

                                                If objID.ToString = attidnum Then

                                                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)

                                                    If TypeOf dbObj Is AttributeDefinition Then

                                                        Dim acAtt As AttributeDefinition = dbObj
                                                        acAtt.UpgradeOpen()

                                                        Dim tagstr As String = acAtt.Tag
                                                        Dim txstr As String = acAtt.TextString
                                                        Dim promstr As String = acAtt.Prompt

                                                        If promstr = "TYPE 1" Then
                                                            dtype = "TYPE 1"
                                                            Exit For

                                                        ElseIf promstr = "TYPE 2" Then
                                                            dtype = "TYPE 2"
                                                            Exit For
                                                        ElseIf promstr = "TYPE 3" Then

                                                            dtype = "TYPE 3"
                                                            Exit For
                                                        End If

                                                    End If

                                                End If
                                            Next
                                            '''''''''''''''''''''''
                                            If dtype = "TYPE 1" Then

                                                If ComboBox2.Text.ToUpper() = "DESIGN 1" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 1 & 2" Then

                                                    attref.TextString = "X"

                                                Else

                                                    attref.TextString = ""

                                                End If

                                            ElseIf dtype = "TYPE 2" Then

                                                If ComboBox2.Text.ToUpper() = "DESIGN 2" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" Then

                                                    attref.TextString = "X"

                                                Else

                                                    attref.TextString = ""

                                                End If

                                            ElseIf dtype = "TYPE 3" Then


                                                If ComboBox2.Text.ToUpper() = "DESIGN 3" OrElse ComboBox2.Text.ToUpper() = "DESIGN 4" OrElse ComboBox2.Text.ToUpper() = "DESIGN 5" OrElse ComboBox2.Text.ToUpper() = "DESIGN 6" OrElse ComboBox2.Text.ToUpper() = "DESIGN 3 & 4" Then

                                                    attref.TextString = "X"

                                                Else

                                                    attref.TextString = ""

                                                End If

                                            End If

                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        End If

                                    Next

                                End If

                            Next

                        End If

                    End If

                Next

                acTrans.Commit()

                ed.WriteMessage(vbLf & "TitleBlock Updated")
                ed.Regen()

            End Using

        End Using

        Button3.BackColor = System.Drawing.SystemColors.Control
        Button3.Enabled = False

        ListBox1.Items.Clear()
        GetLayoutList()
        For Each blicky1 In Selectedlayouts
            ListBox1.SelectedItems.Add(blicky1)
        Next


    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged

        If Me.ActiveControl.Name = ComboBox1.Name Then

            If Button3.BackColor <> System.Drawing.Color.Red Then

                Button3.Enabled = False

            End If

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    For Each item As Object In ListBox1.SelectedItems

                        Dim blkname As String = "Arcxis Title Block"

                        Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.BlockName, "Arcxis Title Block*"), New TypedValue(DxfCode.LayoutName, item)}

                        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

                        Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

                        If prSelRes.Status = PromptStatus.OK Then

                            Dim SS As SelectionSet = prSelRes.Value


                            If SS IsNot Nothing Then

                                For Each brId As ObjectId In prSelRes.Value.GetObjectIds()

                                    Dim strutid As String = brId.ToString

                                    ' Open the block reference
                                    Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                                    Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                                    Dim RevvblockName As String = RevTblRec.Name

                                    For Each attId As ObjectId In RevBlockRef.AttributeCollection
                                        Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                        Dim tagvalue As String = attref.Tag

                                        If tagvalue = "ELEVATION" Then

                                            attref.TextString = ComboBox1.Text

                                        End If

                                    Next

                                Next

                            End If

                        End If

                    Next

                    acTrans.Commit()

                    acDoc.Editor.Regen()

                End Using

            End Using

        End If

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

        If Me.ActiveControl.Name = TextBox7.Name Then



            Button3.Enabled = True
            Button3.BackColor = System.Drawing.SystemColors.Control

            'Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            'Dim acCurDb As Database = acDoc.Database

            ' Lock the new document
            'Using acLckDoc As DocumentLock = acDoc.LockDocument()

            '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '        For Each item As Object In ListBox1.SelectedItems

            '            Dim blkname As String = "Arcxis Title Block"

            '            Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.BlockName, "Arcxis Title Block*"), New TypedValue(DxfCode.LayoutName, item)}

            '            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

            '            Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

            '            If prSelRes.Status = PromptStatus.OK Then

            '                Dim SS As SelectionSet = prSelRes.Value


            '                If SS IsNot Nothing Then

            '                    For Each brId As ObjectId In prSelRes.Value.GetObjectIds()

            '                        Dim strutid As String = brId.ToString

            '                        ' Open the block reference
            '                        Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

            '                        Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

            '                        Dim RevvblockName As String = RevTblRec.Name

            '                        For Each attId As ObjectId In RevBlockRef.AttributeCollection
            '                            Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
            '                            Dim tagvalue As String = attref.Tag

            '                            If tagvalue = "FR-1" Then

            '                                attref.TextString = TextBox7.Text

            '                            End If

            '                        Next

            '                    Next

            '                End If

            '            End If

            '        Next

            '        acTrans.Commit()

            '        acDoc.Editor.Regen()

            '    End Using

            'End Using

        End If

    End Sub

    'Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    '    If Me.ActiveControl.Name = TextBox1.Name Then


    '        If Button3.BackColor <> System.Drawing.Color.Red Then

    '            Button3.Enabled = False

    '        End If

    '        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '        Dim acCurDb As Database = acDoc.Database

    '        '' Lock the new document
    '        Using acLckDoc As DocumentLock = acDoc.LockDocument()

    '            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

    '                For Each item As Object In ListBox1.SelectedItems

    '                    Dim blkname As String = "Arcxis Title Block"

    '                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.BlockName, "Arcxis Title Block*"), New TypedValue(DxfCode.LayoutName, item)}

    '                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

    '                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

    '                    If prSelRes.Status = PromptStatus.OK Then

    '                        Dim SS As SelectionSet = prSelRes.Value

    '                        If SS IsNot Nothing Then

    '                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()

    '                                Dim strutid As String = brId.ToString

    '                                ' Open the block reference
    '                                Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

    '                                Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

    '                                Dim RevvblockName As String = RevTblRec.Name

    '                                For Each attId As ObjectId In RevBlockRef.AttributeCollection
    '                                    Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
    '                                    Dim tagvalue As String = attref.Tag

    '                                    If tagvalue = "PLAN " Then

    '                                        attref.TextString = TextBox1.Text

    '                                    End If

    '                                Next

    '                            Next

    '                        End If

    '                    End If

    '                Next

    '                acTrans.Commit()
    '                acDoc.Editor.Regen()

    '            End Using

    '        End Using

    '    End If

    'End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub


    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub ComboBox4_TextChanged(sender As Object, e As EventArgs) Handles ComboBox4.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub TextBox43_TextChanged(sender As Object, e As EventArgs) Handles TextBox43.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub

    Private Sub ComboBox3_TextChanged(sender As Object, e As EventArgs) Handles ComboBox3.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub


    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Button1.Enabled = True
        Button1.BackColor = System.Drawing.Color.Red

    End Sub


    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged

        Button3.Enabled = True
        Button3.BackColor = System.Drawing.Color.Red

    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim FRM As New Form_Arcxis_TB1

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim Db As Database = acDoc.Database

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(Db.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has("DPIS-CABLE") = True OrElse acLyrTbl.Has("S-FND-BTEND") = True Then

                Dim frm3 As New Form_Arcxis_TB3
                frm3.ShowDialog()
                GetLayoutList()

            Else

                Module_Arcxis_TB.ATB_CheckState = "No"
                Dim frm2 As New Form_Arcxis_TB2
                frm2.ShowDialog()
                GetLayoutList()

            End If

            ' Abort the changes to the database
            acTrans.Commit()

        End Using

        Button8.Enabled = True
        Button8.BackColor = System.Drawing.SystemColors.Control

        ListBox1.Items.Clear()
        GetLayoutList()

        Me.Update()

        If ListBox1.Items.Count > 2 Then
            TextBox1.Enabled = False
            ComboBox1.Enabled = False
        Else
            TextBox1.Enabled = True
            ComboBox1.Enabled = True
        End If

        FRM.BringToFront()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim frm As New Form_Arcxis_TB1
        Module_Arcxis_TB.ATB_UserSelect = ""

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layoutlistsort As New ArrayList
            Dim layoutlist As New ArrayList
            Dim RightLayoutName As String
            Dim NextLayoutName As String

            Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                layAndTab.Add(lay.TabOrder, lay.LayoutName)
            Next

            For Each layStr In layAndTab.Values

                If layStr <> "Model" Then

                    layoutlistsort.Add(layStr)

                End If

            Next

            layoutlistsort.Sort()

            Dim firstinlist As String = layoutlistsort(0)

            Dim firststringlen As Integer = firstinlist.Length

            Dim pagch As String = ""

            Dim filsplit As String()

            If firstinlist.Contains("FR-") Then


                If firstinlist.Contains(".") Then

                    filsplit = firstinlist.Split(".")

                    Dim fs1 As String = filsplit(1)

                    pagch = "." & fs1

                Else

                    If firststringlen = 5 Then

                        pagch = firstinlist.Substring(4)

                    ElseIf firststringlen = 6 Then

                        pagch = firstinlist.Substring(4, 1)

                        Dim i As Integer = Asc(pagch)

                        If i >= 48 AndAlso i <= 57 Then

                            pagch = firstinlist.Substring(5)

                        Else

                            pagch = firstinlist.Substring(4)

                        End If

                    End If



                End If

            ElseIf firstinlist.Contains("WB-") Then


                If firstinlist.Contains(".") Then

                    filsplit = firstinlist.Split(".")

                    Dim fs1 As String = filsplit(1)

                    If fs1.Contains(" (") Then

                        Dim filsplit2 As String() = fs1.Split(" (")

                        Dim fs2 As String = filsplit2(0)

                        pagch = "." & fs2


                    Else

                        pagch = "." & fs1

                    End If

                Else

                    filsplit = firstinlist.Split(" (")

                    Dim fs1 As String = filsplit(0)

                    If fs1.Length = 5 Then

                        pagch = fs1.Substring(4)

                    ElseIf fs1.Length = 6 Then

                        pagch = fs1.Substring(4, 1)

                        Dim i As Integer = Asc(pagch)

                        If i >= 48 AndAlso i <= 57 Then

                            pagch = fs1.Substring(5)

                        Else

                            pagch = fs1.Substring(4)

                        End If

                    End If

                End If

            ElseIf firstinlist.Contains("W-") Then

                If firstinlist.Contains(".") Then

                    filsplit = firstinlist.Split(".")

                    Dim fs1 As String = filsplit(1)

                    If fs1.Contains("(") Then

                        Dim filsplit2 As String() = fs1.Split("(")

                        Dim fs2 As String = filsplit2(0)

                        pagch = "." & fs2

                    Else
                        pagch = "." & fs1

                    End If

                Else

                    filsplit = firstinlist.Split(" (")

                    Dim fs1 As String = filsplit(0)

                    If fs1.Length = 5 Then

                        pagch = firstinlist.Substring(4)

                    ElseIf fs1.Length = 6 Then

                        pagch = firstinlist.Substring(4, 1)

                        Dim i As Integer = Asc(pagch)

                        If i >= 48 AndAlso i <= 57 Then

                            pagch = firstinlist.Substring(5)

                        Else

                            pagch = firstinlist.Substring(4)

                        End If

                    End If

                End If

            End If

            For Each layStrg As String In layoutlistsort

                If layStrg <> "Model" Then

                    Dim pagch2 As String = ""

                    Dim laynm As String

                    If layStrg.Contains(" ") Then

                        Dim layStrgsplit As String() = layStrg.Split(" ")

                        laynm = layStrgsplit(0)

                    Else

                        laynm = layStrg

                    End If


                    Dim laynmlen As Integer = laynm.Length

                    If laynmlen = 4 Then

                        pagch2 = laynm.Substring(3)


                        If pagch2 = pagch Then

                            If Not layStrg.Contains("(LC)") Then

                                layoutlist.Add(layStrg)

                            End If

                        End If


                    ElseIf laynmlen = 5 Then
                        If laynm.Contains("W-") And laynm.Contains(".") Then
                            pagch2 = laynm.Substring(3)
                        Else
                            pagch2 = laynm.Substring(4)
                        End If


                        If pagch2 = pagch Then

                            If Not layStrg.Contains("(LC)") Then

                                layoutlist.Add(layStrg)

                            End If

                        End If

                    ElseIf laynmlen = 6 Then

                        '''''''''''''''''''''''''''''''''''''''''''
                        If laynm.Contains(".") Then

                            Dim laynmsplit As String() = laynm.Split(".")

                            Dim fs1 As String = laynmsplit(1)
                            pagch2 = "." & fs1

                            If pagch2 = pagch Then

                                If Not layStrg.Contains("(LC)") Then

                                    layoutlist.Add(layStrg)


                                End If


                            End If


                        Else

                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            pagch2 = laynm.Substring(4, 1)

                            Dim i As Integer = Asc(pagch2)

                            If i >= 48 AndAlso i <= 57 Then

                                pagch2 = laynm.Substring(5)

                                If pagch2 = pagch Then


                                    If Not layStrg.Contains("(LC)") Then

                                        layoutlist.Add(layStrg)

                                    End If

                                End If


                            Else

                                pagch2 = laynm.Substring(4)

                                If pagch2 = pagch Then


                                    If Not layStrg.Contains("(LC)") Then

                                        layoutlist.Add(layStrg)

                                    End If

                                End If


                            End If

                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    ElseIf laynmlen = 7 Then

                        ''''''''''''''''''''''''''''''''''''''
                        If laynm.Contains(".") Then

                            Dim laynmsplit As String() = laynm.Split(".")

                            Dim fs1 As String = laynmsplit(1)

                            pagch2 = "." & fs1

                            If pagch2 = pagch Then

                                If Not layStrg.Contains("(LC)") Then

                                    layoutlist.Add(layStrg)


                                End If

                            End If

                        Else

                            pagch2 = laynm.Substring(4, 1)

                            Dim i As Integer = Asc(pagch2)


                            If i >= 48 AndAlso i <= 57 Then

                                pagch2 = laynm.Substring(5)


                                If pagch2 = pagch Then


                                    If Not layStrg.Contains("(LC)") Then

                                        layoutlist.Add(layStrg)

                                    End If

                                End If


                            Else

                                pagch2 = laynm.Substring(4)


                                If pagch2 = pagch Then


                                    If Not layStrg.Contains("(LC)") Then

                                        layoutlist.Add(layStrg)

                                    End If

                                End If


                            End If

                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If

                End If

            Next

            Dim elevcount As Integer = layoutlist.Count

            Dim panover As Integer

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            If acLyrTbl.Has("DPIS-CABLE") = True OrElse acLyrTbl.Has("S-FND-BTEND") = True Then

                panover = 21600

            Else

                panover = 1800 * layoutlist.Count

            End If

            ed.WriteMessage(vbLf & "panover = " & panover.ToString)

            For Each layStr In layAndTab.Values
                'ed.WriteMessage(v & vbCrLf)

                If layStr <> "Model" Then

                    RightLayoutName = layStr

                    If layStr.Contains(" (R)") Then

                        NextLayoutName = layStr.Replace(" (R)", " (L)")

                    ElseIf layStr.Contains("-R") Then

                        NextLayoutName = layStr.Replace("-R", "-L")

                    Else

                        NextLayoutName = layStr & " (L)"

                    End If

                    CreateLeftSwing(RightLayoutName, NextLayoutName)
                    ModifyViewPortCenter(NextLayoutName, panover)

                End If

            Next

            acTrans.Commit()

        End Using

        ListBox1.Items.Clear()

        GetLayoutList()

        '' moving this inside the viewport zoom to see if this can be sped up''didnt work still slow

        SetPSLTScale()

        Me.Update()
        Me.BringToFront()
        frm.Update()
        frm.BringToFront()


    End Sub

    Private Sub CreateLeftSwing(RightLayoutName As String, NextlayoutName As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Get the current document and database
        Dim acDocex As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDbex As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Create a transaction for the external drawing
            Using acTransEx As Transaction = acCurDbex.TransactionManager.StartTransaction()

                Dim layoutsEx As DBDictionary = TryCast(acTransEx.GetObject(acCurDbex.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

                ' Check to see if the layout exists in the external drawing
                If layoutsEx.Contains(NextlayoutName) = False Then

                    Dim Layid As ObjectId = layoutsEx.GetAt(RightLayoutName)
                    ' Windows.MessageBox.Show(id.ToString & " id")
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
                                'lay.Extents =


                            End Using

                        End Using

                        ' Regen the drawing to get the layout tab to display
                        acDoc.Editor.Regen()

                        ' Save the changes made
                        acTrans.Commit()
                    End Using
                Else
                    ' Display a message if the layout could not be found in the specified drawing
                    acDoc.Editor.WriteMessage(vbLf & "Layout '" & RightLayoutName & "' could not be imported '" & "'.")
                End If

                ' Discard the changes made to the external drawing file
                acTransEx.Commit()
            End Using
        End Using



    End Sub


    Private Sub ModifyViewPortCenter(NextlayoutName As String, panover As Integer)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed = acDoc.Editor
        Dim acLayoutMgr As LayoutManager = LayoutManager.Current

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
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

                                        Dim elnt As String = attref.TextString
                                        elnt = elnt.Replace("RIGHT", "LEFT")
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

                            layoutviewport.ViewCenter = New Point2d(vcp.X + panover, vcp.Y)

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
        End Using

    End Sub



    Private Function GetReferences(db As Database, bName As String) As ObjectIdCollection
        Dim result As New ObjectIdCollection()
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            If bt.Has(bName) Then
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(bName), OpenMode.ForRead), BlockTableRecord)
                For Each refId As ObjectId In btr.GetBlockReferenceIds(True, False)
                    Dim br As BlockReference = DirectCast(tr.GetObject(refId, OpenMode.ForRead), BlockReference)
                    Dim owner As BlockTableRecord = DirectCast(tr.GetObject(br.OwnerId, OpenMode.ForRead), BlockTableRecord)
                    If owner.IsLayout Then
                        result.Add(br.ObjectId)
                    End If
                Next
                For Each id As ObjectId In btr.GetAnonymousBlockIds()
                    Dim anon As BlockTableRecord = DirectCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                    For Each refId As ObjectId In anon.GetBlockReferenceIds(True, False)
                        Dim br As BlockReference = DirectCast(tr.GetObject(refId, OpenMode.ForRead), BlockReference)
                        Dim owner As BlockTableRecord = DirectCast(tr.GetObject(br.OwnerId, OpenMode.ForRead), BlockTableRecord)
                        If owner.IsLayout Then
                            result.Add(br.ObjectId)
                        End If
                    Next
                Next
            End If
            tr.Commit()
        End Using
        Return result
    End Function



    Private Sub SetPSLTScale()

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
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
                    'ed.WriteMessage(v & vbCrLf)

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
        End Using

    End Sub


    Private Sub GetTB(blkname As String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                Dim curtab As String = acLayoutMgr.CurrentLayout

                Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, curtab)}
                Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

                Dim Rev1List As New ArrayList
                'Dim Rev1List As List(Of String) = New List(Of String)

                If prSelRes.Status = PromptStatus.OK Then

                    Dim SS As SelectionSet = prSelRes.Value
                    Dim sscount As Integer = SS.Count

                    If SS IsNot Nothing Then

                        For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                            If brId.ObjectClass.Name = "AcDbBlockReference" Then

                                Dim strutid As String = brId.ToString

                                ' Open the block reference
                                Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                                Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                                Dim TBName As String = TBTblRec.Name

                                Dim layidd As String = TBTblRec.LayoutId.ToString

                                Dim BLKEntity As Entity = acTrans.GetObject(brId, OpenMode.ForRead)

                                If TBTblRec.Name = "Arcxis Title Block" Then
                                    ZoomObjects(BLKEntity)
                                End If
                            End If

                        Next

                    End If

                End If

                ' Abort the changes to the database
                acTrans.Commit()

            End Using
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

    Public Shared Sub Main()

        ' Say hi in VB.NET.
        Console.WriteLine("Startup Load")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim accurDb As Database = acDoc.Database
        Dim selectList As New ArrayList
        Dim laycount As Integer
        Dim firstlayout As String
        Dim frm As New Form_Arcxis_TB1



        Module_Arcxis_TB.ATB_UserSelect = "Delete Layouts"
        Module_Arcxis_TB.ATB_UserSelect2 = ""

        Dim layAnd As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = accurDb.TransactionManager.StartTransaction()

                Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                Dim curtab As String = acLayoutMgr.CurrentLayout

                If ListBox1.SelectedItems.Count = 0 Then

                    laycount = ListBox1.Items.Count
                    firstlayout = ListBox1.Items(0)

                    For laycount = 1 To laycount - 1

                        acLayoutMgr.DeleteLayout(ListBox1.Items(laycount))

                    Next

                    ListBox1.Items.Clear()

                Else

                    For Each selectedlayout In ListBox1.SelectedItems
                        If selectedlayout <> "FR-1A" Then
                            acLayoutMgr.DeleteLayout(selectedlayout)
                        End If
                    Next

                    ListBox1.Items.Clear()

                End If

                ' Abort the changes to the database
                acTrans.Commit()

            End Using

        End Using

        GetLayoutList()

        frm.Update()
        frm.BringToFront()

        'Button8.Enabled = False

    End Sub

    Private Sub GetLayoutList()

        Dim frm As New Form_Arcxis_TB1

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
            For Each entry As DBDictionaryEntry In layDict
                Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                layAndTab.Add(lay.TabOrder, lay.LayoutName)
            Next

            For Each layStr In layAndTab.Values
                'ed.WriteMessage(v & vbCrLf)

                If layStr <> "Model" Then

                    'EntAddLayout(layStr)

                    ListBox1.Items.Add(layStr)

                End If

            Next

            acTrans.Commit()

        End Using

        frm.Update()
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        If ComboBox8.Text IsNot "" Then
            PlotPDF11x17.Enabled = True
            PDF24x36.Enabled = True
            ComboBox8.BackColor = System.Drawing.SystemColors.Control
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor
        Dim CurrentLayerID As ObjectId = acCurDb.Clayer



        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim lytab As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim alllayers As New ArrayList
                Dim fulllayerstring As String
                Dim selectedlayer As String
                Dim CurrentLayer As LayerTableRecord = acTrans.GetObject(CurrentLayerID, OpenMode.ForRead)
                selectedlayer = "*S-SEAL-" & ComboBox8.SelectedItem

                If CurrentLayer.Name Like selectedlayer Then
                    acCurDb.Clayer = lytab("0")
                End If

                For Each layer In lytab
                    Dim lytr As LayerTableRecord = acTrans.GetObject(layer, OpenMode.ForWrite)
                    If Not lytr.Name Like selectedlayer Then
                        For Each item In ComboBox8.Items
                            fulllayerstring = "*S-SEAL-" & item
                            If lytr.Name Like fulllayerstring And lytr.Name <> selectedlayer Then

                                lytab.UpgradeOpen()
                                lytr.IsFrozen = True
                                lytr.IsOff = True
                            End If
                        Next
                    Else
                        lytab.UpgradeOpen()
                        lytr.IsFrozen = False
                        lytr.IsOff = False

                    End If


                Next
                acTrans.Commit()
                aced.Regen()
            End Using
        End Using
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim frm As New Form_Arcxis_TB1
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Dim selectedtabs As New ArrayList
        For Each Layout1 In ListBox1.SelectedItems
            selectedtabs.Add(Layout1)
        Next
        '' Lock the new document
        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                For Each item As Object In ListBox1.SelectedItems

                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, item)}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)


                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                                If brId.ObjectClass.Name = "AcDbBlockReference" Then
                                    Dim strutid As String = brId.ToString


                                    ' Open the block reference
                                    Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)
                                    Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                    Dim RevvblockName As String = RevTblRec.Name
                                    Dim revisionline = New ArrayList
                                    Dim revisionline1 = New ArrayList
                                    Dim revisionline2 = New ArrayList

                                    ' Iterate the attribute collection
                                    ' Open the attribute reference

                                    For Each attIdd As ObjectId In RevBlockRef.AttributeCollection
                                        Dim attreff As AttributeReference = DirectCast(acTrans.GetObject(attIdd, OpenMode.ForWrite), AttributeReference)
                                        Dim taggvalue As String = attreff.Tag

                                        If taggvalue = "REVDATE2" Then

                                            revisionline.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDESCRIPTION2" Then

                                            revisionline1.Add(attreff.TextString)

                                        ElseIf taggvalue = "INITIAL2" Then

                                            revisionline2.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDATE3" Then

                                            revisionline.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDESCRIPTION3" Then

                                            revisionline1.Add(attreff.TextString)

                                        ElseIf taggvalue = "INITIAL3" Then

                                            revisionline2.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDATE4" Then

                                            revisionline.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDESCRIPTION4" Then

                                            revisionline1.Add(attreff.TextString)

                                        ElseIf taggvalue = "INITIAL4" Then

                                            revisionline2.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDATE5" Then

                                            revisionline.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDESCRIPTION5" Then

                                            revisionline1.Add(attreff.TextString)

                                        ElseIf taggvalue = "INITIAL5" Then

                                            revisionline2.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDATE6" Then

                                            revisionline.Add(attreff.TextString)

                                        ElseIf taggvalue = "REVDESCRIPTION6" Then

                                            revisionline1.Add(attreff.TextString)

                                        ElseIf taggvalue = "INITIAL6" Then

                                            revisionline2.Add(attreff.TextString)

                                        End If

                                    Next

                                    For Each attIdd1 As ObjectId In RevBlockRef.AttributeCollection
                                        Dim attreff As AttributeReference = DirectCast(acTrans.GetObject(attIdd1, OpenMode.ForWrite), AttributeReference)
                                        Dim taggvalue As String = attreff.Tag

                                        If taggvalue = "REVDATE1" Then

                                            attreff.TextString = revisionline(0)

                                        ElseIf taggvalue = "REVDESCRIPTION1" Then

                                            attreff.TextString = revisionline1(0)

                                        ElseIf taggvalue = "INITIAL1" Then

                                            attreff.TextString = revisionline2(0)

                                        ElseIf taggvalue = "REVDATE2" Then

                                            attreff.TextString = revisionline(1)

                                        ElseIf taggvalue = "REVDESCRIPTION2" Then

                                            attreff.TextString = revisionline1(1)

                                        ElseIf taggvalue = "INITIAL2" Then

                                            attreff.TextString = revisionline2(1)

                                        ElseIf taggvalue = "REVDATE3" Then

                                            attreff.TextString = revisionline(2)

                                        ElseIf taggvalue = "REVDESCRIPTION3" Then

                                            attreff.TextString = revisionline1(2)

                                        ElseIf taggvalue = "INITIAL3" Then

                                            attreff.TextString = revisionline2(2)

                                        ElseIf taggvalue = "REVDATE4" Then

                                            attreff.TextString = revisionline(3)

                                        ElseIf taggvalue = "REVDESCRIPTION4" Then

                                            attreff.TextString = revisionline1(3)

                                        ElseIf taggvalue = "INITIAL4" Then

                                            attreff.TextString = revisionline2(3)

                                        ElseIf taggvalue = "REVDATE5" Then

                                            attreff.TextString = revisionline(4)

                                        ElseIf taggvalue = "REVDESCRIPTION5" Then

                                            attreff.TextString = revisionline1(4)

                                        ElseIf taggvalue = "INITIAL5" Then

                                            attreff.TextString = revisionline2(4)

                                        ElseIf taggvalue = "REVDATE6" Then

                                            attreff.TextString = "-"

                                        ElseIf taggvalue = "REVDESCRIPTION6" Then

                                            attreff.TextString = "-"

                                        ElseIf taggvalue = "INITIAL6" Then
                                            attreff.TextString = "-"

                                        End If
                                    Next

                                End If

                            Next

                        End If

                    End If

                Next

                acTrans.Commit()
            End Using

        End Using

        Button1.BackColor = System.Drawing.SystemColors.Control
        Button1.Enabled = False
        ListBox1.Items.Clear()
        GetLayoutList()

        For Each tablay1 In selectedtabs
            ListBox1.SelectedItems.Add(tablay1)
        Next

        Me.Update()
        Me.Refresh()
        frm.Update()
        frm.Refresh()
        ed.Regen()

        ComboBox5.Enabled = True
        DateTimePicker1.Enabled = True
        TextBox9.Enabled = True

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

        If RadioButton7.Checked <> True Then

            Dim psize As String = "11x17"

            WaterMark(psize)

            PageSetUp11x17()

        End If

        If RadioButton11.Checked = True Or RadioButton12.Checked = True Then
            PageSetUp11x17()
        End If

        'Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            'Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim FRM As New Form_Arcxis_TB1
                Dim CurVar As Object
                'Get the BACKGROUNDPLOT SYSTEM VARIABLE
                CurVar = Application.GetSystemVariable("BackGroundPlot")
                'SET BACKGROUNDPLOT SYSTEM VARIABLE TO ZERO
                Application.SetSystemVariable("BackGroundPlot", 0)
                Application.SetSystemVariable("FILEDIA", 0)
                'acDoc.SendStringToExecute("-updatefields all 0 ", True, False, False)
                Try
                    If Not Directory.Exists("c:\temp\") Then
                        Directory.CreateDirectory("c:\temp\")
                    End If

                    Dim layoutManager__2 As LayoutManager = LayoutManager.Current
                    layoutManager__2.CurrentLayout = ATB_FirstLayoutName

                    Dim DWGnm As String = Application.GetSystemVariable("dwgName")

                    'Dim outputDir As String = "c:\temp\"
                    Dim outputDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\"
                    Dim DWGname As String = DirectCast(Application.GetSystemVariable("DWGNAME"), String)
                    DWGname = DWGname.Remove(DWGname.Length - 4)
                    Dim pdfFile As String = outputDir & DWGname & ".pdf"

                    'Export to PDF
                    acDoc.SendStringToExecute("-EXPORT P A " & pdfFile & vbCr, True, False, False)

                Finally

                    'REVERT BACK TO THE ORIGINAL BACKGROUNDPLOT SYSTEM VARIABLE
                    Application.SetSystemVariable("BackGroundPlot", CurVar)

                    Application.SetSystemVariable("FILEDIA", 0)

                End Try

                ' Save the changes made

                'FRM.Close()
                acTrans.Commit()

            End Using

        End Using

    End Sub


End Class