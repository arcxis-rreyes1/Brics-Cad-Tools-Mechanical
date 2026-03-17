Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms
Imports Arcxis_Cad_Tools_Brics
Imports Arcxis_Cad_Tools_Brics.Arcxis_Cad_Tools
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports Color = System.Drawing.Color
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Application = Bricscad.ApplicationServices.Application

Public Class Form_FramingProperties
    Private Sub Form_DrawingProperties_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Try
            Dim fm As New FileManipulation()
            fm.SyncCustomPropsFromAutoPlanPrintInfo()
            acEd.WriteMessage(vbLf & "Custom DWG properties synced from 'AutoPlanPrintInfo'.")
        Catch ex As System.Exception
            acEd.WriteMessage(vbLf & "Sync failed: " & ex.Message)
        End Try

        Dim FormBuilder = GetCustomDwgPropReliable("BUILDER")
        Dim Formplan = GetCustomDwgPropReliable("PLAN")
        Dim FormStamps = GetCustomDwgPropReliable("STAMPS")
        Dim FormProject = GetCustomDwgPropReliable("PROJECT NUMBER")
        Dim FormSheetLabels = GetCustomDwgPropReliable("SHEET LABELING")
        Dim FormBuilderSpec = GetCustomDwgPropReliable("PackageFRWB")

        TextBox1.Text = FormBuilder
        TextBox2.Text = Formplan
        TextBox3.Text = FormProject

        Dim StampLayers As New List(Of String)
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim blockName As String = "Arcxis Title Block"
            Dim sourceDwgPath As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\Arcxis\Engineering\Drafting Standards\CAD Blocks\Arcxis Title Block - Block.dwg"
            FileManipulation.BlockImport(acCurDb, blockName, sourceDwgPath, acTrans)

            Dim lytab As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite)
            Dim cleanlayerstring As String

            For Each layer In lytab

                Dim lytr As LayerTableRecord = acTrans.GetObject(layer, OpenMode.ForWrite)

                If lytr.Name Like "Master Seal File|S-SEAL-*" Then

                    cleanlayerstring = lytr.Name.Substring(24, lytr.Name.Length - 24)

                    StampLayers.Add(cleanlayerstring)

                    lytr.IsOff = True
                    lytr.IsFrozen = True

                End If

            Next


            acTrans.Commit()

        End Using

        Dim frm As New Framing_Vault_Transfer

        Dim normalItems = StampLayers.Where(Function(x) Not x.ToLower().Contains("review")).OrderBy(Function(x) x).ToList()
        Dim reviewItems = StampLayers.Where(Function(x) x.ToLower().Contains("review")).OrderBy(Function(x) x).ToList()


        Dim finalList As New List(Of String)

        'tml-tx is added by default as its most common stamp set to item 0 and checked/selected
        If normalItems.Contains("TML-TX") Then
            'finalList.Add("TML-TX")
            normalItems.Remove("TML-TX")
        End If

        ' Add the rest of the normal items
        finalList.AddRange(normalItems)

        If reviewItems.Contains("For Review") Then
            finalList.Add("For Review")
            reviewItems.Remove("For Review")
        End If
        ' Add the review items at the bottom
        finalList.AddRange(reviewItems)

        For Each seal In finalList
            SealsList.Items.Add(seal)
        Next

        If Not String.IsNullOrWhiteSpace(FormStamps) Then
            Dim tokens = FormStamps.Split(","c).
                        Select(Function(s) s.Trim()).
                        Where(Function(s) s <> "").
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        ToList()

            For Each t In tokens
                Dim idx As Integer = -1
                For i = 0 To frm.SealsList.Items.Count - 1
                    If String.Equals(frm.SealsList.Items(i).ToString().Trim(),
                             t,
                             StringComparison.OrdinalIgnoreCase) Then
                        idx = i : Exit For
                    End If
                Next
                If idx >= 0 Then frm.SealsList.SetItemChecked(idx, True)
            Next
        End If

        If FormBuilderSpec = "WSFW" Then
            WSFW.Checked = True
        ElseIf FormBuilderSpec = "RFR" Then
            RFR.Checked = True
        End If

        If Not String.IsNullOrWhiteSpace(FormStamps) Then
            Dim tokens = FormStamps.Split(","c).
                    Select(Function(s) s.Trim()).
                    Where(Function(s) s <> "").
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

            For Each t In tokens
                Dim idx As Integer = -1
                For i = 0 To SealsList.Items.Count - 1
                    If String.Equals(SealsList.Items(i).ToString().Trim(),
                         t,
                         StringComparison.OrdinalIgnoreCase) Then
                        idx = i : Exit For
                    End If
                Next
                If idx >= 0 Then SealsList.SetItemChecked(idx, True)
            Next
        End If

        Dim found As Boolean = False
        For Each item In SheetLabels.Items
            If String.Equals(item.ToString(), FormSheetLabels, StringComparison.OrdinalIgnoreCase) Then
                SheetLabels.SelectedItem = item
                found = True
                Exit For
            End If
        Next
        If Not found Then
            SheetLabels.BackColor = Color.Red
            Button1.Visible = False
        End If


    End Sub

    Private Sub SheetLabels_SelectedValueChanged(sender As Object, e As EventArgs) Handles SheetLabels.SelectedValueChanged

        If Button1.Visible = False Then
            Button1.Visible = True
            SheetLabels.BackColor = Color.White
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim FormBuilder = GetCustomDwgPropReliable("BUILDER")
        Dim Formplan = GetCustomDwgPropReliable("PLAN")
        Dim FormStamps = GetCustomDwgPropReliable("STAMPS")
        Dim FormProject = GetCustomDwgPropReliable("PROJECT NUMBER")
        Dim FormSheetLabels = GetCustomDwgPropReliable("SHEET LABELING")
        Dim FormBuilderSpec = GetCustomDwgPropReliable("PackageFRWB")
        Dim sealloop As New List(Of String)

        If TextBox1.Text <> FormBuilder Then
            SetCustomDwgPropReliable("BUILDER", TextBox1.Text)
        End If
        If TextBox2.Text <> Formplan Then
            SetCustomDwgPropReliable("PLAN", TextBox2.Text)
        End If
        If TextBox3.Text <> FormProject Then
            SetCustomDwgPropReliable("PROJECT NUMBER", TextBox3.Text)
        End If
        If RFR.Checked = True Then
            SetCustomDwgPropReliable("PackageFRWB", "RFR")
        ElseIf WSFW.Checked = True Then
            SetCustomDwgPropReliable("PackageFRWB", "WSFW")
        End If
        Dim selectedStamps = String.Join(", ", SealsList.CheckedItems.Cast(Of String)())
        If selectedStamps <> FormStamps Then
            SetCustomDwgPropReliable("STAMPS", selectedStamps)
        End If

        For Each item In SealsList.CheckedItems
            sealloop.Add(item)
        Next

        If SheetLabels.SelectedItem IsNot Nothing AndAlso SheetLabels.SelectedItem.ToString() <> FormSheetLabels Then
            SetCustomDwgPropReliable("SHEET LABELING", SheetLabels.SelectedItem.ToString())
        End If

        Dim acDoc = Application.DocumentManager.MdiActiveDocument
        Dim acDb = acDoc.Database
        Dim accurdb = acDoc.Database
        Dim ed = acDoc.Editor
        Dim acEd As Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()


            Using acTrans As Transaction = accurdb.TransactionManager.StartTransaction()

                Dim bt As BlockTable = acTrans.GetObject(accurdb.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                For Each objId As ObjectId In ms

                    Dim ent As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                    If TypeOf ent Is BlockReference Then

                        Dim blkRef As BlockReference = CType(ent, BlockReference)
                        Dim btr As BlockTableRecord = acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead)
                        Dim blockName As String = btr.Name

                        ' === CASE 1: Named "TB-INFO" ===
                        If blockName.Equals("AutoPlanPrintInfo", StringComparison.OrdinalIgnoreCase) Then

                            For Each attId As ObjectId In blkRef.AttributeCollection

                                ' Open the attribute reference
                                Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                Dim tagvalue As String = attref.Tag
                                Dim textvalue As String = attref.TextString

                                If tagvalue.Contains("BUILDER") Then

                                    attref.TextString = TextBox1.Text


                                ElseIf tagvalue.Contains("PLAN") Then

                                    attref.TextString = TextBox2.Text

                                ElseIf tagvalue.Contains("STAMPS") Then

                                    attref.TextString = selectedStamps

                                End If
                            Next

                        End If

                    End If

                Next
            End Using
        End Using

        ' Inside your form's method
        'Framing_Vault_Transfer.EnsurePlanInfoAtOrigin(New Point3d(0, 0, 0), TextBox1.Text, TextBox2.Text, sealloop)

        MessageBox.Show("Drawing properties updated.", "Arcxis Cad Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Public Shared Function GetCustomDwgPropReliable(propName As String) As String
        Dim db = Application.DocumentManager.MdiActiveDocument.Database

        ' Build from current summary info (brings along custom props)
        Dim b As New DatabaseSummaryInfoBuilder(db.SummaryInfo)

        ' Prefer the builder's editable table
        Dim tbl = TryCast(b.CustomPropertyTable, IDictionary)
        If tbl IsNot Nothing Then
            For Each de As DictionaryEntry In tbl
                If String.Equals(CStr(de.Key), propName, StringComparison.OrdinalIgnoreCase) Then
                    Return If(de.Value, Nothing)?.ToString()
                End If
            Next
            Return Nothing
        End If

        ' Fallback: enumerate si.CustomProperties if available (handles odd versions)
        Dim si = db.SummaryInfo
        Dim propsObj As Object = si.CustomProperties
        If propsObj IsNot Nothing Then
            For Each kv As Object In DirectCast(propsObj, IEnumerable)
                Dim t = kv.GetType()
                Dim k As String = CStr(t.GetProperty("Key").GetValue(kv, Nothing))
                If String.Equals(k, propName, StringComparison.OrdinalIgnoreCase) Then
                    Dim v = t.GetProperty("Value").GetValue(kv, Nothing)
                    Return If(v, Nothing)?.ToString()
                End If
            Next
        End If

        Return Nothing
    End Function

    Public Shared Sub SetCustomDwgPropReliable(propName As String, propValue As String)
        Dim db = Application.DocumentManager.MdiActiveDocument.Database
        Dim b As New DatabaseSummaryInfoBuilder(db.SummaryInfo)
        Dim tbl = DirectCast(b.CustomPropertyTable, IDictionary)
        If tbl.Contains(propName) Then
            tbl(propName) = propValue
        Else
            tbl.Add(propName, propValue)
        End If
        db.SummaryInfo = b.ToDatabaseSummaryInfo()
    End Sub

    Private Sub RFR_Click(sender As Object, e As EventArgs) Handles RFR.Click
        If RFR.Checked = True Then
            RFR.Checked = False
            WSFW.Checked = False
            WSFW.Enabled = True
        Else
            RFR.Checked = True
            WSFW.Checked = False
            WSFW.Enabled = False
        End If
    End Sub

    Private Sub WSFW_Click(sender As Object, e As EventArgs) Handles WSFW.Click
        If WSFW.Checked = True Then
            WSFW.Checked = False
            RFR.Checked = False
            RFR.Enabled = True
        Else
            WSFW.Checked = True
            RFR.Checked = False
            RFR.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


End Class