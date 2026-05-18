Imports System
Imports System.Windows.Forms
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Bricscad.PlottingServices
Imports Teigha.Colors
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Color = Teigha.Colors.Color

<Assembly: CommandClass(GetType(PageBlockDivisionCommands))>

Public NotInheritable Class PageBlockDivisionHelper
    Private Sub New()
    End Sub

    Public Shared Sub EnsurePageDivisionsAttribute(tr As Transaction, db As Database)
        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        If Not bt.Has("PAGE") Then Return

        Dim pageBtr As BlockTableRecord = TryCast(tr.GetObject(bt("PAGE"), OpenMode.ForWrite), BlockTableRecord)
        If pageBtr Is Nothing Then Return

        Dim divisionsDef As AttributeDefinition = Nothing

        For Each entId As ObjectId In pageBtr
            Dim ad As AttributeDefinition = TryCast(tr.GetObject(entId, OpenMode.ForRead), AttributeDefinition)
            If ad IsNot Nothing AndAlso String.Equals(ad.Tag, "DIVISIONS", StringComparison.OrdinalIgnoreCase) Then
                divisionsDef = ad
                divisionsDef.UpgradeOpen()
                divisionsDef.Position = New Point3d(11.253, -1068, 0)
                divisionsDef.Height = 48.0092
                divisionsDef.TextStyleId = db.Textstyle
                divisionsDef.Justify = AttachmentPoint.TopLeft
                divisionsDef.Rotation = 0
                divisionsDef.WidthFactor = 0.8
                divisionsDef.Constant = False
                divisionsDef.Verifiable = False
                divisionsDef.Invisible = False
                divisionsDef.LockPositionInBlock = True
                divisionsDef.Layer = "0"
                divisionsDef.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                divisionsDef.HorizontalMode = TextHorizontalMode.TextLeft
                divisionsDef.VerticalMode = TextVerticalMode.TextBase
            End If
        Next

        If divisionsDef Is Nothing Then
            divisionsDef = New AttributeDefinition()
            divisionsDef.Position = New Point3d(11.253, -1068, 0)
            divisionsDef.Height = 48.0092
            divisionsDef.TextStyleId = db.Textstyle
            divisionsDef.Justify = AttachmentPoint.TopLeft
            divisionsDef.Tag = "DIVISIONS"
            divisionsDef.Prompt = "DIVISIONS"
            divisionsDef.TextString = ""
            divisionsDef.Rotation = 0
            divisionsDef.WidthFactor = 0.8
            divisionsDef.Constant = False
            divisionsDef.Verifiable = False
            divisionsDef.Invisible = False
            divisionsDef.LockPositionInBlock = True
            divisionsDef.Layer = "0"
            divisionsDef.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
            divisionsDef.HorizontalMode = TextHorizontalMode.TextLeft
            divisionsDef.VerticalMode = TextVerticalMode.TextBase

            pageBtr.AppendEntity(divisionsDef)
            tr.AddNewlyCreatedDBObject(divisionsDef, True)
        End If

        Dim refIds As ObjectIdCollection = pageBtr.GetBlockReferenceIds(True, True)
        For Each refId As ObjectId In refIds
            Dim br As BlockReference = TryCast(tr.GetObject(refId, OpenMode.ForWrite), BlockReference)
            If br Is Nothing Then Continue For

            Dim hasDivisionsRef As Boolean = False
            For Each attId As ObjectId In br.AttributeCollection
                Dim ar As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                If ar IsNot Nothing AndAlso String.Equals(ar.Tag, "DIVISIONS", StringComparison.OrdinalIgnoreCase) Then
                    hasDivisionsRef = True
                    Exit For
                End If
            Next

            If Not hasDivisionsRef Then
                Dim ar As New AttributeReference()
                ar.SetAttributeFromBlock(divisionsDef, br.BlockTransform)
                ar.TextString = ""
                br.AttributeCollection.AppendAttribute(ar)
                tr.AddNewlyCreatedDBObject(ar, True)
            End If
        Next
    End Sub
End Class

Public Class PageBlockDivisionCommands
    <CommandMethod("EPD", CommandFlags.Modal)>
    Public Shared Sub EnsurePageDivisions()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return

        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                PageBlockDivisionHelper.EnsurePageDivisionsAttribute(tr, db)
                tr.Commit()
            End Using
        End Using

        ed.WriteMessage(vbLf & "PAGE block DIVISIONS attribute check/sync complete.")
    End Sub

    Public Shared Function GetPageBlockDivisions() As List(Of String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim Divisions As New List(Of String)

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
                        Dim blockObjectId As ObjectId = blkRef.ObjectId
                        ' To store the attribute values from this block instance
                        Dim valuesList As New List(Of String)

                        ' Check if the block reference has attributes
                        Dim tagValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                        ' --- Collect pass ---
                        If blkRef.AttributeCollection.Count > 0 Then
                            For Each attId As ObjectId In blkRef.AttributeCollection
                                Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForRead), AttributeReference)
                                If attRef Is Nothing Then Continue For

                                Dim tag As String = If(attRef.Tag, "").Trim()
                                Dim txt As String = If(attRef.TextString, "").Trim()
                                If tag = "" OrElse txt = "" Then Continue For
                                ' Your extra per-tag lists:
                                Select Case tag.ToUpperInvariant()
                                    Case "DIVISIONS"
                                        For Each v In txt.Split(","c).Select(Function(s) s.Trim()).Where(Function(s) s <> "")
                                            If Not Divisions.Contains(v) Then Divisions.Add(v)
                                        Next
                                End Select
                            Next
                        End If

                    End If
                    'End If
                End If
            Next

            ' Commit the transaction
            acTrans.Commit()

        End Using

        Return Divisions

    End Function

    Public Shared Sub SyncCheckedListBoxExact(clb As CheckedListBox, values As IEnumerable(Of String))
        If clb Is Nothing Then Exit Sub
        If values Is Nothing Then values = Enumerable.Empty(Of String)()

        Dim target As List(Of String) =
        values.
            Select(Function(s) If(s, "").Trim()).
            Where(Function(s) s <> "").
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        For i As Integer = clb.Items.Count - 1 To 0 Step -1
            Dim itemText As String = clb.Items(i).ToString().Trim()
            If Not target.Contains(itemText, StringComparer.OrdinalIgnoreCase) Then
                clb.Items.RemoveAt(i)
            End If
        Next

        For Each t As String In target
            Dim idx As Integer = -1
            For i As Integer = 0 To clb.Items.Count - 1
                If String.Equals(clb.Items(i).ToString().Trim(), t, StringComparison.OrdinalIgnoreCase) Then
                    idx = i
                    Exit For
                End If
            Next
            If idx < 0 Then idx = clb.Items.Add(t)
            clb.SetItemChecked(idx, True)
        Next
    End Sub
End Class
