' Option Strict On recommended in your project settings.
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Bricscad.ApplicationServices

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.MaterialListCommands))>

Namespace Arcxis_Cad_Tools

    Public Class MaterialListCommands

        Private Enum VendorMode
            GenericJoist = 1
            Boise = 2
            TrusJoist = 3
        End Enum

        Private Class Item
            Public SortNum As Integer
            Public Product As String = ""
            Public LengthFt As Integer
            Public Qty As Integer
            Public Layer As String = ""
        End Class

        Private Class RolledUp
            Public SortNum As Integer
            Public Product As String = ""
            Public LengthFt As Integer
            Public QtyText As String = "" ' "1", "2", ... or "LF"
        End Class

        <CommandMethod("MML3")>
        Public Sub MMl4()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            Dim ed = doc.Editor

            ' 1) Vendor mode (replaces UserForm1 option buttons)
            Dim mode = PromptVendorMode(ed)
            If mode Is Nothing Then Return

            ' 2) Select objects on screen
            Dim ssRes = ed.GetSelection(New SelectionFilter(New TypedValue() {
                New TypedValue(CInt(DxfCode.Start), "LINE,INSERT,MTEXT,TEXT")
            }))
            If ssRes.Status <> PromptStatus.OK OrElse ssRes.Value Is Nothing OrElse ssRes.Value.Count = 0 Then
                ed.WriteMessage(vbLf & "Nothing selected.")
                Return
            End If

            Dim items As New List(Of Item)()

            Using tr = db.TransactionManager.StartTransaction()

                ' --- Pass A: GROUP / RM-RI / RM-BL labeling & collection
                For Each id In ssRes.Value.GetObjectIds()

                    Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If ent Is Nothing Then Continue For

                    Dim ln = TryCast(ent, Line)
                    If ln Is Nothing Then Continue For

                    Dim layerName As String = ln.Layer
                    Dim layerU As String = layerName.ToUpperInvariant()

                    If layerU.StartsWith("S-FRM-GROUP", StringComparison.OrdinalIgnoreCase) OrElse
                       layerU.StartsWith("S-FRM-FL-GROUP-", StringComparison.OrdinalIgnoreCase) Then

                        ' Extract group number from old/new group layer names
                        Dim grpNumStr As String
                        If layerU.StartsWith("S-FRM-FL-GROUP-", StringComparison.OrdinalIgnoreCase) Then
                            grpNumStr = layerName.Substring("S-FRM-FL-GROUP-".Length)
                        Else
                            grpNumStr = layerName.Substring("S-FRM-GROUP".Length)
                        End If

                        ' Keep all digits in suffix (handles "1" and "-1")
                        Dim digits As String = ""
                        For Each ch As Char In grpNumStr
                            If Char.IsDigit(ch) Then
                                digits &= ch
                            End If
                        Next

                        Dim groupLabel As String = If(digits.Length > 0, "GROUP" & digits, "GROUP")

                        Dim thicknessIn As Double = ln.Thickness
                        Dim lenFt As Integer = CInt(Math.Round((ln.Length + 6.0) / 12.0, 0, MidpointRounding.AwayFromZero))

                        Dim mappedLabel As String = MapGroupLabel(groupLabel, mode.Value)
                        Dim product As String = thicknessIn.ToString() & """" & " " & mappedLabel

                        items.Add(New Item With {
                            .SortNum = 1,
                            .Product = product,
                            .LengthFt = lenFt,
                            .Qty = 1,
                            .Layer = layerName
                        })

                        ' Place rotated label text like VBA did
                        PlaceRotatedLabel(tr, db, mappedLabel & "-" & lenFt & "'", ln)

                    ElseIf layerU.EndsWith("RM-RI", StringComparison.OrdinalIgnoreCase) OrElse
                           layerU.EndsWith("RIM", StringComparison.OrdinalIgnoreCase) OrElse
                           layerU.EndsWith("-RIM", StringComparison.OrdinalIgnoreCase) Then

                        Dim lenFt As Integer = CInt(Math.Round((ln.Length + 6.0) / 12.0, 0, MidpointRounding.AwayFromZero))

                        items.Add(New Item With {
                            .SortNum = 3,
                            .Product = ln.Thickness.ToString() & """" & " RIM BOARD",
                            .LengthFt = lenFt,
                            .Qty = 1,
                            .Layer = layerName
                        })

                    ElseIf layerU.EndsWith("RM-BL", StringComparison.OrdinalIgnoreCase) OrElse
                           layerU.EndsWith("BLK", StringComparison.OrdinalIgnoreCase) OrElse
                           layerU.EndsWith("-BLK", StringComparison.OrdinalIgnoreCase) Then

                        Dim lenFt As Integer = CInt(Math.Round((ln.Length + 6.0) / 12.0, 0, MidpointRounding.AwayFromZero))

                        items.Add(New Item With {
                            .SortNum = 2,
                            .Product = ln.Thickness.ToString() & """" & " BLOCKING PANELS",
                            .LengthFt = lenFt,
                            .Qty = 1,
                            .Layer = layerName
                        })

                    End If
                Next

                ' --- Pass B: S-FRM-* PSL/LVL/LSL logic
                For Each id In ssRes.Value.GetObjectIds()

                    Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    Dim ln = TryCast(ent, Line)
                    If ln Is Nothing Then Continue For

                    Dim layerName = ln.Layer
                    Dim lenFt As Integer = CInt(Math.Round((ln.Length + 9.9) / 12.0, 0, MidpointRounding.AwayFromZero))
                    Dim thick As Double = ln.Thickness

                    Select Case layerName.ToUpperInvariant()
                        Case "S-FRM-3.5", "S-FRM-BM-3.5"
                            items.Add(New Item With {.SortNum = 6, .Product = "3.5x" & thick.ToString() & """" & " PSL 2.0E", .LengthFt = lenFt, .Qty = 1, .Layer = layerName})
                        Case "S-FRM-5.25", "S-FRM-BM-5.25"
                            items.Add(New Item With {.SortNum = 6, .Product = "5.25x" & thick.ToString() & """" & " PSL 2.0E", .LengthFt = lenFt, .Qty = 1, .Layer = layerName})
                        Case "S-FRM-7", "S-FRM-BM-7"
                            items.Add(New Item With {.SortNum = 6, .Product = "7x" & thick.ToString() & """" & " PSL 2.0E", .LengthFt = lenFt, .Qty = 1, .Layer = layerName})
                        Case "S-FRM-1.75", "S-FRM-BM-1.75"
                            Dim prod As String
                            If ln.Color.ColorIndex = 213 Then
                                prod = "1.75x" & thick.ToString() & """" & " LSL 1.55E"
                            Else
                                prod = "1.75x" & thick.ToString() & """" & " LVL 2.0E"
                            End If
                            items.Add(New Item With {.SortNum = 6, .Product = prod, .LengthFt = lenFt, .Qty = 1, .Layer = layerName})
                    End Select
                Next

                ' --- Pass C: Hangers (block inserts with attribute TAG "TYPE")
                For Each id In ssRes.Value.GetObjectIds()

                    Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    Dim br = TryCast(ent, BlockReference)
                    If br Is Nothing Then Continue For
                    If br.AttributeCollection Is Nothing OrElse br.AttributeCollection.Count = 0 Then Continue For

                    Dim hangerType As String = Nothing
                    For Each attId As ObjectId In br.AttributeCollection
                        Dim ar = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                        If ar Is Nothing Then Continue For
                        If String.Equals(ar.Tag, "TYPE", StringComparison.OrdinalIgnoreCase) Then
                            hangerType = ar.TextString
                            Exit For
                        End If
                    Next

                    If Not String.IsNullOrWhiteSpace(hangerType) Then
                        items.Add(New Item With {
                            .SortNum = 8,
                            .Product = "SIMPSON HANGER " & hangerType,
                            .LengthFt = 0,
                            .Qty = 1,
                            .Layer = br.Layer
                        })
                    End If
                Next

                ' --- Clean selected beam texts (your VBA replaced strings based on vendor)
                For Each id In ssRes.Value.GetObjectIds()

                    Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If ent Is Nothing Then Continue For

                    Dim txt = TryCast(ent, DBText)
                    If txt IsNot Nothing Then
                        txt.UpgradeOpen()
                        txt.TextString = ConvertVendorText(txt.TextString, mode.Value)
                        Continue For
                    End If

                    Dim mt = TryCast(ent, MText)
                    If mt IsNot Nothing Then
                        mt.UpgradeOpen()
                        mt.Contents = ConvertVendorText(mt.Contents, mode.Value)
                    End If
                Next

                tr.Commit()
            End Using

            If items.Count = 0 Then
                ed.WriteMessage(vbLf & "No material items found in selection.")
                Return
            End If

            ' 3) Sort like your VBA: sortnum asc, product asc, length desc
            items.Sort(Function(a, b)
                           Dim c = a.SortNum.CompareTo(b.SortNum)
                           If c <> 0 Then Return c
                           c = String.Compare(a.Product, b.Product, StringComparison.OrdinalIgnoreCase)
                           If c <> 0 Then Return c
                           Return b.LengthFt.CompareTo(a.LengthFt) ' longer first
                       End Function)

            ' 4) Roll up (LF behavior for sort types 2 & 3)
            Dim rolled As List(Of RolledUp) = RollUp(items)

            ' 5) OPTION 1: Create a UNIQUE block definition name per run
            Dim mmBlockName As String = GetNextMmListName(db)

            ' (optional) pull a base template dwg, then overwrite its contents with our list
            Dim fileToInsert As String = "\\egnytedrive\energyinspectors\Shared\Arcxis\Engineering\Drafting Standards\CAD Blocks\PTS_FRM\MMList.dwg"
            BuildMmListBlockDefinition(db, fileToInsert, rolled, mmBlockName)

            ' 6) Insert the new unique block
            Dim pPt = ed.GetPoint(vbLf & "Enter a point for MATERIAL LIST: ")
            If pPt.Status <> PromptStatus.OK Then Return

            Using tr = db.TransactionManager.StartTransaction()

                Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If Not bt.Has(mmBlockName) Then
                    ed.WriteMessage(vbLf & "Block definition '" & mmBlockName & "' not found.")
                    Return
                End If

                Dim btrModel = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                Dim br As New BlockReference(pPt.Value, bt(mmBlockName))
                btrModel.AppendEntity(br)
                tr.AddNewlyCreatedDBObject(br, True)

                tr.Commit()
            End Using

            ed.WriteMessage(vbLf & "Created and inserted " & mmBlockName & ".")
        End Sub

        ' ------------------------- OPTION 1: unique name helpers -------------------------

        Private Function GetNextMmListName(db As Database) As String
            ' Finds the highest MMList_### in the drawing and returns the next one.
            Using tr = db.TransactionManager.StartTransaction()

                Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                Dim maxN As Integer = 0

                For Each id As ObjectId In bt
                    Dim btr = CType(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                    Dim nm As String = btr.Name

                    If nm.StartsWith("MMList_", StringComparison.OrdinalIgnoreCase) Then
                        Dim suffix As String = nm.Substring("MMList_".Length) ' "001"
                        Dim n As Integer
                        If Integer.TryParse(suffix, n) Then
                            If n > maxN Then maxN = n
                        End If
                    End If
                Next

                tr.Commit()

                Return "MMList_" & (maxN + 1).ToString("000")
            End Using
        End Function

        ' ------------------------- Existing helpers -------------------------

        Private Function PromptVendorMode(ed As Editor) As VendorMode?
            Dim pko As New PromptKeywordOptions(vbLf & "Vendor mode [Generic/Boise/TrusJoist] <Generic>: ", "Generic Boise TrusJoist")
            pko.AllowNone = True
            Dim res = ed.GetKeywords(pko)
            If res.Status = PromptStatus.Cancel Then Return Nothing
            If res.Status = PromptStatus.None OrElse String.IsNullOrWhiteSpace(res.StringResult) Then Return VendorMode.GenericJoist

            Select Case res.StringResult.ToUpperInvariant()
                Case "GENERIC" : Return VendorMode.GenericJoist
                Case "BOISE" : Return VendorMode.Boise
                Case "TRUSJOIST" : Return VendorMode.TrusJoist
            End Select

            Return VendorMode.GenericJoist
        End Function

        Private Function MapGroupLabel(groupLabel As String, mode As VendorMode) As String
            Dim g = groupLabel.ToUpperInvariant()
            Select Case g
                Case "GROUP1"
                    If mode = VendorMode.TrusJoist Then Return "TJI 110"
                    If mode = VendorMode.Boise Then Return "BCI 4500"
                Case "GROUP2"
                    If mode = VendorMode.TrusJoist Then Return "TJI 210"
                    If mode = VendorMode.Boise Then Return "BCI 5000"
                Case "GROUP3"
                    If mode = VendorMode.TrusJoist Then Return "TJI 230"
                    If mode = VendorMode.Boise Then Return "BCI 6000"
                Case "GROUP4"
                    If mode = VendorMode.TrusJoist Then Return "TJI 360"
                    If mode = VendorMode.Boise Then Return "BCI 60"
            End Select
            Return groupLabel
        End Function

        Private Sub PlaceRotatedLabel(tr As Transaction, db As Database, label As String, ln As Line)

            Dim ang As Double = ln.Angle
            If ang > Math.PI / 2.0 Then ang = Math.PI + ang

            Dim mid As New Point3d((ln.StartPoint.X + ln.EndPoint.X) / 2.0,
                                   (ln.StartPoint.Y + ln.EndPoint.Y) / 2.0,
                                   0.0)

            Dim alignPt As New Point3d(mid.X + 5.0, mid.Y + 5.0, mid.Z)

            Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            Dim t As New DBText() With {
                .TextString = label,
                .Height = 6.0,
                .Position = mid,
                .HorizontalMode = TextHorizontalMode.TextCenter,
                .VerticalMode = TextVerticalMode.TextVerticalMid,
                .AlignmentPoint = alignPt,
                .Rotation = ang
            }

            ms.AppendEntity(t)
            tr.AddNewlyCreatedDBObject(t, True)
        End Sub

        Private Function ConvertVendorText(s As String, mode As VendorMode) As String
            If String.IsNullOrEmpty(s) Then Return s
            Dim t = s

            If mode = VendorMode.TrusJoist Then
                t = t.Replace("VERSA LAM 2.0E", "MICROLLAM 2.0E").
                      Replace("VERSA LAM 1.7E", "MICROLLAM 2.0E").
                      Replace("LVL 2.0E", "MICROLLAM 2.0E").
                      Replace("BCI 4500", "TJI 110").
                      Replace("BCI 5000", "TJI 210").
                      Replace("BCI 6000", "TJI 230").
                      Replace("BCI 60", "TJI 360").
                      Replace("GROUP1", "TJI 110").
                      Replace("GROUP2", "TJI 210").
                      Replace("GROUP3", "TJI 230").
                      Replace("GROUP4", "TJI 360")

            ElseIf mode = VendorMode.Boise Then
                t = t.Replace("MICROLLAM 2.0E", "VERSA LAM 2.0E").
                      Replace("TIMBERSTRAND 1.55E", "VERSA LAM 1.7E").
                      Replace("LVL 2.0E", "VERSA LAM 2.0E").
                      Replace("TJI 110", "BCI 4500").
                      Replace("TJI 210", "BCI 5000").
                      Replace("TJI 230", "BCI 6000").
                      Replace("TJI 360", "BCI 60").
                      Replace("GROUP1", "BCI 4500").
                      Replace("GROUP2", "BCI 5000").
                      Replace("GROUP3", "BCI 6000").
                      Replace("GROUP4", "BCI 60")

            Else
                t = t.Replace("MICROLLAM 2.0E", "LVL 2.0E").
                      Replace("TIMBERSTRAND 1.55E", "LVL 1.7E").
                      Replace("VERSA LAM 2.0E", "LVL 2.0E").
                      Replace("VERSA LAM 1.7E", "LVL 1.7E").
                      Replace("TJI 110", "GROUP1").
                      Replace("TJI 210", "GROUP2").
                      Replace("TJI 230", "GROUP3").
                      Replace("TJI 360", "GROUP4").
                      Replace("BCI 4500", "GROUP1").
                      Replace("BCI 5000", "GROUP2").
                      Replace("BCI 6000", "GROUP3").
                      Replace("BCI 60", "GROUP4")
            End If

            Return t
        End Function

        Private Function RollUp(items As List(Of Item)) As List(Of RolledUp)
            Dim rolled As New List(Of RolledUp)()
            Dim current As RolledUp = Nothing

            For Each it In items

                If current Is Nothing Then
                    current = New RolledUp With {.SortNum = it.SortNum, .Product = it.Product, .LengthFt = it.LengthFt, .QtyText = it.Qty.ToString()}
                    rolled.Add(current)
                    Continue For
                End If

                If current.SortNum = 2 OrElse current.SortNum = 3 Then
                    If String.Equals(current.Product, it.Product, StringComparison.OrdinalIgnoreCase) Then
                        current.LengthFt += it.LengthFt
                        current.QtyText = "LF"
                    Else
                        current = New RolledUp With {.SortNum = it.SortNum, .Product = it.Product, .LengthFt = it.LengthFt, .QtyText = it.Qty.ToString()}
                        rolled.Add(current)
                    End If
                Else
                    If String.Equals(current.Product, it.Product, StringComparison.OrdinalIgnoreCase) AndAlso current.LengthFt = it.LengthFt Then
                        Dim q As Integer
                        If Integer.TryParse(current.QtyText, q) Then
                            q += it.Qty
                            current.QtyText = q.ToString()
                        Else
                            current.QtyText = it.Qty.ToString()
                        End If
                    Else
                        current = New RolledUp With {.SortNum = it.SortNum, .Product = it.Product, .LengthFt = it.LengthFt, .QtyText = it.Qty.ToString()}
                        rolled.Add(current)
                    End If
                End If

            Next

            Return rolled
        End Function

        ' ------------------------- Block creation (UNIQUE name per run) -------------------------

        Private Sub BuildMmListBlockDefinition(db As Database, externalDwgPath As String, rolled As List(Of RolledUp), blockName As String)

            Using tr = db.TransactionManager.StartTransaction()

                Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                ' If template DWG exists, clone it into this drawing UNDER the unique name.
                If Not String.IsNullOrWhiteSpace(externalDwgPath) AndAlso File.Exists(externalDwgPath) Then

                    Using sideDb As New Database(False, True)
                        sideDb.ReadDwgFile(externalDwgPath, FileShare.Read, True, "")

                        ' IMPORTANT:
                        ' Insert the external DWG as a block def using *our* unique name.
                        ' This creates a new definition and avoids touching prior MMList_### blocks.
                        db.Insert(blockName, sideDb, True)
                    End Using
                End If

                ' Ensure block def exists (create blank if template missing)
                If Not bt.Has(blockName) Then
                    bt.UpgradeOpen()
                    Dim btrNew As New BlockTableRecord With {.Name = blockName}
                    bt.Add(btrNew)
                    tr.AddNewlyCreatedDBObject(btrNew, True)
                End If

                Dim btr = CType(tr.GetObject(bt(blockName), OpenMode.ForWrite), BlockTableRecord)

                ' Clear block contents
                Dim toErase As New List(Of ObjectId)()
                For Each id In btr
                    toErase.Add(id)
                Next
                For Each id In toErase
                    Dim dbo = tr.GetObject(id, OpenMode.ForWrite)
                    dbo.Erase(True)
                Next

                ' Build the grid + text inside the block definition

                Dim startPt As New Point2d(30, 0)
                Dim endPt As New Point2d(228, 0)

                Dim textHeight As Double = 6.0

                Dim text2 As New Point2d(100, -11)
                Dim text3 As New Point2d(186.5, -11)
                Dim text4 As New Point2d(214, -11)

                ' Header line + headers
                AddLine(btr, tr, New Point3d(startPt.X, startPt.Y, 0), New Point3d(endPt.X, endPt.Y, 0))
                AddHeaderText(btr, tr, "PRODUCT", text2, textHeight)
                AddHeaderText(btr, tr, "LENGTH", text3, textHeight)
                AddHeaderText(btr, tr, "QTY", text4, textHeight)

                ' Next horizontal line (under header)
                startPt = New Point2d(startPt.X, startPt.Y - 11)
                endPt = New Point2d(endPt.X, endPt.Y - 11)
                text2 = New Point2d(text2.X, text2.Y - 11)
                text3 = New Point2d(text3.X, text3.Y - 11)
                text4 = New Point2d(text4.X, text4.Y - 11)
                AddLine(btr, tr, New Point3d(startPt.X, startPt.Y, 0), New Point3d(endPt.X, endPt.Y, 0))

                For i As Integer = 0 To rolled.Count - 1

                    ' gap line when sort group changes (mimics your VBA behavior)
                    If i >= 2 AndAlso rolled(i - 1).SortNum < rolled(i).SortNum Then
                        startPt = New Point2d(startPt.X, startPt.Y - 11)
                        endPt = New Point2d(endPt.X, endPt.Y - 11)
                        text2 = New Point2d(text2.X, text2.Y - 11)
                        text3 = New Point2d(text3.X, text3.Y - 11)
                        text4 = New Point2d(text4.X, text4.Y - 11)
                        AddLine(btr, tr, New Point3d(startPt.X, startPt.Y, 0), New Point3d(endPt.X, endPt.Y, 0))
                    End If

                    AddCellText(btr, tr, rolled(i).Product, text2, textHeight)
                    AddCellText(btr, tr, rolled(i).LengthFt.ToString(), text3, textHeight)
                    AddCellText(btr, tr, rolled(i).QtyText, text4, textHeight)

                    ' Line under this row
                    startPt = New Point2d(startPt.X, startPt.Y - 11)
                    endPt = New Point2d(endPt.X, endPt.Y - 11)
                    text2 = New Point2d(text2.X, text2.Y - 11)
                    text3 = New Point2d(text3.X, text3.Y - 11)
                    text4 = New Point2d(text4.X, text4.Y - 11)
                    AddLine(btr, tr, New Point3d(startPt.X, startPt.Y, 0), New Point3d(endPt.X, endPt.Y, 0))
                Next

                ' Vertical lines like VBA: extend EXACTLY to the last drawn horizontal line.
                Dim bottomY As Double = startPt.Y

                AddLine(btr, tr, New Point3d(30, 0, 0), New Point3d(30, bottomY, 0))
                AddLine(btr, tr, New Point3d(171, 0, 0), New Point3d(171, bottomY, 0))
                AddLine(btr, tr, New Point3d(202, 0, 0), New Point3d(202, bottomY, 0))
                AddLine(btr, tr, New Point3d(228, 0, 0), New Point3d(228, bottomY, 0))

                tr.Commit()
            End Using
        End Sub

        Private Sub AddLine(btr As BlockTableRecord, tr As Transaction, p1 As Point3d, p2 As Point3d)
            Dim ln As New Line(p1, p2)
            btr.AppendEntity(ln)
            tr.AddNewlyCreatedDBObject(ln, True)
        End Sub

        Private Sub AddHeaderText(btr As BlockTableRecord, tr As Transaction, value As String, pt As Point2d, height As Double)
            Dim t As New DBText()
            t.TextString = value
            t.Height = height
            t.Position = New Point3d(pt.X, pt.Y, 0)
            t.HorizontalMode = TextHorizontalMode.TextCenter
            t.VerticalMode = TextVerticalMode.TextBottom
            t.AlignmentPoint = New Point3d(pt.X, pt.Y, 0)
            btr.AppendEntity(t)
            tr.AddNewlyCreatedDBObject(t, True)
        End Sub

        Private Sub AddCellText(btr As BlockTableRecord, tr As Transaction, value As String, pt As Point2d, height As Double)
            Dim t As New DBText()
            t.TextString = value
            t.Height = height
            t.Position = New Point3d(pt.X, pt.Y, 0)
            t.HorizontalMode = TextHorizontalMode.TextCenter
            t.VerticalMode = TextVerticalMode.TextBottom
            t.AlignmentPoint = New Point3d(pt.X, pt.Y, 0)
            btr.AppendEntity(t)
            tr.AddNewlyCreatedDBObject(t, True)
        End Sub

    End Class

End Namespace
