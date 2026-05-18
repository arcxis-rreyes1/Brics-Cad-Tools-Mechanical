Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.Cealing_All))>
Namespace Arcxis_Cad_Tools
    Public Class Cealing_All

        Private Shared _lastSpacing As String = "24"
        Private Shared _lastBearing As String = "Y"
        Private Shared _lastCeilingOrRidge As String = "C"

        Private Enum CeilingTable
            ZZ
            ZZZ
            ZZ3
            ZZZ3
        End Enum

        Private Enum SizeTable
            CS
            CS3
            CSS
        End Enum

        <CommandMethod("zz", CommandFlags.Modal)>
        Public Sub CeilingLongLabels_ZZ()
            RunCeilingLongLabels(CeilingTable.ZZ)
        End Sub

        <CommandMethod("zzz", CommandFlags.Modal)>
        Public Sub CeilingLongLabels_ZZZ()
            RunCeilingLongLabels(CeilingTable.ZZZ)
        End Sub

        <CommandMethod("zz3", CommandFlags.Modal)>
        Public Sub CeilingLongLabels_ZZ3()
            RunCeilingLongLabels(CeilingTable.ZZ3)
        End Sub

        <CommandMethod("zzz3", CommandFlags.Modal)>
        Public Sub CeilingLongLabels_ZZZ3()
            RunCeilingLongLabels(CeilingTable.ZZZ3)
        End Sub

        ' <CommandMethod("cl", CommandFlags.Modal)>
        ' Public Sub CeilingLengthLabels_CL()
        '     RunClLengthLabels()
        ' End Sub

        ' <CommandMethod("cs", CommandFlags.Modal)>
        ' Public Sub CeilingSizeLabels_CS()
        '     RunSizeLabels(SizeTable.CS)
        ' End Sub

        ' <CommandMethod("cs3", CommandFlags.Modal)>
        ' Public Sub CeilingSizeLabels_CS3()
        '     RunSizeLabels(SizeTable.CS3)
        ' End Sub

        ' <CommandMethod("css", CommandFlags.Modal)>
        ' Public Sub CeilingSizeLabels_CSS()
        '     RunSizeLabels(SizeTable.CSS)
        ' End Sub

        Private Sub RunCeilingLongLabels(table As CeilingTable)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim spacing As String = PromptForSpacing(ed)
            If String.IsNullOrEmpty(spacing) Then
                Return
            End If

            Dim hasBearing As Boolean? = PromptForBearing(ed)
            If Not hasBearing.HasValue Then
                Return
            End If

            Dim oldLayer As String = CStr(Application.GetSystemVariable("CLAYER"))
            Dim oldTextSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))

            Try
                Using docLock As DocumentLock = doc.LockDocument()
                    ' First transaction: ensure layers and isolate them
                    Dim layerStates As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CJOIST")
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CLG-JOIST")
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CLG-TAG")
                        Application.SetSystemVariable("TEXTSIZE", 6.0)

                        ' Isolate layers and save their states
                        layerStates = IsolateSelectionLayers(tr, db, "S-FRM-CJOIST", "S-FRM-CLG-JOIST", "S-FRM-CLG-TAG")
                        tr.Commit()
                    End Using

                    ' Force regen to display layer changes
                    ed.Regen()

                    ' Now do selection with layers isolated
                    Dim filter As New SelectionFilter(New TypedValue() {
                        New TypedValue(CInt(DxfCode.Start), "LINE,LWPOLYLINE,POLYLINE")
                    })

                    Dim selRes As PromptSelectionResult = ed.GetSelection(filter)

                    ' Restore layers in another transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        RestoreLayerStates(tr, layerStates)
                        tr.Commit()
                    End Using

                    If selRes.Status <> PromptStatus.OK Then
                        Return
                    End If

                    ' Continue with processing in a final transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()

                        If table = CeilingTable.ZZ3 OrElse table = CeilingTable.ZZZ3 Then
                            EnsureLayerFromTemplateAndColor(tr, db, "S-FRM-CLG-JOIST-3", "S-FRM-CLG-JOIST", 20)
                            MoveEntitiesToLayer(selRes.Value.GetObjectIds(), tr, "S-FRM-CLG-JOIST-3")
                        End If

                        Application.SetSystemVariable("CLAYER", "S-FRM-CLG-TAG")

                        Dim btr As BlockTableRecord = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Dim textSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))
                        Dim segments As List(Of Tuple(Of Point3d, Point3d)) = CollectLinearSegments(selRes.Value.GetObjectIds(), tr)
                        Dim groupedSegments As List(Of List(Of Tuple(Of Point3d, Point3d))) = BuildParallelGroups(segments, 2.2)

                        For Each group As List(Of Tuple(Of Point3d, Point3d)) In groupedSegments
                            Dim sequences As List(Of List(Of Tuple(Of Point3d, Point3d))) = SplitIntoSequences(group, 2.2, 4)

                            For Each sequence As List(Of Tuple(Of Point3d, Point3d)) In sequences
                                Dim segment As Tuple(Of Point3d, Point3d) = PickRepresentativeSegment(sequence)
                                Dim p10 As Point3d = segment.Item1
                                Dim p11 As Point3d = segment.Item2
                                Dim midpoint As New Point3d((p10.X + p11.X) * 0.5, (p10.Y + p11.Y) * 0.5, (p10.Z + p11.Z) * 0.5)

                                Dim clearSpan As Double = New Point2d(p10.X, p10.Y).GetDistanceTo(New Point2d(p11.X, p11.Y))

                                ' Bearing affects member size lookups only.
                                Dim sizeSpan As Double = If(hasBearing.Value, clearSpan - 4.0, clearSpan)
                                Dim prefix As String = GetMemberPrefix(spacing, sizeSpan, table)

                                ' Length rule: actual span for bearing, or +6 in when no bearing; then round up to even feet.
                                Dim lengthSpan As Double = If(hasBearing.Value, clearSpan, clearSpan + 6.0)
                                Dim lengthFeet As Integer = CInt(Math.Ceiling(lengthSpan / 12.0))
                                Dim taggedFeet As Integer
                                If lengthFeet < 6 Then
                                    taggedFeet = lengthFeet
                                Else
                                    taggedFeet = If(lengthFeet Mod 2 = 0, lengthFeet, lengthFeet + 1)
                                End If
                                Dim oldLabel As String = BuildLabel(prefix, taggedFeet, table)

                                Dim forward As Vector3d = p10.GetVectorTo(p11)
                                Dim rawAngle As Double = Math.Atan2(forward.Y, forward.X)
                                If rawAngle < 0 Then
                                    rawAngle += Math.PI * 2.0
                                End If
                                Dim angleDeg As Double = rawAngle * (180.0 / Math.PI)
                                If angleDeg > 90.1 AndAlso angleDeg <= 270.1 Then
                                    Dim reverse As Vector3d = p11.GetVectorTo(p10)
                                    rawAngle = Math.Atan2(reverse.Y, reverse.X)
                                    If rawAngle < 0 Then
                                        rawAngle += Math.PI * 2.0
                                    End If
                                End If

                                Dim textPoint As Point3d = GetTagTextPoint(midpoint, p10, p11, rawAngle, textSize)

                                Dim text As New DBText()
                                text.SetDatabaseDefaults()
                                text.Layer = CStr(Application.GetSystemVariable("CLAYER"))
                                text.TextStyleId = db.Textstyle
                                text.Height = textSize
                                text.WidthFactor = 0.8
                                text.Rotation = rawAngle
                                text.HorizontalMode = TextHorizontalMode.TextMid
                                text.AlignmentPoint = textPoint
                                text.Position = textPoint
                                If sequence.Count > 1 Then
                                    text.TextString = sequence.Count.ToString() & "~" & oldLabel
                                Else
                                    text.TextString = oldLabel
                                End If
                                text.AdjustAlignment(db)

                                btr.AppendEntity(text)
                                tr.AddNewlyCreatedDBObject(text, True)
                            Next
                        Next

                        tr.Commit()
                    End Using
                End Using
            Finally
                Application.SetSystemVariable("CLAYER", oldLayer)
                Application.SetSystemVariable("TEXTSIZE", oldTextSize)
            End Try
        End Sub

        Private Sub RunClLengthLabels()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim mode As String = PromptForCeilingOrRidge(ed)
            If String.IsNullOrEmpty(mode) Then
                Return
            End If

            Dim oldLayer As String = CStr(Application.GetSystemVariable("CLAYER"))
            Dim oldTextSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))

            Try
                Using docLock As DocumentLock = doc.LockDocument()
                    Dim sourceLayer As String = If(mode = "R", "S-FRM-ROOF", "S-FRM-CLG-JOIST")
                    Dim layerStates As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))

                    ' First transaction: ensure layers and isolate them
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        EnsureLayerExistsAndOn(tr, db, sourceLayer)
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CLG-TAG")

                        ' Ensure both framing layers exist for isolation
                        If sourceLayer <> "S-FRM-ROOF" Then
                            EnsureLayerExistsAndOn(tr, db, "S-FRM-CJOIST")
                        End If

                        Application.SetSystemVariable("TEXTSIZE", 6.0)

                        ' Isolate layers and save their states
                        If sourceLayer = "S-FRM-ROOF" Then
                            layerStates = IsolateSelectionLayer(tr, db, sourceLayer)
                        Else
                            layerStates = IsolateSelectionLayers(tr, db, "S-FRM-CJOIST", "S-FRM-CLG-JOIST")
                        End If
                        tr.Commit()
                    End Using

                    ' Force regen to display layer changes
                    ed.Regen()

                    ' Now do selection with layers isolated
                    Dim filter As New SelectionFilter(New TypedValue() {
                        New TypedValue(CInt(DxfCode.Start), "LINE,LWPOLYLINE,POLYLINE")
                    })

                    Dim selRes As PromptSelectionResult = ed.GetSelection(filter)

                    ' Restore layers in another transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        RestoreLayerStates(tr, layerStates)
                        tr.Commit()
                    End Using

                    If selRes.Status <> PromptStatus.OK Then
                        Return
                    End If

                    ' Continue with processing in a final transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()

                        Application.SetSystemVariable("CLAYER", "S-FRM-CLG-TAG")

                        Dim btr As BlockTableRecord = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Dim textSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))
                        Dim segments As List(Of Tuple(Of Point3d, Point3d)) = CollectLinearSegments(selRes.Value.GetObjectIds(), tr)
                        Dim groupedSegments As List(Of List(Of Tuple(Of Point3d, Point3d))) = BuildParallelGroups(segments, 2.2)

                        For Each group As List(Of Tuple(Of Point3d, Point3d)) In groupedSegments
                            Dim segment As Tuple(Of Point3d, Point3d) = PickRepresentativeSegment(group)
                            Dim p10 As Point3d = segment.Item1
                            Dim p11 As Point3d = segment.Item2
                            Dim midpoint As New Point3d((p10.X + p11.X) * 0.5, (p10.Y + p11.Y) * 0.5, (p10.Z + p11.Z) * 0.5)

                            Dim d2d As Double = New Point2d(p10.X, p10.Y).GetDistanceTo(New Point2d(p11.X, p11.Y))
                            Dim d4d As Integer = CInt(Math.Truncate(1.0 + (d2d / 12.0)))
                            Dim d5d As Integer = If(d4d <= 7, d4d, MakeEven(d4d))

                            Dim forward As Vector3d = p10.GetVectorTo(p11)
                            Dim rawAngle As Double = Math.Atan2(forward.Y, forward.X)
                            If rawAngle < 0 Then
                                rawAngle += Math.PI * 2.0
                            End If
                            Dim angleDeg As Double = rawAngle * (180.0 / Math.PI)
                            If angleDeg > 90.1 AndAlso angleDeg <= 270.1 Then
                                Dim reverse As Vector3d = p11.GetVectorTo(p10)
                                rawAngle = Math.Atan2(reverse.Y, reverse.X)
                                If rawAngle < 0 Then
                                    rawAngle += Math.PI * 2.0
                                End If
                            End If

                            Dim textPoint As Point3d = GetTagTextPoint(midpoint, p10, p11, rawAngle, textSize)

                            Dim text As New DBText()
                            text.SetDatabaseDefaults()
                            text.Layer = CStr(Application.GetSystemVariable("CLAYER"))
                            text.TextStyleId = db.Textstyle
                            text.Height = textSize
                            text.WidthFactor = 0.8
                            text.Rotation = rawAngle
                            text.HorizontalMode = TextHorizontalMode.TextMid
                            text.AlignmentPoint = textPoint
                            text.Position = textPoint
                            text.TextString = group.Count.ToString() & " ~ " & d5d.ToString()
                            text.AdjustAlignment(db)

                            btr.AppendEntity(text)
                            tr.AddNewlyCreatedDBObject(text, True)
                        Next

                        tr.Commit()
                    End Using
                End Using
            Finally
                Application.SetSystemVariable("CLAYER", oldLayer)
                Application.SetSystemVariable("TEXTSIZE", oldTextSize)
            End Try
        End Sub

        Private Sub RunSizeLabels(table As SizeTable)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim spacing As String = PromptForSpacing(ed)
            If String.IsNullOrEmpty(spacing) Then
                Return
            End If

            Dim oldLayer As String = CStr(Application.GetSystemVariable("CLAYER"))
            Dim oldTextSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))

            Try
                Using docLock As DocumentLock = doc.LockDocument()
                    Dim layerStates As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))

                    ' First transaction: ensure layers and isolate them
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CJOIST")
                        EnsureLayerExistsAndOn(tr, db, "S-FRM-CLG-JOIST")
                        EnsureLayerExistsAndOn(tr, db, "s-anno-egtext")
                        Application.SetSystemVariable("TEXTSIZE", 6.0)

                        ' Isolate layers and save their states
                        layerStates = IsolateSelectionLayers(tr, db, "S-FRM-CJOIST", "S-FRM-CLG-JOIST")
                        tr.Commit()
                    End Using

                    ' Force regen to display layer changes
                    ed.Regen()

                    ' Now do selection with layers isolated
                    Dim filter As New SelectionFilter(New TypedValue() {
                        New TypedValue(CInt(DxfCode.Start), "LINE,LWPOLYLINE,POLYLINE")
                    })

                    Dim selRes As PromptSelectionResult = ed.GetSelection(filter)

                    ' Restore layers in another transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        RestoreLayerStates(tr, layerStates)
                        tr.Commit()
                    End Using

                    If selRes.Status <> PromptStatus.OK Then
                        Return
                    End If

                    ' Continue with processing in a final transaction
                    Using tr As Transaction = db.TransactionManager.StartTransaction()

                        Application.SetSystemVariable("CLAYER", "s-anno-egtext")

                        Dim btr As BlockTableRecord = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Dim textSize As Double = CDbl(Application.GetSystemVariable("TEXTSIZE"))
                        Dim segments As List(Of Tuple(Of Point3d, Point3d)) = CollectLinearSegments(selRes.Value.GetObjectIds(), tr)

                        For Each segment As Tuple(Of Point3d, Point3d) In segments
                            Dim p10 As Point3d = segment.Item1
                            Dim p11 As Point3d = segment.Item2
                            Dim midpoint As New Point3d((p10.X + p11.X) * 0.5, (p10.Y + p11.Y) * 0.5, (p10.Z + p11.Z) * 0.5)

                            Dim clearSpan As Double = New Point2d(p10.X, p10.Y).GetDistanceTo(New Point2d(p11.X, p11.Y))
                            Dim d2d As Double = If(table = SizeTable.CSS, clearSpan, clearSpan - 7.0)
                            Dim label As String = GetSizeLabel(spacing, d2d, table)

                            Dim forward As Vector3d = p10.GetVectorTo(p11)
                            Dim rawAngle As Double = Math.Atan2(forward.Y, forward.X)
                            If rawAngle < 0 Then
                                rawAngle += Math.PI * 2.0
                            End If
                            Dim angleDeg As Double = rawAngle * (180.0 / Math.PI)
                            If angleDeg > 90.1 AndAlso angleDeg <= 270.1 Then
                                Dim reverse As Vector3d = p11.GetVectorTo(p10)
                                rawAngle = Math.Atan2(reverse.Y, reverse.X)
                                If rawAngle < 0 Then
                                    rawAngle += Math.PI * 2.0
                                End If
                            End If

                            Dim textPoint As Point3d = GetTagTextPoint(midpoint, p10, p11, rawAngle, textSize)

                            Dim text As New DBText()
                            text.SetDatabaseDefaults()
                            text.Layer = CStr(Application.GetSystemVariable("CLAYER"))
                            text.TextStyleId = db.Textstyle
                            text.Height = textSize
                            text.WidthFactor = 0.8
                            text.Rotation = rawAngle
                            text.HorizontalMode = TextHorizontalMode.TextMid
                            text.AlignmentPoint = textPoint
                            text.Position = textPoint
                            text.TextString = label
                            text.AdjustAlignment(db)

                            btr.AppendEntity(text)
                            tr.AddNewlyCreatedDBObject(text, True)
                        Next

                        tr.Commit()
                    End Using
                End Using
            Finally
                Application.SetSystemVariable("CLAYER", oldLayer)
                Application.SetSystemVariable("TEXTSIZE", oldTextSize)
            End Try
        End Sub

        Private Function PromptForSpacing(ed As Editor) As String
            Dim pso As New PromptStringOptions(vbLf & "What is the spacing? (24,19.2,16,12) <" & _lastSpacing & ">: ")
            pso.AllowSpaces = False
            Dim result As PromptResult = ed.GetString(pso)
            If result.Status <> PromptStatus.OK Then
                Return Nothing
            End If

            Dim value As String = result.StringResult.Trim()
            If value = "" Then
                value = _lastSpacing
            End If

            Select Case value
                Case "24", "19.2", "16", "12"
                    _lastSpacing = value
                    Return value
                Case Else
                    ed.WriteMessage(vbLf & "Invalid spacing. Use 24, 19.2, 16, or 12.")
                    Return Nothing
            End Select
        End Function

        Private Function PromptForBearing(ed As Editor) As Boolean?
            Dim pbo As New PromptStringOptions(vbLf & "Ceiling have bearing Yes or No? (Y/N) <" & _lastBearing & ">: ")
            pbo.AllowSpaces = False

            Dim result As PromptResult = ed.GetString(pbo)
            If result.Status <> PromptStatus.OK Then
                Return Nothing
            End If

            Dim value As String = result.StringResult.Trim().ToUpperInvariant()
            If value = "" Then
                value = _lastBearing
            End If

            If value = "Y" Then
                _lastBearing = "Y"
                Return True
            End If

            If value = "N" Then
                _lastBearing = "N"
                Return False
            End If

            ed.WriteMessage(vbLf & "Invalid entry. Use Y or N.")
            Return Nothing
        End Function

        Private Function PromptForCeilingOrRidge(ed As Editor) As String
            Dim pbo As New PromptStringOptions(vbLf & "Ceiling or Ridges?(C/R) <" & _lastCeilingOrRidge & ">: ")
            pbo.AllowSpaces = False

            Dim result As PromptResult = ed.GetString(pbo)
            If result.Status <> PromptStatus.OK Then
                Return Nothing
            End If

            Dim value As String = result.StringResult.Trim().ToUpperInvariant()
            If value = "" Then
                value = _lastCeilingOrRidge
            End If

            If value = "C" OrElse value = "R" Then
                _lastCeilingOrRidge = value
                Return value
            End If

            ed.WriteMessage(vbLf & "Invalid entry. Use C or R.")
            Return Nothing
        End Function

        Private Function GetMemberPrefix(spacing As String, d2d As Double, table As CeilingTable) As String
            Select Case table
                Case CeilingTable.ZZ
                    Select Case spacing
                        Case "24" : Return PickPrefix(d2d, 118.0, 150.0, 177.0, 209.0)
                        Case "19.2" : Return PickPrefix(d2d, 132.0, 167.0, 198.0, 234.0)
                        Case "16" : Return PickPrefix(d2d, 144.0, 183.0, 217.0, 256.0)
                        Case "12" : Return PickPrefix(d2d, 167.0, 211.0, 251.0, 296.0)
                    End Select
                Case CeilingTable.ZZZ
                    Select Case spacing
                        Case "24" : Return PickPrefix(d2d, 167.0, 211.0, 251.0, 296.0)
                        Case "19.2" : Return PickPrefix(d2d, 187.0, 236.0, 281.0, 330.0)
                        Case "16" : Return PickPrefix(d2d, 203.0, 259.0, 307.0, 362.0)
                        Case "12" : Return PickPrefix(d2d, 224.0, 295.0, 329.0, 400.0)
                    End Select
                Case CeilingTable.ZZ3
                    ' Source table: ARC_Ceiling_All_Brics.lsp, defun c:zz3
                    Select Case spacing
                        Case "24" : Return PickPrefix(d2d, 89.0, 113.0, 137.0, 162.0)
                        Case "19.2" : Return PickPrefix(d2d, 100.0, 126.0, 153.0, 181.0)
                        Case "16" : Return PickPrefix(d2d, 110.0, 138.0, 167.0, 198.0)
                        Case "12" : Return PickPrefix(d2d, 127.0, 159.0, 193.0, 229.0)
                    End Select
                Case CeilingTable.ZZZ3
                    ' Source table: ARC_Ceiling_All_Brics.lsp, defun c:zzz3
                    Select Case spacing
                        Case "24" : Return PickPrefix(d2d, 126.0, 159.0, 193.0, 229.0)
                        Case "19.2" : Return PickPrefix(d2d, 141.0, 178.0, 216.0, 256.0)
                        Case "16" : Return PickPrefix(d2d, 155.0, 195.0, 237.0, 280.0)
                        Case "12" : Return PickPrefix(d2d, 179.0, 216.0, 273.0, 288.0)
                    End Select
            End Select

            Return ""
        End Function

        Private Function BuildLabel(prefix As String, feet As Integer, table As CeilingTable) As String
            Dim label As String = prefix & feet.ToString() & "'"
            Return label
        End Function

        Private Function GetSizeLabel(spacing As String, d2d As Double, table As SizeTable) As String
            Select Case table
                Case SizeTable.CS
                    Select Case spacing
                        Case "24" : Return PickSizeLabel(d2d, 118.0, 150.0, 177.0, 209.0, False)
                        Case "19.2" : Return PickSizeLabel(d2d, 132.0, 167.0, 198.0, 234.0, False)
                        Case "16" : Return PickSizeLabel(d2d, 144.0, 183.0, 217.0, 256.0, False)
                        Case "12" : Return PickSizeLabel(d2d, 167.0, 211.0, 251.0, 296.0, False)
                    End Select
                Case SizeTable.CS3
                    Select Case spacing
                        Case "24" : Return PickSizeLabel(d2d, 89.0, 113.0, 137.0, 162.0, False)
                        Case "19.2" : Return PickSizeLabel(d2d, 100.0, 126.0, 153.0, 181.0, True)
                        Case "16" : Return PickSizeLabel(d2d, 110.0, 138.0, 167.0, 198.0, True)
                        Case "12" : Return PickSizeLabel(d2d, 127.0, 159.0, 193.0, 229.0, True)
                    End Select
                Case SizeTable.CSS
                    Select Case spacing
                        Case "24" : Return PickSizeLabel(d2d, 118.0, 150.0, 177.0, 209.0, False)
                        Case "19.2" : Return PickSizeLabel(d2d, 132.0, 167.0, 198.0, 234.0, False)
                        Case "16" : Return PickSizeLabel(d2d, 144.0, 183.0, 217.0, 256.0, False)
                        Case "12" : Return PickSizeLabel(d2d, 167.0, 211.0, 251.0, 296.0, False)
                    End Select
            End Select

            Return "Change Spacing"
        End Function

        Private Function PickSizeLabel(d2d As Double, max2x6 As Double, max2x8 As Double, max2x10 As Double, max2x12 As Double, addNo3 As Boolean) As String
            If d2d <= max2x6 Then
                Return WithNo3("2x6", addNo3)
            End If
            If d2d <= max2x8 Then
                Return WithNo3("2x8", addNo3)
            End If
            If d2d <= max2x10 Then
                Return WithNo3("2x10", addNo3)
            End If
            If d2d <= max2x12 Then
                Return WithNo3("2x12", addNo3)
            End If

            Return "Change Spacing"
        End Function

        Private Function WithNo3(value As String, addNo3 As Boolean) As String
            Return value
        End Function

        Private Function PickPrefix(d2d As Double, maxBlank As Double, max2x8 As Double, max2x10 As Double, max2x12 As Double) As String
            If d2d <= maxBlank Then
                Return ""
            End If
            If d2d <= max2x8 Then
                Return "2x8-"
            End If
            If d2d <= max2x10 Then
                Return "2x10-"
            End If
            If d2d <= max2x12 Then
                Return "2x12-"
            End If
            If d2d <= 500.0 Then
                Return "Change Spacing"
            End If

            Return "Change Spacing"
        End Function

        Private Shared Function MakeEven(value As Integer) As Integer
            If value Mod 2 = 0 Then
                Return value
            End If

            Return value + 1
        End Function

        Private Iterator Function GetLinearSegments(ent As Entity) As IEnumerable(Of Tuple(Of Point3d, Point3d))
            Dim ln As Line = TryCast(ent, Line)
            If ln IsNot Nothing Then
                Yield Tuple.Create(ln.StartPoint, ln.EndPoint)
                Return
            End If

            Dim pl As Polyline = TryCast(ent, Polyline)
            If pl Is Nothing OrElse pl.NumberOfVertices < 2 Then
                Return
            End If

            Dim segmentCount As Integer = If(pl.Closed, pl.NumberOfVertices, pl.NumberOfVertices - 1)
            For i As Integer = 0 To segmentCount - 1
                If pl.GetSegmentType(i) <> SegmentType.Line Then
                    Continue For
                End If

                Dim p0 As Point3d = pl.GetPoint3dAt(i)
                Dim p1 As Point3d = pl.GetPoint3dAt((i + 1) Mod pl.NumberOfVertices)
                Yield Tuple.Create(p0, p1)
            Next
        End Function

        Private Function CollectLinearSegments(ids As IEnumerable(Of ObjectId), tr As Transaction) As List(Of Tuple(Of Point3d, Point3d))
            Dim result As New List(Of Tuple(Of Point3d, Point3d))()

            For Each id As ObjectId In ids
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing Then
                    Continue For
                End If

                For Each segment As Tuple(Of Point3d, Point3d) In GetLinearSegments(ent)
                    result.Add(segment)
                Next
            Next

            Return result
        End Function

        Private Function BuildParallelGroups(segments As List(Of Tuple(Of Point3d, Point3d)), maxSpacing As Double) As List(Of List(Of Tuple(Of Point3d, Point3d)))
            Dim groups As New List(Of List(Of Tuple(Of Point3d, Point3d)))()
            Dim visited As New HashSet(Of Integer)()

            For i As Integer = 0 To segments.Count - 1
                If visited.Contains(i) Then
                    Continue For
                End If

                Dim queue As New Queue(Of Integer)()
                Dim indices As New List(Of Integer)()
                queue.Enqueue(i)
                visited.Add(i)

                While queue.Count > 0
                    Dim current As Integer = queue.Dequeue()
                    indices.Add(current)

                    For j As Integer = 0 To segments.Count - 1
                        If visited.Contains(j) Then
                            Continue For
                        End If

                        If Not AreParallelWithinTolerance(segments(current), segments(j), 5.0) Then
                            Continue For
                        End If

                        Dim spacing As Double = PerpendicularSpacing2d(segments(current), segments(j))
                        Dim endpointDistance As Double = MinEndpointDistance2d(segments(current), segments(j))
                        Dim horizontal As Boolean = IsMostlyHorizontal(segments(current))
                        Dim hasAxisOverlap As Boolean = HasStrongAxisOverlap(segments(current), segments(j), horizontal)

                        If Not hasAxisOverlap Then
                            Continue For
                        End If

                        If spacing > maxSpacing AndAlso endpointDistance > maxSpacing Then
                            Continue For
                        End If

                        visited.Add(j)
                        queue.Enqueue(j)
                    Next
                End While

                Dim group As New List(Of Tuple(Of Point3d, Point3d))()
                For Each idx As Integer In indices
                    group.Add(segments(idx))
                Next
                groups.Add(group)
            Next

            Return groups
        End Function

        Private Function SplitIntoSequences(group As List(Of Tuple(Of Point3d, Point3d)), maxSpacing As Double, maxCount As Integer) As List(Of List(Of Tuple(Of Point3d, Point3d)))
            Dim result As New List(Of List(Of Tuple(Of Point3d, Point3d)))()
            If group.Count = 0 Then
                Return result
            End If

            Dim horizontal As Boolean = IsMostlyHorizontal(group(0))
            Dim ordered As List(Of Tuple(Of Point3d, Point3d)) = group.OrderBy(Function(s) GetPerpendicularAxisValue(s, horizontal)).ToList()

            Dim current As New List(Of Tuple(Of Point3d, Point3d))()
            current.Add(ordered(0))

            For i As Integer = 1 To ordered.Count - 1
                Dim previousValue As Double = GetPerpendicularAxisValue(ordered(i - 1), horizontal)
                Dim currentValue As Double = GetPerpendicularAxisValue(ordered(i), horizontal)
                Dim gap As Double = Math.Abs(currentValue - previousValue)
                Dim endpointDistance As Double = MinEndpointDistance2d(ordered(i - 1), ordered(i))
                Dim hasAxisOverlap As Boolean = HasStrongAxisOverlap(ordered(i - 1), ordered(i), horizontal)

                If (Not hasAxisOverlap) OrElse ((gap > maxSpacing AndAlso endpointDistance > maxSpacing) OrElse current.Count >= maxCount) Then
                    result.Add(current)
                    current = New List(Of Tuple(Of Point3d, Point3d))()
                End If

                current.Add(ordered(i))
            Next

            If current.Count > 0 Then
                result.Add(current)
            End If

            Return result
        End Function

        Private Function GetPerpendicularAxisValue(segment As Tuple(Of Point3d, Point3d), horizontal As Boolean) As Double
            If horizontal Then
                Return (segment.Item1.Y + segment.Item2.Y) * 0.5
            End If

            Return (segment.Item1.X + segment.Item2.X) * 0.5
        End Function

        Private Function AreParallelWithinTolerance(a As Tuple(Of Point3d, Point3d), b As Tuple(Of Point3d, Point3d), angleToleranceDeg As Double) As Boolean
            Dim va As Vector3d = a.Item1.GetVectorTo(a.Item2)
            Dim vb As Vector3d = b.Item1.GetVectorTo(b.Item2)
            If va.Length <= 0.000000001 OrElse vb.Length <= 0.000000001 Then
                Return False
            End If

            Dim aa As Double = Math.Atan2(va.Y, va.X)
            Dim ab As Double = Math.Atan2(vb.Y, vb.X)
            Dim delta As Double = Math.Abs(aa - ab)
            If delta > Math.PI Then
                delta = (2.0 * Math.PI) - delta
            End If
            If delta > (Math.PI * 0.5) Then
                delta = Math.Abs(Math.PI - delta)
            End If

            Return delta <= (angleToleranceDeg * (Math.PI / 180.0))
        End Function

        Private Function PerpendicularSpacing2d(baseSegment As Tuple(Of Point3d, Point3d), testSegment As Tuple(Of Point3d, Point3d)) As Double
            Dim dx As Double = baseSegment.Item2.X - baseSegment.Item1.X
            Dim dy As Double = baseSegment.Item2.Y - baseSegment.Item1.Y
            Dim length As Double = Math.Sqrt((dx * dx) + (dy * dy))
            If length <= 0.000000001 Then
                Return Double.MaxValue
            End If

            Dim ux As Double = dx / length
            Dim uy As Double = dy / length
            Dim nx As Double = -uy
            Dim ny As Double = ux

            Dim baseMid As New Point3d((baseSegment.Item1.X + baseSegment.Item2.X) * 0.5, (baseSegment.Item1.Y + baseSegment.Item2.Y) * 0.5, 0.0)
            Dim testMid As New Point3d((testSegment.Item1.X + testSegment.Item2.X) * 0.5, (testSegment.Item1.Y + testSegment.Item2.Y) * 0.5, 0.0)
            Dim vx As Double = testMid.X - baseMid.X
            Dim vy As Double = testMid.Y - baseMid.Y
            Return Math.Abs((vx * nx) + (vy * ny))
        End Function

        Private Function MinEndpointDistance2d(a As Tuple(Of Point3d, Point3d), b As Tuple(Of Point3d, Point3d)) As Double
            Dim d1 As Double = Distance2d(a.Item1, b.Item1)
            Dim d2 As Double = Distance2d(a.Item1, b.Item2)
            Dim d3 As Double = Distance2d(a.Item2, b.Item1)
            Dim d4 As Double = Distance2d(a.Item2, b.Item2)
            Return Math.Min(Math.Min(d1, d2), Math.Min(d3, d4))
        End Function

        Private Function Distance2d(a As Point3d, b As Point3d) As Double
            Dim dx As Double = a.X - b.X
            Dim dy As Double = a.Y - b.Y
            Return Math.Sqrt((dx * dx) + (dy * dy))
        End Function

        Private Function HasStrongAxisOverlap(a As Tuple(Of Point3d, Point3d), b As Tuple(Of Point3d, Point3d), horizontal As Boolean) As Boolean
            Dim aStart As Double
            Dim aEnd As Double
            Dim bStart As Double
            Dim bEnd As Double

            If horizontal Then
                aStart = Math.Min(a.Item1.X, a.Item2.X)
                aEnd = Math.Max(a.Item1.X, a.Item2.X)
                bStart = Math.Min(b.Item1.X, b.Item2.X)
                bEnd = Math.Max(b.Item1.X, b.Item2.X)
            Else
                aStart = Math.Min(a.Item1.Y, a.Item2.Y)
                aEnd = Math.Max(a.Item1.Y, a.Item2.Y)
                bStart = Math.Min(b.Item1.Y, b.Item2.Y)
                bEnd = Math.Max(b.Item1.Y, b.Item2.Y)
            End If

            Dim overlapStart As Double = Math.Max(aStart, bStart)
            Dim overlapEnd As Double = Math.Min(aEnd, bEnd)
            Dim overlapLength As Double = Math.Max(0.0, overlapEnd - overlapStart)
            Dim aLength As Double = Math.Max(0.0, aEnd - aStart)
            Dim bLength As Double = Math.Max(0.0, bEnd - bStart)
            Dim minLength As Double = Math.Min(aLength, bLength)

            If minLength <= 0.000000001 Then
                Return False
            End If

            ' Require strong directional overlap to avoid grouping staggered members.
            Return overlapLength >= (0.9 * minLength)
        End Function

        Private Function PickRepresentativeSegment(group As List(Of Tuple(Of Point3d, Point3d))) As Tuple(Of Point3d, Point3d)
            Dim representative As Tuple(Of Point3d, Point3d) = group(0)
            Dim horizontal As Boolean = IsMostlyHorizontal(group(0))

            For Each candidate As Tuple(Of Point3d, Point3d) In group
                If horizontal Then
                    Dim yCandidate As Double = (candidate.Item1.Y + candidate.Item2.Y) * 0.5
                    Dim yCurrent As Double = (representative.Item1.Y + representative.Item2.Y) * 0.5
                    If yCandidate > yCurrent Then
                        representative = candidate
                    End If
                Else
                    Dim xCandidate As Double = (candidate.Item1.X + candidate.Item2.X) * 0.5
                    Dim xCurrent As Double = (representative.Item1.X + representative.Item2.X) * 0.5
                    If xCandidate < xCurrent Then
                        representative = candidate
                    End If
                End If
            Next

            Return representative
        End Function

        Private Function IsMostlyHorizontal(segment As Tuple(Of Point3d, Point3d)) As Boolean
            Dim dx As Double = Math.Abs(segment.Item2.X - segment.Item1.X)
            Dim dy As Double = Math.Abs(segment.Item2.Y - segment.Item1.Y)
            Return dx >= dy
        End Function

        Private Function GetTagTextPoint(midpoint As Point3d, p10 As Point3d, p11 As Point3d, rawAngle As Double, textSize As Double) As Point3d
            Dim dx As Double = Math.Abs(p11.X - p10.X)
            Dim dy As Double = Math.Abs(p11.Y - p10.Y)

            If dy > dx Then
                Dim vx As Double = p11.X - p10.X
                Dim vy As Double = p11.Y - p10.Y
                Dim length As Double = Math.Sqrt((vx * vx) + (vy * vy))
                If length > 0.000000001 Then
                    Dim n1x As Double = -vy / length
                    Dim n1y As Double = vx / length
                    Dim n2x As Double = -n1x
                    Dim n2y As Double = -n1y

                    Dim nx As Double
                    Dim ny As Double
                    If n1x <= n2x Then
                        nx = n1x
                        ny = n1y
                    Else
                        nx = n2x
                        ny = n2y
                    End If

                    Return New Point3d(midpoint.X + (textSize * nx), midpoint.Y + (textSize * ny), midpoint.Z)
                End If
            End If

            Dim offsetAngle As Double = If(rawAngle > (Math.PI * 0.5) AndAlso rawAngle < (Math.PI * 1.5), rawAngle - 1.0, rawAngle + 1.0)
            Return New Point3d(midpoint.X + (textSize * Math.Cos(offsetAngle)), midpoint.Y + (textSize * Math.Sin(offsetAngle)), midpoint.Z)
        End Function

        Private Function CountOtherSegmentsNearPoint(point As Point3d, currentSegment As Tuple(Of Point3d, Point3d), segments As IEnumerable(Of Tuple(Of Point3d, Point3d)), tolerance As Double) As Integer
            Dim count As Integer = 0

            For Each segment As Tuple(Of Point3d, Point3d) In segments
                If SegmentEquals(segment, currentSegment, 0.001) Then
                    Continue For
                End If

                If DistancePointToSegment2d(point, segment.Item1, segment.Item2) <= tolerance Then
                    count += 1
                End If
            Next

            Return count
        End Function

        Private Function SegmentEquals(a As Tuple(Of Point3d, Point3d), b As Tuple(Of Point3d, Point3d), tol As Double) As Boolean
            Return (PointsEqual2d(a.Item1, b.Item1, tol) AndAlso PointsEqual2d(a.Item2, b.Item2, tol)) OrElse
                   (PointsEqual2d(a.Item1, b.Item2, tol) AndAlso PointsEqual2d(a.Item2, b.Item1, tol))
        End Function

        Private Function PointsEqual2d(a As Point3d, b As Point3d, tol As Double) As Boolean
            Return Math.Abs(a.X - b.X) <= tol AndAlso Math.Abs(a.Y - b.Y) <= tol
        End Function

        Private Function DistancePointToSegment2d(p As Point3d, a As Point3d, b As Point3d) As Double
            Dim dx As Double = b.X - a.X
            Dim dy As Double = b.Y - a.Y
            Dim lengthSquared As Double = (dx * dx) + (dy * dy)

            If lengthSquared <= 0.000000001 Then
                Dim ddx As Double = p.X - a.X
                Dim ddy As Double = p.Y - a.Y
                Return Math.Sqrt((ddx * ddx) + (ddy * ddy))
            End If

            Dim t As Double = (((p.X - a.X) * dx) + ((p.Y - a.Y) * dy)) / lengthSquared
            If t < 0.0 Then
                t = 0.0
            ElseIf t > 1.0 Then
                t = 1.0
            End If

            Dim projX As Double = a.X + (t * dx)
            Dim projY As Double = a.Y + (t * dy)
            Dim distX As Double = p.X - projX
            Dim distY As Double = p.Y - projY
            Return Math.Sqrt((distX * distX) + (distY * distY))
        End Function

        Private Function ApplyJoistCount(baseLabel As String, joistCount As Integer) As String
            If joistCount > 1 Then
                Return "#" & joistCount.ToString() & " " & baseLabel
            End If

            Return baseLabel
        End Function

        Private Function IsolateSelectionLayer(tr As Transaction, db As Database, keepLayerName As String) As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))
            Return IsolateSelectionLayers(tr, db, keepLayerName)
        End Function

        Private Function IsolateSelectionLayers(tr As Transaction, db As Database, ParamArray keepLayerNames As String()) As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))
            Dim states As New Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean))()
            Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            Dim keepLayerNamesSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each layerName As String In keepLayerNames
                keepLayerNamesSet.Add(layerName)
            Next

            For Each layerId As ObjectId In lt
                Dim ltr As LayerTableRecord = CType(tr.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)
                states(layerId) = Tuple.Create(ltr.IsOff, ltr.IsFrozen)

                If keepLayerNamesSet.Contains(ltr.Name) Then
                    ltr.IsOff = False
                    ltr.IsFrozen = False
                Else
                    ltr.IsOff = True
                End If
            Next

            Return states
        End Function

        Private Sub RestoreLayerStates(tr As Transaction, states As Dictionary(Of ObjectId, Tuple(Of Boolean, Boolean)))
            If states Is Nothing Then
                Return
            End If

            For Each kvp As KeyValuePair(Of ObjectId, Tuple(Of Boolean, Boolean)) In states
                If kvp.Key.IsNull OrElse kvp.Key.IsErased Then
                    Continue For
                End If

                Dim ltr As LayerTableRecord = TryCast(tr.GetObject(kvp.Key, OpenMode.ForWrite), LayerTableRecord)
                If ltr Is Nothing Then
                    Continue For
                End If

                ltr.IsOff = kvp.Value.Item1
                ltr.IsFrozen = kvp.Value.Item2
            Next
        End Sub

        Private Sub EnsureLayerExistsAndOn(tr As Transaction, db As Database, layerName As String)
            Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

            If Not lt.Has(layerName) Then
                lt.UpgradeOpen()
                Dim ltr As New LayerTableRecord()
                ltr.Name = layerName
                lt.Add(ltr)
                tr.AddNewlyCreatedDBObject(ltr, True)
                Return
            End If

            Dim layerId As ObjectId = lt(layerName)
            Dim layer As LayerTableRecord = CType(tr.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)
            If layer.IsOff Then
                layer.IsOff = False
            End If
            If layer.IsFrozen Then
                layer.IsFrozen = False
            End If
        End Sub

        Private Sub EnsureLayerFromTemplateAndColor(tr As Transaction, db As Database, layerName As String, templateLayerName As String, colorIndex As Short)
            EnsureLayerExistsAndOn(tr, db, templateLayerName)

            Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            Dim templateId As ObjectId = lt(templateLayerName)
            Dim templateLayer As LayerTableRecord = CType(tr.GetObject(templateId, OpenMode.ForRead), LayerTableRecord)

            If Not lt.Has(layerName) Then
                lt.UpgradeOpen()
                Dim newLayer As New LayerTableRecord()
                newLayer.Name = layerName
                newLayer.LinetypeObjectId = templateLayer.LinetypeObjectId
                newLayer.LineWeight = templateLayer.LineWeight
                newLayer.PlotStyleNameId = templateLayer.PlotStyleNameId
                newLayer.IsPlottable = templateLayer.IsPlottable
                newLayer.Description = templateLayer.Description
                newLayer.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, colorIndex)

                lt.Add(newLayer)
                tr.AddNewlyCreatedDBObject(newLayer, True)
            End If

            Dim layerId As ObjectId = lt(layerName)
            Dim layer As LayerTableRecord = CType(tr.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)
            layer.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, colorIndex)
            If layer.IsOff Then
                layer.IsOff = False
            End If
            If layer.IsFrozen Then
                layer.IsFrozen = False
            End If
        End Sub

        Private Sub MoveEntitiesToLayer(ids As IEnumerable(Of ObjectId), tr As Transaction, targetLayer As String)
            For Each id As ObjectId In ids
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForWrite), Entity)
                If ent Is Nothing Then
                    Continue For
                End If

                ent.Layer = targetLayer
            Next
        End Sub

    End Class
End Namespace