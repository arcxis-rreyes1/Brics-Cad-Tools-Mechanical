Imports System
Imports System.Globalization
Imports System.Windows.Forms
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.Colors
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.WsPasteCc))>
Namespace Arcxis_Cad_Tools
    Public Class WsPasteCc

        <CommandMethod("DD", CommandFlags.Modal)>
        Public Sub WS()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return

            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database

            Dim clipboardText As String = GetClipboardText()
            If String.IsNullOrWhiteSpace(clipboardText) Then
                ed.WriteMessage(vbLf & "Clipboard is empty or contains no text.")
                Return
            End If

            Dim lines As String() = clipboardText.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
            If lines.Length = 0 Then
                ed.WriteMessage(vbLf & "No text lines found in clipboard.")
                Return
            End If

            Dim ppr = ed.GetPoint(vbLf & "Pick top left of empty box:")
            If ppr.Status <> PromptStatus.OK Then Return

            Dim textHeight As Double = 6.0
            Dim lineStep As Double = 8.0
            Dim startPt As Point3d = ppr.Value
            Dim textStart As New Point3d(startPt.X + 2.0, startPt.Y - 6.0, startPt.Z)

            Try
                Using doc.LockDocument()
                    Using tr = db.TransactionManager.StartTransaction()
                        EnsureLayer(db, tr, "S-FRM-LOAD-AREA-NOPLT", 3, "Continuous", False)
                        'Dim txtStyleId As ObjectId = EnsureTextStyle(db, tr, "arcxis_std", "simplex.shx", 0.0, 0.8, 0.0)
                        Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                        Dim rowIndex As Integer = 0
                        For Each line As String In lines
                            Dim rowText As String = line.Replace(vbTab, " ").Trim()
                            If rowText.Length = 0 Then
                                rowIndex += 1
                                Continue For
                            End If

                            Dim mt As New MText()
                            mt.SetDatabaseDefaults()
                            mt.Location = New Point3d(textStart.X, textStart.Y - (rowIndex * lineStep), textStart.Z)
                            mt.Contents = rowText
                            mt.TextHeight = textHeight
                            mt.Layer = "S-FRM-LOAD-AREA-NOPLT"
                            mt.ColorIndex = 7
                            'mt.TextStyleId = txtStyleId
                            btr.AppendEntity(mt)
                            tr.AddNewlyCreatedDBObject(mt, True)

                            rowIndex += 1
                        Next

                        tr.Commit()
                    End Using
                End Using

                ed.WriteMessage(vbLf & $"DD complete. Placed {lines.Length} row(s) from clipboard.")

            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "DD failed: " & ex.Message)
            End Try
        End Sub

        <CommandMethod("OSC", CommandFlags.Modal)>
        Public Sub ClipboardPasteAsOle()

            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor

            ' 1) Get basepoint
            Dim ppr = ed.GetPoint(vbLf & "Get Basepoint: ")
            If ppr.Status <> PromptStatus.OK Then Return

            Dim pt = ppr.Value
            Dim ptStr As String =
            pt.X.ToString(CultureInfo.InvariantCulture) & "," &
            pt.Y.ToString(CultureInfo.InvariantCulture) & "," &
            pt.Z.ToString(CultureInfo.InvariantCulture)

            ' 2) PASTECLIP at point
            ' 3) SCALE previous selection by 40 about the same point
            '
            ' SCALE command sequence:
            '   SCALE <select objects>  (we use Previous = P)
            '   <enter> to finish selection
            '   <basepoint>
            '   <scale factor>
            '
            Dim cmd As String =
            "_.PASTECLIP " & ptStr & vbLf &
            "_.SCALE P " & vbLf &
            ptStr & vbLf &
            "40" & vbLf

            doc.SendStringToExecute(cmd, True, False, False)

        End Sub



        Private Function GetClipboardText() As String
            Try
                If Clipboard.ContainsText() Then
                    Return Clipboard.GetText()
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function GetCurrentLayerName(db As Database) As String
            If db Is Nothing Then Return "0"

            Using tr = db.TransactionManager.StartTransaction()
                Dim ltr = TryCast(tr.GetObject(db.Clayer, OpenMode.ForRead), LayerTableRecord)
                If ltr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ltr.Name) Then
                    Return ltr.Name
                End If
                tr.Commit()
            End Using

            Return "0"
        End Function

        Private Function FormatPoint(pt As Point3d) As String
            Return pt.X.ToString("0.######", Globalization.CultureInfo.InvariantCulture) & "," &
                   pt.Y.ToString("0.######", Globalization.CultureInfo.InvariantCulture) & "," &
                   pt.Z.ToString("0.######", Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Function QuoteForCommand(value As String) As String
            Return """" & value.Replace("""", "") & """"
        End Function

        Private Function EnsureTextStyle(db As Database,
                                 tr As Transaction,
                                 styleName As String,
                                 Optional fontFileName As String = "simplex.shx",
                                 Optional textHeight As Double = 0.0,
                                 Optional widthFactor As Double = 0.8,
                                 Optional obliqueAngleDegrees As Double = 0.0) As ObjectId
            Dim tst = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)

            If tst.Has(styleName) Then
                Return tst(styleName)
            End If

            tst.UpgradeOpen()

            Dim ts As New TextStyleTableRecord() With {
        .Name = styleName,
        .FileName = fontFileName,
        .TextSize = textHeight,
        .XScale = widthFactor,
        .ObliquingAngle = obliqueAngleDegrees * (Math.PI / 180.0) ' degrees -> radians
    }

            Dim id = tst.Add(ts)
            tr.AddNewlyCreatedDBObject(ts, True)
            Return id
        End Function
        Private Sub EnsureLayer(db As Database,
                                tr As Transaction,
                                layerName As String,
                                colorIndex As Integer,
                                linetypeName As String,
                                isPlottable As Boolean)
            Dim lt = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            Dim linetypeId As ObjectId = EnsureLineTypeRecursive(db, tr, linetypeName)

            Dim ltr As LayerTableRecord = Nothing
            If lt.Has(layerName) Then
                ltr = CType(tr.GetObject(lt(layerName), OpenMode.ForWrite), LayerTableRecord)
            Else
                lt.UpgradeOpen()
                ltr = New LayerTableRecord()
                ltr.Name = layerName
                lt.Add(ltr)
                tr.AddNewlyCreatedDBObject(ltr, True)
            End If

            If colorIndex > 0 AndAlso colorIndex <= 255 Then
                ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, CShort(colorIndex))
            End If

            If Not linetypeId.IsNull Then
                ltr.LinetypeObjectId = linetypeId
            End If

            ltr.IsPlottable = isPlottable
        End Sub

        Private Function EnsureLineTypeRecursive(db As Database,
                                                 tr As Transaction,
                                                 linetypeName As String,
                                                 Optional attempt As Integer = 0) As ObjectId
            Dim ltt = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
            If ltt.Has(linetypeName) Then Return ltt(linetypeName)

            Dim candidates As String() = {"acad.lin", "default.lin"}
            If attempt >= candidates.Length Then Return ObjectId.Null

            Try
                ltt.UpgradeOpen()
                db.LoadLineTypeFile(linetypeName, candidates(attempt))
            Catch
            End Try

            Return EnsureLineTypeRecursive(db, tr, linetypeName, attempt + 1)
        End Function

    End Class
End Namespace
