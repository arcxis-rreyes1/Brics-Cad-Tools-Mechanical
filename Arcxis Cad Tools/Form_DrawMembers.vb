Imports System.Linq
Imports System.Windows
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class Form_DrawMembers
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.TextChanged
        Dim sel As String = TryCast(ComboBox3.Text, String)
        If sel Is Nothing Then
            ComboBox4.Items.Clear()
            Exit Sub
        End If
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        Select Case sel.ToUpperInvariant()
            Case "2X"
                For Each s In New String() {"4", "6", "8", "10", "12"} : ComboBox4.Items.Add(s) : Next
            Case "1.75", "3.5", "5.25", "7"
                For Each s In New String() {"5.5", "7.25", "9.5", "11.25", "11.875", "14", "16", "18", "24"} : ComboBox4.Items.Add(s) : Next
            Case "RIM", "BLK"
                For Each s In New String() {"4", "6", "8", "10", "12", "9.5", "11.25", "11.875", "14", "16", "18"} : ComboBox4.Items.Add(s) : Next
            Case "3 1/2", "4"
                For Each s In New String() {"1/2", "5/16", "3/8", "1/4"} : ComboBox4.Items.Add(s) : Next
        End Select
    End Sub


    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.TextChanged
        Dim sel As String = TryCast(ComboBox6.SelectedItem, String)
        If sel Is Nothing Then sel = String.Empty
        Dim isSteel As Boolean = String.Equals(sel, "STEEL", StringComparison.OrdinalIgnoreCase)
        GroupBox12.Visible = isSteel
        ComboBox9.Visible = isSteel
        ComboBox3.Enabled = True
        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        Select Case sel.ToUpperInvariant()
            Case "PSL 2.0E", "LVL 2.0E", "LSL 1.55E", "VERSA-LAM 1.7E", "VERSA-LAM 2.0E"
                ComboBox3.Items.Add("1.75")
                ComboBox3.SelectedIndex = 0
                setplyto1(False)
            Case "GROUP 1", "GROUP 2", "GROUP 3", "GROUP 4"
                ComboBox3.Enabled = False
                ComboBox3.Items.Clear()
                For Each s In New String() {"9.5", "11.875", "14", "16", "18"} : ComboBox4.Items.Add(s) : Next
                setplyto1(False)
            Case "BLK"
                ComboBox3.Items.Add("Blk")
                ComboBox3.SelectedIndex = 0
                ComboBox3.Enabled = False
                ComboBox6.Items.Clear()
                ComboBox6.Items.Add("Blk")
                ComboBox6.SelectedIndex = 0
                ComboBox6.Enabled = False
                setplyto1(True)
            Case "RIM"
                ComboBox3.Items.Add("Rim")
                ComboBox3.SelectedIndex = 0
                ComboBox3.Enabled = False
                ComboBox6.Items.Add("Rim")
                ComboBox6.SelectedIndex = 0
                ComboBox6.Enabled = False
                setplyto1(True)
            Case "3 1/2", "4", "5", "6", "7", "8"
                For Each s In New String() {"3 1/2", "4"} : ComboBox3.Items.Add(s) : Next
                setplyto1(True)
            Case Else
                setplyto1(False)
        End Select



        Select Case sel.ToUpperInvariant()
            Case "BLK", "RIM", "3 1/2", "4", "5", "6", "7", "8"
                ComboBox7.Enabled = False : ComboBox8.Enabled = False
            Case Else
                ComboBox7.Enabled = True : ComboBox8.Enabled = True
        End Select
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Dim sel As String = TryCast(ComboBox2.SelectedItem, String)
        If sel Is Nothing Then sel = String.Empty
        Dim isSteel As Boolean = String.Equals(sel, "STEEL", StringComparison.OrdinalIgnoreCase)
        GroupBox12.Visible = isSteel
        ComboBox9.Visible = isSteel
        ComboBox3.Enabled = True
        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        ComboBox6.Enabled = False
        ComboBox6.Text = ""
        Select Case sel.ToUpperInvariant()
            Case "I-JOIST"
                ComboBox6.Items.Clear()
                For Each s In New String() {"GROUP 1", "GROUP 2", "GROUP 3", "GROUP 4"} : ComboBox6.Items.Add(s) : Next
                ComboBox6.Enabled = True
                ComboBox3.Enabled = False
                ComboBox3.Text = ""
                setplyto1(False)
            Case "COMMODITY"
                setplyto1(False)
                For Each s In New String() {"#2 SYP"} : ComboBox6.Items.Add(s) : Next
                ComboBox6.Enabled = True
                ComboBox3.Items.Clear()
                ComboBox3.Enabled = False
            Case "FLUSH BEAM", "DROP BEAM", "HEADER"
                ComboBox6.Enabled = True
                setplyto1(False)
            Case "LINTEL"
                ComboBox6.Items.Clear()
                For Each s In New String() {"3 1/2", "4", "5", "6", "7", "8"} : ComboBox6.Items.Add(s) : Next
                ComboBox6.Enabled = True
                GroupBox8.Text = "Height"
                setplyto1(False)
            Case "BLK"
                ComboBox3.Items.Clear()
                ComboBox3.Enabled = False
                ComboBox6.Items.Add("Blk")
                ComboBox6.SelectedIndex = 0
                ComboBox6.Enabled = False
                setplyto1(True)
            Case "RIM"
                ComboBox3.Items.Clear()
                ComboBox3.Enabled = False
                ComboBox6.Items.Add("Rim")
                ComboBox6.SelectedIndex = 0
                ComboBox6.Enabled = False
                setplyto1(True)
            Case Else
                setplyto1(False)
        End Select

        If sel.ToUpperInvariant() <> "LINTEL" Then
            GroupBox8.Text = "Product Type"
        End If

        Select Case sel.ToUpperInvariant()
            Case "BLK", "RIM"
                ComboBox7.Enabled = False : ComboBox8.Enabled = False
            Case Else
                ComboBox7.Enabled = True : ComboBox8.Enabled = True
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim width As String = TryCast(ComboBox3.SelectedItem, String)
        Dim member As String = TryCast(ComboBox2.SelectedItem, String)
        Dim materialType As String = TryCast(ComboBox6.SelectedItem, String) ' material type supplied from form (adjust control if different)
        Dim MemberLayer As String = ""
        Dim MemberLayerColor As Integer = 0
        Dim MemberLineScale As Integer = 0
        Dim MemberLineType As String = ""
        Dim MemberThickness As Double
        Dim stringsuffix As String = ComboBox8.SelectedItem
        Dim stringprefix As String = ComboBox7.SelectedItem
        Dim selText As String = If(ComboBox4.SelectedItem?.ToString(), String.Empty)
        Me.Hide()

        If member <> "Lintel" Then
            If Not Double.TryParse(selText, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, MemberThickness) Then
                MemberThickness = 0R
            End If
        Else
            ' For Lintel, parse the width as a fraction (e.g., "3 1/2" -> 3.5)
            MemberThickness = ParseFraction(width)
        End If
        Dim LocationPrefix As String = If(ComboBox1.SelectedItem = "Ceiling", "S-FRM-CLG-", "S-FRM-FL-")
        Select Case member.ToUpperInvariant()
            Case "2X"
                MemberLayer = "S-FRM-2X" : MemberLayerColor = 12 : MemberLineScale = 1
                MemberLineType = If(member = "Header", "hidden2", "ByLayer")
            Case "I-JOIST"
                Select Case materialType.ToUpperInvariant()
                    Case "GROUP1"
                        MemberLayer = "S-FRM-FL-GROUP-1" : MemberLayerColor = 80 : MemberLineScale = 1 : MemberLineType = "ByLayer"
                    Case "GROUP2"
                        MemberLayer = "S-FRM-FL-GROUP-2" : MemberLayerColor = 20 : MemberLineScale = 1 : MemberLineType = "ByLayer"
                    Case "GROUP3"
                        MemberLayer = "S-FRM-FL-GROUP-3" : MemberLayerColor = 6 : MemberLineScale = 1 : MemberLineType = "ByLayer"
                    Case "GROUP4"
                        MemberLayer = "S-FRM-FL-GROUP-4" : MemberLayerColor = 4 : MemberLineScale = 1 : MemberLineType = "ByLayer"
                    Case Else
                        MessageBox.Show("Select a valid material.") : Me.Show() : Exit Sub
                End Select
        'Case "3.5"
        '    MemberLayer = "S-FRM-BM-3.5" : MemberLayerColor = If(member = "Column", 1, 221) : MemberLineScale = 1.75 : MemberLineType = "phantom2"
        'Case "5.25"
        '    MemberLayer = "S-FRM-BM-5.25" : MemberLayerColor = 231 : MemberLineScale = 2 : MemberLineType = "phantom2"
        'Case "7"
        '    MemberLayer = "S-FRM-BM-7" : MemberLayerColor = 241 : MemberLineScale = 2 : MemberLineType = "phantom2"
        'Case "2.5"
        '    MemberLayer = LocationPrefix & "2.5" : MemberLayerColor = 12 : MemberLineScale = 1 : MemberLineType = "ByLayer"
        'Case "1.5"
        '    MemberLayer = LocationPrefix & "1.5" : MemberLayerColor = 12 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "2"
                MemberLayer = "S-FRM-BBO" : MemberLayerColor = 12 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "BLK"
                MemberLayer = LocationPrefix & "BLK" : MemberLayerColor = 80 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "GROUP1"
                MemberLayer = "S-FRM-FL-GROUP-1" : MemberLayerColor = 80 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "GROUP2"
                MemberLayer = "S-FRM-FL-GROUP-2" : MemberLayerColor = 20 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "GROUP3"
                MemberLayer = "S-FRM-FL-GROUP-3" : MemberLayerColor = 6 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "GROUP4"
                MemberLayer = "S-FRM-FL-GROUP-4" : MemberLayerColor = 4 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "RIM"
                MemberLayer = "S-FRM-FL-RIM" : MemberLayerColor = 240 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case "LINTEL"
                MemberLayer = "S-FRM-BM-STEEL-1" : MemberLayerColor = 256 : MemberLineScale = 1 : MemberLineType = "ByLayer"
            Case Else
                MessageBox.Show("Select a valid material.") : Me.Show() : Exit Sub
        End Select


        Dim ply As Integer
        If Not Integer.TryParse(ComboBox5.Text, ply) Then ply = 1
        Try
            DrawMemberLinesForPly(MemberLayer, MemberLayerColor, MemberLineType, MemberLineScale, ply, width, member, MemberThickness, materialType, stringprefix, stringsuffix)
        Catch ex As Exception
            MessageBox.Show("Error drawing line: " & ex.Message)
        End Try
        Me.Show()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sel As String = TryCast(ComboBox1.SelectedItem, String)
        If sel Is Nothing Then sel = String.Empty
        ComboBox2.Items.Clear()
        Select Case sel.ToUpperInvariant()
            Case "CEILING"
                For Each s In New String() {"Flush Beam", "Drop Beam", "Header", "I-Joist", "Commodity", "Blk", "Steel", "Column", "Girder Truss", "Truss", "Lintel"}
                    ComboBox2.Items.Add(s)
                Next
            Case "FLOOR"
                For Each s In New String() {"Flush Beam", "Drop Beam", "Header", "I-Joist", "Commodity", "Rim", "Blk", "Steel", "Column", "Girder Truss", "Truss", "Lintel"}
                    ComboBox2.Items.Add(s)
                Next
        End Select
    End Sub

    Private Sub DrawMemberLinesForPly(layerName As String,
                                  colorIndex As Integer,
                                  lineTypeName As String,
                                  lineTypeScale As Double,
                                  ply As Integer,
                                  width As String,
                                  member As String,
                                  memberThickness As Double,
                                  materialType As String,
                                  stringprefix As String,
                                  stringsuffix As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Throw New InvalidOperationException("No active AutoCAD document.")
        Dim ed As Editor = doc.Editor

        Dim res1 = ed.GetPoint(New PromptPointOptions(vbLf & "Specify center start point:"))
        If res1.Status <> PromptStatus.OK Then Exit Sub
        Dim res2 = ed.GetPoint(New PromptPointOptions(vbLf & "Specify center end point:") With {.BasePoint = res1.Value, .UseBasePoint = True})
        If res2.Status <> PromptStatus.OK Then Exit Sub

        Dim startPt = res1.Value
        Dim endPt = res2.Value
        Dim dir = endPt - startPt
        If dir.Length < Tolerance.Global.EqualPoint Then
            ed.WriteMessage(vbLf & "Zero length; aborted.")
            Exit Sub
        End If

        ' Direction (centerline) and its LEFT perpendicular (CCW 90°)
        Dim dirUnit As Vector3d = dir.GetNormal()
        Dim perp As Vector3d = New Vector3d(-dirUnit.Y, dirUnit.X, 0) ' LEFT side

        Dim OffsetPrimary As Double
        Dim OffsetSecondaryDelta As Double

        If width = "2x" Then
            OffsetPrimary = 0.875
            OffsetSecondaryDelta = 1.75
        Else
            OffsetPrimary = 1.0
            OffsetSecondaryDelta = 2.0
        End If

        Dim offsets As New List(Of Double)
        Select Case ply
            Case 1 : offsets.Add(0)
            Case 2 : offsets.AddRange({OffsetPrimary, -OffsetPrimary})
            Case 3 : offsets.AddRange({0, OffsetPrimary, -OffsetPrimary})
            Case 4 : offsets.AddRange({OffsetPrimary, -OffsetPrimary, OffsetPrimary + OffsetSecondaryDelta, -(OffsetPrimary + OffsetSecondaryDelta)})
            Case Else : offsets.Add(0)
        End Select

        Dim wKey = If(width, "").ToUpperInvariant()
        Dim extension As Double = If(wKey = "BLK" OrElse wKey = "RIM" OrElse wKey = "LINTEL", 0.0,
                                 If(String.Equals(member, "Header", StringComparison.OrdinalIgnoreCase), 1.5, 3.5))

        Using doc.LockDocument()
            Dim db = doc.Database
            Using tr = db.TransactionManager.StartTransaction()
                EnsureLayer(db, tr, layerName, colorIndex, lineTypeName)
                If Not String.Equals(lineTypeName, "BYLAYER", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(lineTypeName) Then
                    EnsureLineType(db, tr, lineTypeName)
                End If
                Dim btr = CType(tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite), BlockTableRecord)

                Dim centerExtendedStart As Point3d = startPt - dirUnit * extension
                Dim centerExtendedEnd As Point3d = endPt + dirUnit * extension

                ' Only create polylines if NOT a lintel
                If Not String.Equals(member, "LINTEL", StringComparison.OrdinalIgnoreCase) Then
                    For Each d In offsets
                        Dim baseP1 = startPt + perp * d
                        Dim baseP2 = endPt + perp * d
                        Dim p1 = baseP1 - dirUnit * extension
                        Dim p2 = baseP2 + dirUnit * extension

                        Dim pl As New Polyline()
                        pl.SetDatabaseDefaults()
                        pl.AddVertexAt(0, New Point2d(p1.X, p1.Y), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(p2.X, p2.Y), 0, 0, 0)
                        pl.Layer = layerName
                        If memberThickness > 0 Then pl.Thickness = memberThickness
                        If Not String.Equals(lineTypeName, "BYLAYER", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(lineTypeName) Then
                            pl.Linetype = lineTypeName
                            pl.LinetypeScale = lineTypeScale
                        End If
                        btr.AppendEntity(pl)
                        tr.AddNewlyCreatedDBObject(pl, True)
                    Next
                End If

                Dim DimensionText As String
                If member = "BLK" Or width = "RIM" Then
                    DimensionText = "BLK"
                ElseIf member = "Lintel" Then
                    ' Format as 2 lines for lintel: height on first line, description on second
                    ' Convert thickness back to fraction string for display
                    Dim thicknessAsString As String = DecimalToFraction(memberThickness)
                    DimensionText = "L" & materialType & "x" & width & "x" & thicknessAsString & vbLf & "STEEL LINTEL"
                End If
                ' When BLK is selected, add an aligned dimension offset 6" from the line with text "BLK" and dimstyle "jd64a"
                If String.Equals(member, "BLK", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(width, "BLK", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(member, "LINTEL", StringComparison.OrdinalIgnoreCase) Then
                    Dim dimStyleId As ObjectId = EnsureDimStyle(db, tr, "ARCXIS-FRM-DIM")

                    ' Midpoint of picked line
                    Dim mid As Point3d = New Point3d((startPt.X + endPt.X) / 2.0, (startPt.Y + endPt.Y) / 2.0, (startPt.Z + endPt.Z) / 2.0)

                    ' Decide offset direction independent of draw direction
                    Dim rawAngle As Double = dir.AngleOnPlane(New Plane(Point3d.Origin, Vector3d.ZAxis))
                    rawAngle = rawAngle Mod (2.0 * Math.PI) : If rawAngle < 0 Then rawAngle += 2.0 * Math.PI

                    Const axisTol As Double = 0.01 ' ~0.57°
                    Dim isHorizontal As Boolean = (Math.Abs(Math.Sin(rawAngle)) < axisTol) Or (Math.Abs(Math.Sin(rawAngle - Math.PI)) < axisTol)
                    Dim isVertical As Boolean = (Math.Abs(Math.Cos(rawAngle)) < axisTol) Or (Math.Abs(Math.Cos(rawAngle - Math.PI)) < axisTol)

                    Dim offsetVec As Vector3d
                    If isVertical Then
                        ' Always to LEFT of vertical (negative X)
                        offsetVec = New Vector3d(-1, 0, 0) * 6.0
                    ElseIf isHorizontal Then
                        ' Always ABOVE horizontal (positive Y)
                        offsetVec = New Vector3d(0, 1, 0) * 6.0
                    Else
                        ' For angled lines, use mathematical left but force ABOVE (positive Y)
                        Dim perpLeft As Vector3d = New Vector3d(-dirUnit.Y, dirUnit.X, 0).GetNormal()
                        If perpLeft.Y < 0 Then perpLeft = perpLeft.Negate()
                        offsetVec = perpLeft * 6.0
                    End If

                    Dim dimLinePoint As Point3d = mid + offsetVec

                    Dim ad As New AlignedDimension(startPt, endPt, dimLinePoint, DimensionText, dimStyleId)
                    ad.Layer = "S-ANNO-TEXT" ' or layerName if preferred
                    btr.AppendEntity(ad)
                    tr.AddNewlyCreatedDBObject(ad, True)
                End If


                If Not String.Equals(member, "LINTEL", StringComparison.OrdinalIgnoreCase) Then
                    ' Label build (unchanged)
                    Dim label As String = BuildMemberLabel(member, width, memberThickness, ply, centerExtendedStart.DistanceTo(centerExtendedEnd), materialType, stringprefix, stringsuffix)
                    If Not String.IsNullOrWhiteSpace(label) Then
                        Dim txtStyleId As ObjectId = EnsureTextStyle(db, tr, "arcxis_std", "simplex.shx", 0.0, 0.8, 0.0)
                        Dim mid As Point3d = New Point3d((centerExtendedStart.X + centerExtendedEnd.X) / 2.0,
                                                 (centerExtendedStart.Y + centerExtendedEnd.Y) / 2.0,
                                                 (centerExtendedStart.Z + centerExtendedEnd.Z) / 2.0)
                        Dim textOffsetDistance As Double = 6.0 + (ply - 1)
                        Dim rawAngle As Double = dir.AngleOnPlane(New Plane(Point3d.Origin, Vector3d.ZAxis))
                        rawAngle = rawAngle Mod (2.0 * Math.PI) : If rawAngle < 0 Then rawAngle += 2.0 * Math.PI
                        Const axisTol As Double = 0.01
                        Dim isHorizontal As Boolean = (Math.Abs(Math.Sin(rawAngle)) < axisTol) Or (Math.Abs(Math.Sin(rawAngle - Math.PI)) < axisTol)
                        Dim isVertical As Boolean = (Math.Abs(Math.Cos(rawAngle)) < axisTol) Or (Math.Abs(Math.Cos(rawAngle - Math.PI)) < axisTol)
                        Dim leftVec As Vector3d = New Vector3d(-1, 0, 0)
                        Dim aboveVec As Vector3d = New Vector3d(0, 1, 0)
                        Dim perpLeft As Vector3d = New Vector3d(-dirUnit.Y, dirUnit.X, 0).GetNormal()
                        Dim finalOffsetVec As Vector3d
                        If isVertical Then
                            finalOffsetVec = leftVec * textOffsetDistance
                        ElseIf isHorizontal Then
                            finalOffsetVec = aboveVec * textOffsetDistance
                        Else
                            If perpLeft.X > 0 Then perpLeft = perpLeft.Negate()
                            finalOffsetVec = perpLeft * textOffsetDistance
                        End If
                        Dim textPoint As Point3d = mid + finalOffsetVec
                        Dim readableAngle As Double = NormalizeTextRotation(rawAngle)

                        Dim mt As New MText()
                        mt.SetDatabaseDefaults()
                        mt.Location = textPoint
                        mt.Attachment = AttachmentPoint.MiddleCenter
                        mt.Contents = label
                        mt.Layer = "S-ANNO-TEXT"
                        mt.TextStyleId = txtStyleId
                        mt.Rotation = readableAngle
                        btr.AppendEntity(mt)
                        tr.AddNewlyCreatedDBObject(mt, True)
                    End If
                End If

                tr.Commit()
            End Using
        End Using
    End Sub

    Private Function BuildMemberLabel(member As String,
                                      widthKey As String,
                                      memberThickness As Double,
                                      ply As Integer,
                                      fullLength As Double,
                                      materialType As String,
                                      Stringprefix As String,
                                      StringSuffix As String) As String

        If String.IsNullOrWhiteSpace(member) Then Return String.Empty

        Dim plyTag As String = ""
        If Stringprefix <> "" Then Stringprefix = Stringprefix & " "
        If StringSuffix <> "" Then StringSuffix = " " & StringSuffix
        Dim upperMember = member.ToUpperInvariant()
        Dim feet As Integer
        If upperMember = "I-JOIST" Then

            plyTag = If(ply = 2, "DLB ", If(ply = 3, "TRPL ", If(ply = 4, "QUAD ", "")))
            ' length in inches -> round up to nearest foot
            feet = CInt(Math.Ceiling(fullLength / 12.0))
            Return $"{plyTag}{widthKey}-{feet}'"
        Else
            Dim thicknessText As String = If(memberThickness > 0, memberThickness.ToString("0.###", Globalization.CultureInfo.InvariantCulture), "?")
            Dim mat As String = If(String.IsNullOrWhiteSpace(materialType), "", materialType)
            feet = CInt(Math.Ceiling(fullLength / 12.0))
            ' Only 2x beams carry a length in the label here
            If String.Equals(widthKey, "2x", StringComparison.OrdinalIgnoreCase) Then
                If feet > 8 Then
                    ' Round up to next multiple of 2 (even number)
                    feet = CInt(Math.Ceiling(feet / 2.0)) * 2
                    ' Alternatively: If (feet Mod 2) <> 0 Then feet += 1
                End If
            End If
            If ply > 1 Then
                If upperMember = "COMMODITY" Then
                    Return $"{Stringprefix}{ply}~2x{thicknessText}-{feet}' {mat}{StringSuffix}".Trim()
                Else
                    Return $"{Stringprefix}{ply}~{widthKey}x{thicknessText}-{feet}' {mat}{StringSuffix}".Trim()
                End If

            Else
                Return $"{widthKey}x{thicknessText}-{feet}' {mat}".Trim()
            End If
        End If
        ' Other members not labeled for now
        Return String.Empty
    End Function

    ' Ensures a dimension style exists (minimal setup). If missing, creates it.
    Private Function EnsureDimStyle(db As Database, tr As Transaction, styleName As String) As ObjectId
        Dim dst = CType(tr.GetObject(db.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
        If dst.Has(styleName) Then Return dst(styleName)

        dst.UpgradeOpen()
        Dim ds As New DimStyleTableRecord() With {
            .Name = styleName,
            .Dimtxt = 0.125,    ' Text height = 1/8"
            .Dimscale = 64.0     ' Overall dimension scale
        }
        Dim id = dst.Add(ds)
        tr.AddNewlyCreatedDBObject(ds, True)
        Return id
    End Function
    Private Function EnsureLayer(db As Database,
                                 tr As Transaction,
                                 layerName As String,
                                 colorIndex As Integer,
                                 lineTypeName As String) As ObjectId
        Dim lt = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
        If lt.Has(layerName) Then Return lt(layerName)
        lt.UpgradeOpen()
        Dim newLayer As New LayerTableRecord() With {
            .Name = layerName,
            .Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, CShort(Math.Max(1, Math.Min(255, colorIndex))))
        }
        If Not String.Equals(lineTypeName, "BYLAYER", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(lineTypeName) Then
            Dim ltt = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
            If ltt.Has(lineTypeName) Then newLayer.LinetypeObjectId = ltt(lineTypeName)
        End If
        Dim id = lt.Add(newLayer)
        tr.AddNewlyCreatedDBObject(newLayer, True)
        Return id
    End Function

    Private Sub EnsureLineType(db As Database, tr As Transaction, lineTypeName As String)
        Dim ltt = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
        If ltt.Has(lineTypeName) Then Return
        ltt.UpgradeOpen()
        db.LoadLineTypeFile(lineTypeName, "acad.lin")
    End Sub

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
    Sub setplyto1(Remove As Boolean)
        If Remove Then
            ComboBox5.Enabled = False
            ComboBox5.SelectedIndex = 0
            ComboBox5.Text = 1
        Else
            ComboBox5.Enabled = True
            ComboBox5.Text = ""
        End If
    End Sub
    Private Function NormalizeTextRotation(angle As Double) As Double
        ' Keep angle within 0..2π
        angle = angle Mod (2.0 * Math.PI)
        If angle < 0 Then angle += 2.0 * Math.PI
        ' Flip if upside-down; include 270° by using <= on the upper bound
        If angle > Math.PI / 2.0 AndAlso angle <= 3.0 * Math.PI / 2.0 Then
            angle -= Math.PI
        End If
        Return angle
    End Function

    Private Function ParseFraction(input As String) As Double
        ' Converts "3 1/2" to 3.5, "4" to 4.0, "1/2" to 0.5, etc.
        If String.IsNullOrWhiteSpace(input) Then Return 0.0

        input = input.Trim()
        Dim result As Double = 0.0

        ' Split by space to separate whole number from fraction
        Dim parts = input.Split(" "c)

        ' Parse whole number if present
        Dim numerator As Double
        Dim denominator As Double

        If parts.Length > 0 AndAlso Double.TryParse(parts(0), result) Then
            ' First part is a whole number
            If parts.Length > 1 Then
                ' There's a fraction part too
                Dim fractionParts = parts(1).Split("/"c)
                If fractionParts.Length = 2 Then
                    If Double.TryParse(fractionParts(0), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, numerator) AndAlso
                       Double.TryParse(fractionParts(1), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, denominator) Then
                        If denominator <> 0 Then
                            result += numerator / denominator
                        End If
                    End If
                End If
            End If
        Else
            ' Try parsing as just a fraction (e.g., "1/2")
            Dim fractionParts = input.Split("/"c)
            If fractionParts.Length = 2 Then
                If Double.TryParse(fractionParts(0), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, numerator) AndAlso
                   Double.TryParse(fractionParts(1), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, denominator) Then
                    If denominator <> 0 Then
                        result = numerator / denominator
                    End If
                End If
            End If
        End If

        Return result
    End Function

    Private Function DecimalToFraction(value As Double) As String
        ' Converts 3.5 to "3 1/2", 4.0 to "4", 0.5 to "1/2", etc.
        If value = 0.0 Then Return "0"

        ' Common fractions to check (denominator 16 for common building measurements)
        Dim whole As Integer = CInt(Math.Floor(value))
        Dim remainder As Double = value - whole

        ' Check for common fractions with tolerance
        Const tolerance As Double = 0.01

        ' Check common denominators: 2, 4, 8, 16
        For denominator As Integer = 2 To 16
            For numerator As Integer = 1 To denominator - 1
                Dim fractionValue As Double = numerator / CDbl(denominator)
                If Math.Abs(remainder - fractionValue) < tolerance Then
                    If whole > 0 Then
                        Return $"{whole} {numerator}/{denominator}"
                    Else
                        Return $"{numerator}/{denominator}"
                    End If
                End If
            Next
        Next

        ' If no common fraction found, return the decimal value
        If whole > 0 Then
            Return $"{whole} {remainder:F4}".TrimEnd("0"c).TrimEnd("."c)
        Else
            Return value.ToString("F4").TrimEnd("0"c).TrimEnd("."c)
        End If
    End Function

    Private Function GCD(a As Integer, b As Integer) As Integer
        While b <> 0
            Dim temp As Integer = b
            b = a Mod b
            a = temp
        End While
        Return a
    End Function
End Class