Imports System
Imports System.Reflection
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Exception = Teigha.Runtime.Exception

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.MechanicalFunctions))>

Namespace Arcxis_Cad_Tools
    Public Class MechanicalFunctions
        <CommandMethod("EXCEL2TABLE")>
        Public Sub ExcelSelectionToAcadTable()

            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            ' --- Get Excel selection ---
            Dim xlApp As Excel.Application = Nothing
            Try
                ' Use VB Interaction.GetObject to bind to the running Excel instance (COM ROT)
                xlApp = CType(Interaction.GetObject("", "Excel.Application"), Excel.Application)
            Catch ex As COMException
                ed.WriteMessage(vbLf & "Excel is not running or no selection found.")
                Return
            Catch
                ed.WriteMessage(vbLf & "Excel is not running or no selection found.")
                Return
            End Try

            Dim rng As Excel.Range = TryCast(xlApp.Selection, Excel.Range)
            If rng Is Nothing Then
                ed.WriteMessage(vbLf & "Please select a cell range in Excel before running EXCEL2TABLE.")
                Return
            End If

            ' --------------------------------------------------------------------
            '   Build lists of visible rows/columns
            ' --------------------------------------------------------------------
            Dim visibleRows As New List(Of Excel.Range)
            Dim visibleCols As New List(Of Excel.Range)

            For r As Integer = 1 To rng.Rows.Count
                Dim row As Excel.Range = rng.Rows(r)
                If Not row.EntireRow.Hidden Then
                    visibleRows.Add(row)
                End If
            Next

            For c As Integer = 1 To rng.Columns.Count
                Dim col As Excel.Range = rng.Columns(c)
                If Not col.EntireColumn.Hidden Then
                    visibleCols.Add(col)
                End If
            Next

            If visibleRows.Count = 0 OrElse visibleCols.Count = 0 Then
                ed.WriteMessage(vbLf & "Selection contains no visible rows or columns.")
                Return
            End If

            ' --- Ask insertion point ---
            Dim ppr = ed.GetPoint(vbLf & "Specify insertion point:")
            If ppr.Status <> PromptStatus.OK Then Return

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()

                    Dim bt = CType(tr.GetObject(db.BlockTableId, Teigha.DatabaseServices.OpenMode.ForRead), BlockTable)
                    Dim btr = CType(tr.GetObject(db.CurrentSpaceId, Teigha.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)

                    ' --- Create AutoCAD table with visible row/column count ---
                    Dim tbl As New Table()
                    tbl.TableStyle = db.Tablestyle

                    Dim nRows As Integer = visibleRows.Count
                    Dim nCols As Integer = visibleCols.Count

                    tbl.SetSize(nRows, nCols)

                    ' Default sizes
                    For r = 0 To nRows - 1
                        tbl.SetRowHeight(r, 3.0)
                    Next
                    For c = 0 To nCols - 1
                        tbl.SetColumnWidth(c, 15.0)
                    Next

                    ' --------------------------------------------------------------------
                    '   Fill visible rows & columns ONLY
                    ' --------------------------------------------------------------------
                    Dim acRow As Integer = 0

                    For Each excelRow As Excel.Range In visibleRows
                        Dim acCol As Integer = 0

                        For Each excelCol As Excel.Range In visibleCols
                            ' Get cell
                            Dim cell As Excel.Range = rng.Cells(excelRow.Row - rng.Row + 1,
                                                        excelCol.Column - rng.Column + 1)

                            Dim txt As String = CStr(cell.Text)
                            Dim acCell = tbl.Cells(acRow, acCol)

                            acCell.TextString = txt
                            acCell.TextHeight = 1.5

                            ' Alignment
                            Try
                                Select Case CInt(cell.HorizontalAlignment)
                                    Case Excel.XlHAlign.xlHAlignLeft
                                        acCell.Alignment = CellAlignment.MiddleLeft
                                    Case Excel.XlHAlign.xlHAlignRight
                                        acCell.Alignment = CellAlignment.MiddleRight
                                    Case Excel.XlHAlign.xlHAlignCenter
                                        acCell.Alignment = CellAlignment.MiddleCenter
                                End Select
                            Catch
                            End Try

                            ' Bold = slightly larger text
                            Try
                                If CBool(cell.Font.Bold) Then
                                    acCell.TextHeight = 1.7
                                End If
                            Catch
                            End Try

                            acCol += 1
                        Next

                        acRow += 1
                    Next

                    ' Place table
                    tbl.Position = ppr.Value
                    btr.AppendEntity(tbl)
                    tr.AddNewlyCreatedDBObject(tbl, True)

                    tr.Commit()
                End Using
            End Using

        End Sub

        ' Non-interactive CHSPACE-like command: move all titleblock-contained PS entities to MS via matched viewport transform.
        <CommandMethod("ARCXISCHSPACEALL", CommandFlags.Modal)>
        Public Sub RunChspaceOnTitleBlocks()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            If db.TileMode Then
                ed.WriteMessage(vbLf & "Switch to a layout (paper space) before running ARCXISCHSPACEALL.")
                Return
            End If

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim psBtr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead), BlockTableRecord)
                    Dim msBtr = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                    ' Collect candidate title blocks and viewports in current paperspace
                    Dim titleBlockIds As New List(Of ObjectId)()
                    Dim viewportIds As New List(Of ObjectId)()

                    For Each id As ObjectId In psBtr
                        If id.IsNull OrElse id.ObjectClass Is Nothing Then Continue For
                        Dim dxf = id.ObjectClass.DxfName
                        If String.Equals(dxf, "INSERT", StringComparison.OrdinalIgnoreCase) Then
                            Dim br = TryCast(tr.GetObject(id, OpenMode.ForRead, True), BlockReference)
                            If br Is Nothing Then Continue For
                            If br.IsErased Then Continue For
                            Dim brName As String = String.Empty
                            Try
                                Dim brBtr = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                brName = brBtr.Name
                            Catch
                            End Try
                            If Not String.IsNullOrEmpty(brName) AndAlso
                               (brName.IndexOf("Title", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                brName.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("DPIS TitleBlock", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("DPIS TitleBlock 24x36", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("A$C5F731B9A", StringComparison.OrdinalIgnoreCase)) Then
                                titleBlockIds.Add(id)
                            End If
                        ElseIf String.Equals(dxf, "VIEWPORT", StringComparison.OrdinalIgnoreCase) Then
                            viewportIds.Add(id)
                        End If
                    Next

                    If titleBlockIds.Count = 0 Then
                        ed.WriteMessage(vbLf & "No title blocks found in the current layout.")
                        Return
                    End If

                    Dim processed As Integer = 0

                    For Each tbId In titleBlockIds
                        ' Refresh layout entity list each iteration to avoid accessing erased ids
                        Dim currentLayoutEntities As New List(Of ObjectId)()
                        For Each eid As ObjectId In psBtr
                            currentLayoutEntities.Add(eid)
                        Next

                        Dim tbBr = TryCast(tr.GetObject(tbId, OpenMode.ForRead, True), BlockReference)
                        If tbBr Is Nothing OrElse tbBr.IsErased Then Continue For

                        Dim tbExt As Extents3d
                        Try
                            tbExt = tbBr.GeometricExtents
                        Catch
                            Continue For
                        End Try

                        ' Find a suitable viewport (inside titleblock extents; skip Number=1)
                        Dim vpMatchId As ObjectId = ObjectId.Null
                        Dim vpMatch As Viewport = Nothing
                        For Each vpId In viewportIds
                            Dim vp = TryCast(tr.GetObject(vpId, OpenMode.ForRead, True), Viewport)
                            If vp Is Nothing OrElse vp.IsErased Then Continue For
                            If vp.Number = 1 Then Continue For
                            Dim center As New Point3d(vp.CenterPoint.X, vp.CenterPoint.Y, tbBr.Position.Z)
                            If PointInExtents(center, tbExt) Then
                                vpMatchId = vpId
                                vpMatch = vp
                                Exit For
                            End If
                        Next
                        If vpMatch Is Nothing Then
                            ' Fallback: any non-1 viewport
                            For Each vpId In viewportIds
                                Dim vp = TryCast(tr.GetObject(vpId, OpenMode.ForRead, True), Viewport)
                                If vp IsNot Nothing AndAlso Not vp.IsErased AndAlso vp.Number <> 1 Then
                                    vpMatchId = vpId
                                    vpMatch = vp
                                    Exit For
                                End If
                            Next
                        End If
                        If vpMatch Is Nothing Then
                            ed.WriteMessage(vbLf & "No suitable viewport found for a title block; skipping.")
                            Continue For
                        End If

                        ' Log viewport parameters for diagnostics
                        Try
                            ed.WriteMessage(vbLf & $"Viewport Number={vpMatch.Number}")
                            ed.WriteMessage(vbLf & $"  CenterPoint=({vpMatch.CenterPoint.X:F6}, {vpMatch.CenterPoint.Y:F6}, {vpMatch.CenterPoint.Z:F6})")
                            ed.WriteMessage(vbLf & $"  ViewCenter=({vpMatch.ViewCenter.X:F6}, {vpMatch.ViewCenter.Y:F6})")
                            ed.WriteMessage(vbLf & $"  ViewTarget=({vpMatch.ViewTarget.X:F6}, {vpMatch.ViewTarget.Y:F6}, {vpMatch.ViewTarget.Z:F6})")
                            ed.WriteMessage(vbLf & $"  ViewDirection=({vpMatch.ViewDirection.X:F6}, {vpMatch.ViewDirection.Y:F6}, {vpMatch.ViewDirection.Z:F6})")
                            ed.WriteMessage(vbLf & $"  TwistAngle={vpMatch.TwistAngle:F6}")
                            ed.WriteMessage(vbLf & $"  CustomScale={vpMatch.CustomScale:F6}")
                        Catch
                        End Try

                        ' Build transform from paperspace to model space using viewport
                        Dim M As Matrix3d = BuildPsToMsTransform(vpMatch)

                        ' Diagnostic: log mapping of titleblock position and extents
                        Try
                            Dim tbPos = tbBr.Position
                            Dim tbPosMs = tbPos.TransformBy(M)
                            ed.WriteMessage(vbLf & $"  Diagnostic: TB Position PS=({tbPos.X:F6}, {tbPos.Y:F6}, {tbPos.Z:F6}) -> MS=({tbPosMs.X:F6}, {tbPosMs.Y:F6}, {tbPosMs.Z:F6})")
                            Try
                                Dim minP = tbExt.MinPoint
                                Dim maxP = tbExt.MaxPoint
                                Dim cP As New Point3d((minP.X + maxP.X) / 2.0, (minP.Y + maxP.Y) / 2.0, tbBr.Position.Z)
                                Dim cPms = cP.TransformBy(M)
                                ed.WriteMessage(vbLf & $"  Diagnostic: TB Extents PS Min=({minP.X:F6}, {minP.Y:F6}) Max=({maxP.X:F6}, {maxP.Y:F6}) CenterPS=({cP.X:F6}, {cP.Y:F6}) -> CenterMS=({cPms.X:F6}, {cPms.Y:F6}, {cPms.Z:F6})")
                            Catch
                            End Try
                        Catch
                        End Try

                        ' Gather all PS entities inside/overlapping titleblock extents (excluding the viewport itself)
                        Dim moveIds As New List(Of ObjectId)()
                        For Each entId In currentLayoutEntities
                            If entId = vpMatchId Then Continue For
                            Dim ent = TryCast(tr.GetObject(entId, OpenMode.ForRead, True), Entity)
                            If ent Is Nothing OrElse ent.IsErased Then Continue For
                            ' Skip overall paperspace viewport
                            If TypeOf ent Is Viewport Then
                                Dim v = DirectCast(ent, Viewport)
                                If v.Number = 1 Then Continue For
                            End If
                            Dim eExt As Extents3d
                            Try
                                eExt = ent.GeometricExtents
                            Catch
                                Continue For
                            End Try
                            If ExtentsInside(tbExt, eExt) OrElse ExtentsOverlap(tbExt, eExt) Then
                                moveIds.Add(entId)
                            End If
                        Next
                        ' Ensure titleblock insert moves too
                        If Not moveIds.Contains(tbId) Then moveIds.Add(tbId)

                        If moveIds.Count = 0 Then
                            ed.WriteMessage(vbLf & "Nothing to move for this title block; skipping.")
                            Continue For
                        End If

                        ' Clone to model space, transform, and erase originals
                        For Each entId In moveIds
                            Dim srcEnt = TryCast(tr.GetObject(entId, OpenMode.ForWrite, True), Entity)
                            If srcEnt Is Nothing OrElse srcEnt.IsErased Then Continue For
                            Dim cloned = TryCast(srcEnt.Clone(), Entity)
                            If cloned Is Nothing Then Continue For

                            msBtr.AppendEntity(cloned)
                            tr.AddNewlyCreatedDBObject(cloned, True)

                            Try
                                cloned.TransformBy(M)
                            Catch
                                cloned.Erase()
                                Continue For
                            End Try

                            Try
                                srcEnt.Erase()
                            Catch
                            End Try
                        Next

                        processed += 1
                    Next

                    tr.Commit()
                    ed.WriteMessage(vbLf & $"ARCXISCHSPACEALL processed {processed} title block(s).")
                End Using
            End Using
        End Sub

        ' Command: pick a viewport and dump safe properties to the command line
        <CommandMethod("ARCXISVPINFO", CommandFlags.Modal)>
        Public Sub ViewportInfo()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            If db.TileMode Then
                ed.WriteMessage(vbLf & "Switch to a layout (paper space) before running ARCXISVPINFO.")
                Return
            End If

            ' Prompt to select a viewport entity
            Dim peo As New PromptEntityOptions(vbLf & "Select a viewport")
            peo.SetRejectMessage(vbLf & "Entity is not a viewport.")
            peo.AddAllowedClass(GetType(Viewport), True)
            Dim per = ed.GetEntity(peo)
            If per.Status <> PromptStatus.OK Then Return

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim vp = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead), Viewport)
                    If vp Is Nothing Then
                        ed.WriteMessage(vbLf & "Selected entity was not a viewport.")
                        Return
                    End If

                    ' Basic state (safe)
                    ed.WriteMessage(vbLf & $"Viewport Id={per.ObjectId}, Number={vp.Number}, On={vp.On}, Locked={vp.Locked}")

                    ' Geometry and placement (safe)
                    ed.WriteMessage(vbLf & $"CenterPoint=({vp.CenterPoint.X:F6}, {vp.CenterPoint.Y:F6}, {vp.CenterPoint.Z:F6})")
                    ed.WriteMessage(vbLf & $"Width={vp.Width:F6}, Height={vp.Height:F6}")
                    ed.WriteMessage(vbLf & $"TwistAngle={vp.TwistAngle:F6}")

                    ' View parameters (safe)
                    ed.WriteMessage(vbLf & $"ViewCenter=({vp.ViewCenter.X:F6}, {vp.ViewCenter.Y:F6})")
                    ed.WriteMessage(vbLf & $"ViewTarget=({vp.ViewTarget.X:F6}, {vp.ViewTarget.Y:F6}, {vp.ViewTarget.Z:F6})")
                    ed.WriteMessage(vbLf & $"ViewDirection=({vp.ViewDirection.X:F6}, {vp.ViewDirection.Y:F6}, {vp.ViewDirection.Z:F6})")
                    ed.WriteMessage(vbLf & $"ViewHeight={vp.ViewHeight:F6}")
                    ed.WriteMessage(vbLf & $"CustomScale={vp.CustomScale:F6}")

                    ' Optional/sometimes unsupported properties (guarded)
                    Try
                        ed.WriteMessage(vbLf & $"StandardScale={vp.StandardScale}")
                    Catch
                    End Try

                    tr.Commit()
                End Using
            End Using
        End Sub

        ' Command: diagnose viewport locations and mapped model-space corners
        <CommandMethod("ARCXISVPDIAG", CommandFlags.Modal)>
        Public Sub ViewportDiagnostics()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            If db.TileMode Then
                ed.WriteMessage(vbLf & "Switch to a layout (paper space) before running ARCXISVPDIAG.")
                Return
            End If

            ' Prompt to select a viewport entity
            Dim peo As New PromptEntityOptions(vbLf & "Select a viewport for diagnostics")
            peo.SetRejectMessage(vbLf & "Entity is not a viewport.")
            peo.AddAllowedClass(GetType(Viewport), True)
            Dim per = ed.GetEntity(peo)
            If per.Status <> PromptStatus.OK Then Return

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim vp = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead), Viewport)
                    If vp Is Nothing Then
                        ed.WriteMessage(vbLf & "Selected entity was not a viewport.")
                        Return
                    End If

                    ed.WriteMessage(vbLf & $"--- Viewport Diagnostic ---")
                    ed.WriteMessage(vbLf & $"Viewport Id={per.ObjectId}, Number={vp.Number}, On={vp.On}, Locked={vp.Locked}")
                    ed.WriteMessage(vbLf & $"CenterPoint (PS)=({vp.CenterPoint.X:F6}, {vp.CenterPoint.Y:F6}, {vp.CenterPoint.Z:F6})")
                    ed.WriteMessage(vbLf & $"ViewCenter (model XY)=({vp.ViewCenter.X:F6}, {vp.ViewCenter.Y:F6})")
                    ed.WriteMessage(vbLf & $"ViewTarget (model XYZ)=({vp.ViewTarget.X:F6}, {vp.ViewTarget.Y:F6}, {vp.ViewTarget.Z:F6})")
                    ed.WriteMessage(vbLf & $"ViewDirection=({vp.ViewDirection.X:F6}, {vp.ViewDirection.Y:F6}, {vp.ViewDirection.Z:F6})")
                    ed.WriteMessage(vbLf & $"TwistAngle={vp.TwistAngle:F6}, CustomScale={vp.CustomScale:F6}")
                    ed.WriteMessage(vbLf & $"Width={vp.Width:F6}, Height={vp.Height:F6}")

                    ' Build transform and show mapped points
                    Dim M = BuildPsToMsTransform(vp)

                    Try
                        ' Map PS center point
                        Dim centerPS = vp.CenterPoint
                        Dim centerMS = centerPS.TransformBy(M)
                        ed.WriteMessage(vbLf & $"Mapped PS Center -> MS = ({centerMS.X:F6}, {centerMS.Y:F6}, {centerMS.Z:F6})")

                        ' Map computed model center used by transform (ViewCenter XY + ViewTarget.Z)
                        Dim vc = vp.ViewCenter
                        Dim modelCenter As New Point3d(vc.X, vc.Y, vp.ViewTarget.Z)
                        ed.WriteMessage(vbLf & $"Model center used = ({modelCenter.X:F6}, {modelCenter.Y:F6}, {modelCenter.Z:F6})")

                        ' Compute PS rectangle corners (in PS coordinates) around viewport CenterPoint using Width/Height
                        Dim halfW = vp.Width / 2.0
                        Dim halfH = vp.Height / 2.0
                        Dim cornersPS As New List(Of Point3d) From {
                            New Point3d(centerPS.X - halfW, centerPS.Y - halfH, centerPS.Z),
                            New Point3d(centerPS.X - halfW, centerPS.Y + halfH, centerPS.Z),
                            New Point3d(centerPS.X + halfW, centerPS.Y + halfH, centerPS.Z),
                            New Point3d(centerPS.X + halfW, centerPS.Y - halfH, centerPS.Z)
                        }

                        Dim i As Integer = 0
                        Dim minX As Double = Double.MaxValue, minY As Double = Double.MaxValue
                        Dim maxX As Double = -Double.MaxValue, maxY As Double = -Double.MaxValue

                        For Each cp In cornersPS
                            Dim mp = cp.TransformBy(M)
                            ed.WriteMessage(vbLf & $"Corner {i}: PS=({cp.X:F6}, {cp.Y:F6}) -> MS=({mp.X:F6}, {mp.Y:F6}, {mp.Z:F6})")
                            If mp.X < minX Then minX = mp.X
                            If mp.Y < minY Then minY = mp.Y
                            If mp.X > maxX Then maxX = mp.X
                            If mp.Y > maxY Then maxY = mp.Y
                            i += 1
                        Next

                        ed.WriteMessage(vbLf & $"Mapped MS extents from corners: MinX={minX:F6}, MinY={minY:F6}, MaxX={maxX:F6}, MaxY={maxY:F6}")

                        ' Map the ViewTarget point for comparison
                        Dim vt = vp.ViewTarget
                        Dim vtMapped = vt.TransformBy(M)
                        ed.WriteMessage(vbLf & $"ViewTarget mapped -> MS = ({vtMapped.X:F6}, {vtMapped.Y:F6}, {vtMapped.Z:F6})")

                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "Diagnostic transform failed: " & ex.Message)
                    End Try

                    tr.Commit()
                End Using
            End Using
        End Sub

        ' Command: arrange titleblocks in a grid with 150' spacing
        <CommandMethod("ARCXISGRIDTB", CommandFlags.Modal)>
        Public Sub ArrangeTitleblocksGrid()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            If db.TileMode Then
                'ed.WriteMessage(vbLf & "Run this command in Model Space.")
                'Return
            End If

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim ms = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                    ' Collect title block references in model space (matching same names as earlier)
                    Dim tbRefs As New List(Of BlockReference)()
                    For Each id As ObjectId In ms
                        Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead, True), Entity)
                        If ent Is Nothing OrElse ent.IsErased Then Continue For

                        ' Only consider block references (INSERT)
                        Dim br = TryCast(ent, BlockReference)
                        If br Is Nothing Then Continue For

                        Try
                            Dim brBtr = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            Dim brName = brBtr.Name
                            If Not String.IsNullOrEmpty(brName) AndAlso
                               (brName.IndexOf("Title", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                brName.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("DPIS TitleBlock", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("DPIS TitleBlock 24x36", StringComparison.OrdinalIgnoreCase) OrElse
                                brName.Equals("A$C5F731B9A", StringComparison.OrdinalIgnoreCase)) Then
                                tbRefs.Add(br)
                            End If
                        Catch
                        End Try
                    Next

                    If tbRefs.Count = 0 Then
                        ed.WriteMessage(vbLf & "No title block inserts found in model space.")
                        Return
                    End If

                    ' Sort by current position: top-to-bottom (Y desc), left-to-right (X asc)
                    tbRefs.Sort(Function(a, b)
                                    Dim ay = a.Position.Y
                                    Dim by = b.Position.Y
                                    If Math.Abs(ay - by) > 0.000001 Then
                                        Return -ay.CompareTo(by) ' descending Y
                                    End If
                                    Return a.Position.X.CompareTo(b.Position.X) ' ascending X
                                End Function)

                    Dim n = tbRefs.Count
                    Dim cols As Integer = CInt(Math.Ceiling(Math.Sqrt(n)))
                    Dim rows As Integer = CInt(Math.Ceiling(n / cols))

                    ' Compute bounding left/top from existing positions to preserve relative origin
                    Dim minX = Double.MaxValue, maxX = -Double.MaxValue, maxY = -Double.MaxValue
                    For Each br In tbRefs
                        If br.Position.X < minX Then minX = br.Position.X
                        If br.Position.X > maxX Then maxX = br.Position.X
                        If br.Position.Y > maxY Then maxY = br.Position.Y
                    Next

                    ' Spacing: 150 feet -> 1800 drawing units (12 units/ft). Anchor offset: 2500 ft -> 30000 units.
                    Dim spacingUnits As Double = 1800.0
                    Dim anchorOffsetUnits As Double = 2500.0 * 12.0
                    ' Start far to the right of existing content
                    Dim startX As Double = maxX + anchorOffsetUnits

                    ' Cluster by Y within ~50 feet to preserve rows
                    Dim yTolUnits As Double = 50.0 * 12.0
                    Dim rowsClusters As New List(Of List(Of BlockReference))()

                    ' Sort by Y descending, then X ascending to build clusters
                    Dim sorted As New List(Of BlockReference)(tbRefs)
                    sorted.Sort(Function(a, b)
                                    Dim ay = a.Position.Y
                                    Dim by = b.Position.Y
                                    If Math.Abs(ay - by) > 0.000001 Then
                                        Return -ay.CompareTo(by)
                                    End If
                                    Return a.Position.X.CompareTo(b.Position.X)
                                End Function)

                    For Each br In sorted
                        Dim placed As Boolean = False
                        For Each cluster In rowsClusters
                            ' Compare with the first item's Y in the cluster
                            Dim anchorY = cluster(0).Position.Y
                            If Math.Abs(br.Position.Y - anchorY) <= yTolUnits Then
                                cluster.Add(br)
                                placed = True
                                Exit For
                            End If
                        Next
                        If Not placed Then
                            rowsClusters.Add(New List(Of BlockReference) From {br})
                        End If
                    Next

                    ' Sort each cluster by X ascending
                    For Each cluster In rowsClusters
                        cluster.Sort(Function(a, b) a.Position.X.CompareTo(b.Position.X))
                    Next

                    ' Arrange clusters into grid using anchor startX and spacingUnits
                    Dim processedEnts As New HashSet(Of ObjectId)()

                    For rowIdx As Integer = 0 To rowsClusters.Count - 1
                        Dim cluster = rowsClusters(rowIdx)
                        For colIdx As Integer = 0 To cluster.Count - 1
                            Dim br = cluster(colIdx)

                            Dim targetX = startX + colIdx * spacingUnits
                            Dim targetY = maxY - rowIdx * spacingUnits
                            Dim targetZ = br.Position.Z

                            Dim oldPos = br.Position

                            ' Capture extents BEFORE moving to identify the group
                            Dim brExtBefore As Extents3d
                            Try
                                brExtBefore = br.GeometricExtents
                            Catch
                                brExtBefore = New Extents3d(New Point3d(oldPos.X - 1, oldPos.Y - 1, oldPos.Z), New Point3d(oldPos.X + 1, oldPos.Y + 1, oldPos.Z))
                            End Try

                            Dim disp As New Vector3d(targetX - oldPos.X, targetY - oldPos.Y, targetZ - oldPos.Z)

                            ' Move the block reference
                            Try
                                Dim brWrite = CType(tr.GetObject(br.ObjectId, OpenMode.ForWrite), BlockReference)
                                brWrite.TransformBy(Matrix3d.Displacement(disp))
                                processedEnts.Add(br.ObjectId)
                            Catch ex As Exception
                                ed.WriteMessage(vbLf & $"Failed moving titleblock {br.ObjectId}: {ex.Message}")
                                Continue For
                            End Try

                            ' Move overlapping entities with the same displacement
                            For Each entId As ObjectId In ms
                                If processedEnts.Contains(entId) Then Continue For
                                If entId = br.ObjectId Then Continue For
                                Dim ent = TryCast(tr.GetObject(entId, OpenMode.ForRead, True), Entity)
                                If ent Is Nothing OrElse ent.IsErased Then Continue For
                                If TypeOf ent Is Viewport Then Continue For

                                Dim entExt As Extents3d
                                Try
                                    entExt = ent.GeometricExtents
                                Catch
                                    Continue For
                                End Try

                                If ExtentsOverlap(brExtBefore, entExt) OrElse ExtentsInside(brExtBefore, entExt) OrElse ExtentsInside(entExt, brExtBefore) Then
                                    Try
                                        Dim entW = CType(tr.GetObject(entId, OpenMode.ForWrite, True), Entity)
                                        entW.TransformBy(Matrix3d.Displacement(disp))
                                        processedEnts.Add(entId)
                                    Catch
                                    End Try
                                End If
                            Next

                            ' Label sequence number within row-major order
                            Try
                                Using txt As New DBText()
                                    txt.SetDatabaseDefaults()
                                    txt.Position = New Point3d(targetX, targetY, targetZ)
                                    txt.Height = 24.0
                                    Dim seqNum As Integer = rowIdx * cluster.Count + (colIdx + 1)
                                    txt.TextString = seqNum.ToString()
                                    txt.Layer = "0"
                                    ms.AppendEntity(txt)
                                    tr.AddNewlyCreatedDBObject(txt, True)
                                End Using
                            Catch
                            End Try
                        Next
                    Next

                    tr.Commit()
                    Dim spacingFeetVal As Double = spacingUnits / 12.0
                    ed.WriteMessage(vbLf & $"ARCXISGRIDTB clustered {tbRefs.Count} titleblocks into {rowsClusters.Count} row(s) with Y tolerance {yTolUnits / 12.0}' and spacing {spacingFeetVal}'.")
                End Using
            End Using
        End Sub

        Private Function BuildPsToMsTransform(vp As Viewport) As Matrix3d
            ' Compose PS→MS using viewport properties (no UCS):
            ' - CenterPoint (PS)
            ' - TwistAngle
            ' - Fixed scale (1/96)
            ' - ViewDirection (WCS)
            ' - ViewCenter (WCS XY) + ViewTarget (WCS XYZ) for final model translation
            Dim T_toOrigin = Matrix3d.Displacement(New Vector3d(-vp.CenterPoint.X, -vp.CenterPoint.Y, 0))
            Dim R_untwist = Matrix3d.Rotation(-vp.TwistAngle, Vector3d.ZAxis, Point3d.Origin)

            ' Build orthonormal basis from ViewDirection
            Dim viewDir = vp.ViewDirection.GetNormal()
            Dim xSeed As Vector3d = Vector3d.XAxis
            If Math.Abs(xSeed.DotProduct(viewDir)) > 0.95 Then xSeed = Vector3d.YAxis
            Dim xProj = (xSeed - viewDir * xSeed.DotProduct(viewDir)).GetNormal()
            Dim yVec = viewDir.CrossProduct(xProj).GetNormal()
            Dim A_align = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                                                         Point3d.Origin, xProj, yVec, viewDir)

            ' Use ViewCenter (2D) as the view center in model XY and add ViewTarget XY; use ViewTarget.Z for Z
            Dim vc = vp.ViewCenter
            Dim modelCenterX = vc.X + vp.ViewTarget.X
            Dim modelCenterY = vc.Y + vp.ViewTarget.Y
            Dim modelCenterZ = vp.ViewTarget.Z
            Dim T_toModel = Matrix3d.Displacement(New Vector3d(modelCenterX, modelCenterY, modelCenterZ))

            ' Force fixed scale of 96
            Dim scaleFactor As Double = 96.0
            Dim S_scale_fixed = Matrix3d.Scaling(scaleFactor, Point3d.Origin)

            Return T_toModel * A_align * S_scale_fixed * R_untwist * T_toOrigin
        End Function

        Private Function PointInExtents(pt As Point3d, ext As Extents3d) As Boolean
            Return pt.X >= ext.MinPoint.X AndAlso pt.X <= ext.MaxPoint.X AndAlso
                   pt.Y >= ext.MinPoint.Y AndAlso pt.Y <= ext.MaxPoint.Y
        End Function

        Private Function ExtentsOverlap(a As Extents3d, b As Extents3d) As Boolean
            Return Not (b.MinPoint.X > a.MaxPoint.X OrElse b.MaxPoint.X < a.MinPoint.X OrElse
                        b.MinPoint.Y > a.MaxPoint.Y OrElse b.MaxPoint.Y < a.MinPoint.Y)
        End Function

        Private Function ExtentsInside(container As Extents3d, inner As Extents3d) As Boolean
            Return inner.MinPoint.X >= container.MinPoint.X AndAlso inner.MaxPoint.X <= container.MaxPoint.X AndAlso
                   inner.MinPoint.Y >= container.MinPoint.Y AndAlso inner.MaxPoint.Y <= container.MaxPoint.Y
        End Function
        <CommandMethod("UPDATEATTRIBDEFS")>
        Public Shared Sub UpdateAttributeDefinitions()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            ' Define attribute mappings (old tag -> new tag)
            Dim attributeMappings As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"LXS", "VENTTYPE"},
        {"TND", "ATTICTYPE"},
        {"MOD", "PLANTYPE"}
    }

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                    ' Check if the AtticVentPage block exists
                    If Not bt.Has("AtticVentPage") Then
                        ed.WriteMessage(vbLf & "Block 'AtticVentPage' not found in this drawing.")
                        tr.Commit()
                        Return
                    End If

                    ' Get the block definition for AtticVentPage
                    Dim btrId As ObjectId = bt("AtticVentPage")
                    Dim btr = CType(tr.GetObject(btrId, OpenMode.ForWrite), BlockTableRecord)

                    ' Dictionary to store ALL attribute values: BlockRefId -> (Tag -> Value)
                    Dim blockAttributeValues As New Dictionary(Of ObjectId, Dictionary(Of String, String))()

                    ' Step 1: Collect ALL attribute values from existing block references
                    ed.WriteMessage(vbLf & "Collecting attribute values...")
                    For Each spaceId As ObjectId In {bt(BlockTableRecord.ModelSpace), bt(BlockTableRecord.PaperSpace)}
                        Dim spaceBtr = CType(tr.GetObject(spaceId, OpenMode.ForRead), BlockTableRecord)
                        For Each entId As ObjectId In spaceBtr
                            Dim br = TryCast(tr.GetObject(entId, OpenMode.ForRead), BlockReference)
                            If br Is Nothing OrElse br.BlockTableRecord <> btrId Then Continue For

                            ' Store ALL attribute values for this block reference
                            Dim attValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                            Dim attCol = br.AttributeCollection
                            If attCol IsNot Nothing Then
                                For Each attRefId As ObjectId In attCol
                                    Dim attRef = CType(tr.GetObject(attRefId, OpenMode.ForRead), AttributeReference)
                                    attValues(attRef.Tag.ToUpper()) = attRef.TextString
                                Next
                            End If

                            blockAttributeValues(br.ObjectId) = attValues
                        Next
                    Next

                    If blockAttributeValues.Count = 0 Then
                        ed.WriteMessage(vbLf & "No 'AtticVentPage' block references found in this drawing.")
                        tr.Commit()
                        Return
                    End If

                    ' Step 2: Update attribute definitions in the block definition
                    ed.WriteMessage(vbLf & "Updating attribute definitions...")
                    For Each entId As ObjectId In btr
                        Dim attDef = TryCast(tr.GetObject(entId, OpenMode.ForRead), AttributeDefinition)
                        If attDef Is Nothing Then Continue For

                        Dim tagUpper = attDef.Tag.ToUpper()
                        If attributeMappings.ContainsKey(tagUpper) Then
                            Dim attDefWrite = CType(tr.GetObject(attDef.ObjectId, OpenMode.ForWrite), AttributeDefinition)
                            attDefWrite.Tag = attributeMappings(tagUpper)
                        End If
                    Next

                    ' Step 3: Update all block references with new attribute structure
                    ed.WriteMessage(vbLf & "Synchronizing block references...")
                    For Each kvp In blockAttributeValues
                        Dim brId = kvp.Key
                        Dim oldValues = kvp.Value

                        Dim br = CType(tr.GetObject(brId, OpenMode.ForWrite), BlockReference)

                        ' Remove old attributes
                        For Each attRefId As ObjectId In br.AttributeCollection
                            Dim attRef = CType(tr.GetObject(attRefId, OpenMode.ForWrite), AttributeReference)
                            attRef.Erase()
                        Next

                        ' Recreate ALL attributes with preserved values
                        For Each defEntId As ObjectId In btr
                            Dim attDef = TryCast(tr.GetObject(defEntId, OpenMode.ForRead), AttributeDefinition)
                            If attDef Is Nothing OrElse attDef.Constant Then Continue For

                            Using attRef As New AttributeReference()
                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform)

                                ' Restore value - check both old tag and new tag
                                Dim newTag = attDef.Tag.ToUpper()
                                Dim valueRestored As Boolean = False

                                ' First, check if this new tag came from a mapped old tag
                                For Each mapping In attributeMappings
                                    If newTag = mapping.Value.ToUpper() Then
                                        ' This is a renamed attribute - look for old tag value
                                        If oldValues.ContainsKey(mapping.Key.ToUpper()) Then
                                            attRef.TextString = oldValues(mapping.Key.ToUpper())
                                            valueRestored = True
                                            Exit For
                                        End If
                                    End If
                                Next

                                ' If not restored yet, check if value exists under current tag name
                                If Not valueRestored AndAlso oldValues.ContainsKey(newTag) Then
                                    attRef.TextString = oldValues(newTag)
                                End If

                                br.AttributeCollection.AppendAttribute(attRef)
                                tr.AddNewlyCreatedDBObject(attRef, True)
                            End Using
                        Next
                    Next

                    tr.Commit()
                    ed.WriteMessage(vbLf & $"Successfully updated {blockAttributeValues.Count} 'AtticVentPage' block reference(s).")
                    ed.WriteMessage(vbLf & "Attribute mappings: LXS→VENTTYPE, TND→ATTICTYPE, MOD→PLANTYPE")
                End Using
            End Using

            Using doc.LockDocument()

                Dim folder As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
                Dim name As String = CStr(Application.GetSystemVariable("DWGNAME"))
                Dim path = System.IO.Path.Combine(folder, name)

                db.SaveAs(path, True, DwgVersion.Current, db.SecurityParameters)

            End Using

        End Sub
    End Class

End Namespace
