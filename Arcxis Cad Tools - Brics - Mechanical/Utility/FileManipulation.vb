Imports System
Imports System.Collections
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Security.Policy
Imports System.Threading
Imports System.Windows.Forms
Imports PdfSharp
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Path = System.IO.Path
Imports Bricscad.ApplicationServices
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document
Imports Exception = Teigha.Runtime.Exception
Imports Layout = Teigha.DatabaseServices.Layout
Imports Color = Teigha.Colors.Color
Imports System.Collections.Specialized


' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.FileManipulation))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class FileManipulation

        ' === CSV accumulation (list of lists) ===
        Private Shared _pendingRows As New List(Of List(Of String))()
        ' Add this at the top of your FileManipulation class
        Private Shared _fontResolverInitialized As Boolean = False
        Private Shared ReadOnly _fontResolverLock As New Object()

        Private Shared Sub EnsureFontResolver()
            If _fontResolverInitialized Then Return

            SyncLock _fontResolverLock
                If Not _fontResolverInitialized Then
                    Try
                        PdfSharp.Fonts.GlobalFontSettings.FontResolver = New SystemFontResolver()
                        _fontResolverInitialized = True
                    Catch ex As Exception
                        ' Log but don't fail - let individual operations handle font issues
                        Debug.WriteLine($"Failed to initialize font resolver: {ex.Message}")
                    End Try
                End If
            End SyncLock
        End Sub

        <CommandMethod("RedoPaths", CommandFlags.Modal)>
        Sub FixPaths()
            Module_Arcxis_TB.AddNewPaths()
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

        Public Shared Function GetCustomDwgPropForDoc(db As Database, propName As String) As String
            'Dim db = Application.DocumentManager.MdiActiveDocument.Database

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




        <CommandMethod("AMP")>
        Sub CallMechanicalProps()
            Dim frm As New Form_MechanicalPrinting
            frm.ShowDialog()
        End Sub

        Private Shared Sub DeleteFilesSafe(files As IEnumerable(Of String))
            If files Is Nothing Then Exit Sub
            For Each f In files.Distinct(StringComparer.OrdinalIgnoreCase)
                Try
                    If IO.File.Exists(f) Then IO.File.Delete(f)
                Catch
                    ' best-effort; ignore individual failures
                End Try
            Next
        End Sub

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

        Public Sub TurnOnOrOffLayer(layerName As String, TurnOn As Boolean)
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acDoc.LockDocument()
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim lt As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                    If lt.Has(layerName) Then

                        Dim lyrId As ObjectId = lt(layerName)
                        Dim layer As LayerTableRecord = acTrans.GetObject(lyrId, OpenMode.ForWrite)

                        If TurnOn Then

                            layer.IsOff = False
                            layer.IsFrozen = False

                        Else

                            layer.IsOff = True
                            layer.IsFrozen = True

                        End If

                    End If

                    acTrans.Commit()
                End Using
            End Using
        End Sub

        ' Reworked: write ONE logical set per PDF as Identifier/Key/Value rows.
        ' Column 1: Identifier = "Builder - PdfName"
        ' Column 2: Key        = PLAN | ELEV | SW | MOD
        ' Column 3: Value
        ' Creates a small CSV per PDF (4 data rows + header).

        ' Helper to write one key/value row with proper CSV quoting.
        Private Sub WriteKv(sw As StreamWriter, identifier As String, key As String, valueStr As String)
            sw.WriteLine($"{CsvQuote(identifier)},{CsvQuote(key)},{CsvQuote(If(valueStr, ""))}")
        End Sub
        ' Reworked: one logical set per PDF.
        ' Columns:
        '   Col1 Identifier: "Builder - PdfName"
        '   Col2 Key: PLAN | ELEV | SW | MOD
        '   Col3 Value
        ' Uses System.Tuple explicitly to avoid ambiguity with OpenXml types.

        Public Shared Sub CreateViewportSizeRectangle(layoutid As ObjectId)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutid, OpenMode.ForWrite), Layout)

                    Dim btr As BlockTableRecord = acTrans.GetObject(acLayout.BlockTableRecordId, OpenMode.ForWrite)

                    ' Add a rectangle (as polyline)
                    Dim rect As New Polyline()
                    rect.Layer = "0"
                    rect.AddVertexAt(0, New Point2d(0, 0), 0, 0, 0)
                    rect.AddVertexAt(1, New Point2d(17, 0), 0, 0, 0)
                    rect.AddVertexAt(2, New Point2d(17, 11), 0, 0, 0)
                    rect.AddVertexAt(3, New Point2d(0, 11), 0, 0, 0)
                    rect.AddVertexAt(4, New Point2d(0, 0), 0, 0, 0)
                    rect.Closed = True
                    btr.AppendEntity(rect)
                    acTrans.AddNewlyCreatedDBObject(rect, True)
                    acTrans.Commit()
                End Using

            End Using

        End Sub

        Public Shared Function BlockImport(targetDb As Database, blockName As String, sourceDwgPath As String, actrans As Transaction) As ObjectId

            Dim bt As BlockTable = actrans.GetObject(targetDb.BlockTableId, OpenMode.ForRead)

            ' If the block already exists, check attribute count and possibly rename
            If bt.Has(blockName) Then
                Dim btr As BlockTableRecord = actrans.GetObject(bt(blockName), OpenMode.ForWrite)
                ' Count AttributeDefinitions (not AttributeReferences)
                Dim attrCount As Integer = 0
                For Each entId As ObjectId In btr
                    Dim ent As DBObject = actrans.GetObject(entId, OpenMode.ForRead)
                    If TypeOf ent Is AttributeDefinition Then
                        attrCount += 1
                    End If
                Next
                If attrCount <> 40 Then
                    ' Rename block to "old " + blockName if not already renamed
                    If Not btr.Name.StartsWith("old ", StringComparison.OrdinalIgnoreCase) Then
                        ' Ensure the new name does not already exist
                        Dim newName As String = "old " & blockName
                        Dim uniqueName As String = newName
                        Dim suffix As Integer = 1
                        While bt.Has(uniqueName)
                            uniqueName = newName & "_" & suffix
                            suffix += 1
                        End While
                        btr.Name = uniqueName
                    End If
                    ' Refresh BlockTable reference after rename
                    bt = actrans.GetObject(targetDb.BlockTableId, OpenMode.ForRead)
                Else
                    Return bt(blockName)
                End If
            End If

            ' Load external DWG in a side database
            Using sourceDb As New Database(False, True)
                sourceDb.ReadDwgFile(sourceDwgPath, FileOpenMode.OpenForReadAndAllShare, False, "")

                ' Get block table from source
                Using sourceTrans As Transaction = sourceDb.TransactionManager.StartTransaction()

                    Dim sourceBT As BlockTable = sourceTrans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)

                    Dim sourceBlockId As ObjectId = sourceBT(blockName)
                    Dim idsToClone As New ObjectIdCollection()

                    idsToClone.Add(sourceBlockId)

                    ' Clone block definition into target DB
                    Dim idMapping As New IdMapping()

                    sourceDb.WblockCloneObjects(idsToClone, targetDb.BlockTableId, idMapping, DuplicateRecordCloning.Replace, False)

                    sourceTrans.Commit()
                End Using
            End Using

            ' Return the newly added block
            bt = actrans.GetObject(targetDb.BlockTableId, OpenMode.ForRead)

            ' Avoid reloading xrefs inside an active transaction; caller should reload after commit if needed.

            Return bt(blockName)
        End Function

        Public Shared Sub CreateViewport(layoutid)
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acDoc.LockDocument()
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim layId As ObjectId = CType(layoutid, ObjectId)
                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layId, OpenMode.ForWrite), Layout)
                    Dim btr As BlockTableRecord = acTrans.GetObject(acLayout.BlockTableRecordId, OpenMode.ForWrite)

                    ' Desired geometry
                    Dim viewportCenter As New Point3d(7.706, 5.5, 0)
                    Dim viewportWidth As Double = 14.79
                    Dim viewportHeight As Double = 10.5
                    Dim modelViewCenter As New Point2d(8.5, 5.5)
                    Dim modelViewHeight As Double = 11.588

                    ' Try to find an existing paperspace viewport (Number > 1)
                    Dim targetVp As Viewport = Nothing
                    For Each id As ObjectId In btr
                        Dim vp As Viewport = TryCast(acTrans.GetObject(id, OpenMode.ForWrite), Viewport)
                        If vp Is Nothing Then Continue For
                        If vp.Number > 1 Then
                            If vp.Height < 10.51 Then
                                targetVp = vp
                                Exit For
                            End If
                        End If
                    Next

                    If targetVp Is Nothing Then
                        ' Create new if none found
                        targetVp = New Viewport()
                        btr.AppendEntity(targetVp)
                        acTrans.AddNewlyCreatedDBObject(targetVp, True)
                    End If

                    ' Initialize/update the viewport the same way
                    targetVp.SetDatabaseDefaults()
                    targetVp.CustomScale = 1 / 96
                    targetVp.CenterPoint = viewportCenter
                    targetVp.Width = viewportWidth
                    targetVp.Height = viewportHeight
                    targetVp.ViewCenter = modelViewCenter
                    targetVp.ViewHeight = modelViewHeight
                    targetVp.StandardScale = StandardScaleType.Scale1To8inchAnd1ft
                    targetVp.ViewTarget = Point3d.Origin
                    targetVp.ViewDirection = New Vector3d(0, 0, 1)
                    targetVp.TwistAngle = 0.0
                    targetVp.On = True
                    targetVp.Layer = "0"
                    'targetVp.UpdateDisplay()
                    acTrans.Commit()
                End Using
            End Using
        End Sub

        Public Shared Sub CreateLayoutsWithTitleblock(Optional ByVal NumOfLayouts As Integer = 0, Optional ByVal RanFromAutomation As Boolean = False, Optional ByVal CustomCTB As String = "")
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurdb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor

            If Application.GetSystemVariable("LAYOUTREGENCTL") <> 0 Then
                Application.SetSystemVariable("LAYOUTREGENCTL", 0)
            End If

            Using acLckDoc As DocumentLock = acDoc.LockDocument()
                Using acTrans As Transaction = acCurdb.TransactionManager.StartTransaction()

                    Dim lm As LayoutManager = LayoutManager.Current
                    Dim bt As BlockTable = acTrans.GetObject(acCurdb.BlockTableId, OpenMode.ForRead)
                    Dim layoutDict As DBDictionary = TryCast(acTrans.GetObject(acCurdb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                    Dim KeepNames As New List(Of String)
                    Dim toDelete As New List(Of ObjectId)()

                    ' === BLOCK NAME AND PATH ===
                    Dim blockName As String = "Arcxis Title Block"
                    Dim sourceDwgPath As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\Arcxis\Engineering\Drafting Standards\CAD Blocks\Arcxis Title Block - Block.dwg"


                    ' === Check if block is already loaded ===
                    Dim blockDefId As ObjectId
                    ' === Load block from external DWG ===
                    blockDefId = BlockImport(acCurdb, blockName, sourceDwgPath, acTrans)

                    Dim existingInitialRevValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For Each btrId As ObjectId In bt
                        Dim btr As BlockTableRecord = acTrans.GetObject(btrId, OpenMode.ForRead)
                        For Each entId As ObjectId In btr
                            Dim ent As Entity = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Entity)
                            If TypeOf ent Is BlockReference Then
                                Dim blkRef As BlockReference = CType(ent, BlockReference)
                                Dim blkDef As BlockTableRecord = acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead)
                                If blkDef.Name.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) Or blkDef.Name.Equals("DPIS RevisionBlock", StringComparison.OrdinalIgnoreCase) Then
                                    For Each attId As ObjectId In blkRef.AttributeCollection
                                        Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForRead), AttributeReference)
                                        If attRef IsNot Nothing AndAlso (attRef.Tag Like "*INITIAL*" Or attRef.Tag Like "*REV*") Then
                                            If attRef.TextString <> "" And attRef.TextString <> "-" Then
                                                existingInitialRevValues(attRef.Tag) = attRef.TextString
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    Next

                    ' === Customize These Parameters ===
                    Dim layoutBaseName As String = ""  ' "" gives layout names like "1", "2", etc.

                    For i As Integer = 1 To NumOfLayouts

                        Dim layoutName As String = layoutBaseName & i.ToString()

                        Dim layoutId As ObjectId
                        KeepNames.Add(layoutName)
                        If LayoutExistsByDictionary(acCurdb, layoutName) Then
                            layoutId = lm.GetLayoutId(layoutName)
                        Else
                            layoutId = lm.CreateLayout(layoutName)
                        End If

                        Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)
                        Dim btr1 As BlockTableRecord = acTrans.GetObject(acLayout.BlockTableRecordId, OpenMode.ForWrite)

                        lm.CurrentLayout = acLayout.LayoutName

                        ' === Now erase all objects in the layout ===
                        For Each objId As ObjectId In btr1
                            Dim obj As DBObject = acTrans.GetObject(objId, OpenMode.ForWrite)

                            Dim keep As Boolean = False

                            ' Keep any existing paperspace viewport(s)
                            If TypeOf obj Is Viewport Then
                                Dim vpp As Viewport = TryCast(acTrans.GetObject(objId, OpenMode.ForWrite), Viewport)
                                If vpp IsNot Nothing Then
                                    ' Never delete the overall paperspace viewport (Number = 1)
                                    If vpp.BlockName = "*Paper_Space" Then
                                        keep = True
                                    ElseIf vpp.Height > 11 Or vpp.Width > 17 Then
                                        keep = False
                                    End If
                                End If
                            ElseIf TypeOf obj Is BlockReference Then
                                Dim blkRef1 As BlockReference = CType(obj, BlockReference)
                                Dim blkDef As BlockTableRecord = acTrans.GetObject(blkRef1.BlockTableRecord, OpenMode.ForRead)
                                If blkDef.Name.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) Then
                                    keep = True
                                End If
                            End If

                            If Not keep Then
                                obj.Erase()
                            End If
                        Next

                        Application.SetSystemVariable("PSLTSCALE", 0)

                        ' Ensure the rectangle and a viewport exist
                        CreateViewportSizeRectangle(layoutId)
                        CreateViewport(layoutId)
                        PageSetUp11x17(acLayout.LayoutName, CustomCTB)

                        ' === Insert Title Block only if not already present in this layout ===
                        Dim hasTitleBlock As Boolean = False
                        For Each objId As ObjectId In btr1
                            Dim br As BlockReference = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), BlockReference)
                            If br IsNot Nothing Then
                                Dim blkDef As BlockTableRecord = acTrans.GetObject(br.BlockTableRecord, OpenMode.ForRead)
                                If blkDef.Name.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) Then
                                    hasTitleBlock = True
                                    Exit For
                                End If
                            End If
                        Next

                        If Not hasTitleBlock Then
                            Dim insertPoint As New Point3d(0.311, 0.25, 0)
                            Dim blkRef As New BlockReference(insertPoint, blockDefId)
                            blkRef.Layer = "0"
                            btr1.AppendEntity(blkRef)
                            acTrans.AddNewlyCreatedDBObject(blkRef, True)

                            ' Populate attributes for the newly inserted TB
                            Dim blockDef As BlockTableRecord = acTrans.GetObject(blockDefId, OpenMode.ForWrite)

                            For Each entId As ObjectId In blockDef
                                Dim attDef As AttributeDefinition = TryCast(acTrans.GetObject(entId, OpenMode.ForWrite), AttributeDefinition)
                                If attDef IsNot Nothing AndAlso Not attDef.Constant Then
                                    Dim attRef As New AttributeReference()
                                    attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                    If attRef IsNot Nothing Then
                                        If attRef.Tag Like "*INITIAL*" Or attRef.Tag Like "*REV*" Then
                                            If existingInitialRevValues.ContainsKey(attRef.Tag) Then
                                                attRef.TextString = existingInitialRevValues(attRef.Tag)
                                            Else
                                                attRef.TextString = "-"
                                            End If
                                        Else
                                            attRef.TextString = ""
                                        End If
                                    End If
                                    blkRef.AttributeCollection.AppendAttribute(attRef)
                                    acTrans.AddNewlyCreatedDBObject(attRef, True)
                                End If
                            Next
                        Else
                            ' Title block already exists in this layout - update its attributes (INITIAL / REV) if values were collected
                            For Each objId2 As ObjectId In btr1
                                Dim br As BlockReference = TryCast(acTrans.GetObject(objId2, OpenMode.ForWrite), BlockReference)
                                If br Is Nothing Then Continue For
                                Dim defRec As BlockTableRecord = TryCast(acTrans.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                If defRec Is Nothing Then Continue For
                                If defRec.Name.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) Then
                                    For Each attId As ObjectId In br.AttributeCollection
                                        Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                        If attRef Is Nothing Then Continue For
                                        If attRef.Tag Like "*INITIAL*" Or attRef.Tag Like "*REV*" Then
                                            If existingInitialRevValues.ContainsKey(attRef.Tag) Then
                                                attRef.TextString = existingInitialRevValues(attRef.Tag)
                                            ElseIf String.IsNullOrWhiteSpace(attRef.TextString) Then
                                                attRef.TextString = "-"
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next

                    ' Delete layouts not in KeepNames
                    Dim lt As DBDictionary = acTrans.GetObject(acCurdb.LayoutDictionaryId, OpenMode.ForRead)
                    For Each entry As DBDictionaryEntry In lt
                        Dim layout As Layout = acTrans.GetObject(entry.Value, OpenMode.ForRead)
                        If layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase) Then Continue For
                        If Not KeepNames.Contains(layout.LayoutName) Then
                            toDelete.Add(layout.ObjectId)
                        End If
                    Next
                    For Each id In toDelete
                        Dim layout As Layout = acTrans.GetObject(id, OpenMode.ForWrite)
                        layout.Erase()
                    Next

                    acTrans.Commit()

                    ReloadNestedXrefs(acCurdb)

                    lm.CurrentLayout = "Model"

                End Using

            End Using

            Using acDoc.LockDocument()

                Dim folder As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
                Dim name As String = CStr(Application.GetSystemVariable("DWGNAME"))
                Dim path = System.IO.Path.Combine(folder, name)

                acCurdb.SaveAs(path, True, DwgVersion.Current, acCurdb.SecurityParameters)

            End Using

        End Sub

        Public Shared Function LayoutExistsByDictionary(db As Database, layoutName As String) As Boolean
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim layouts As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                If layouts Is Nothing Then Return False
                For Each entry As DBDictionaryEntry In layouts
                    Dim lo As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                    If lo IsNot Nothing AndAlso String.Equals(lo.LayoutName, layoutName, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
                tr.Commit()
            End Using
            Return False
        End Function

        Private Shared Function TryApply11x17PlotConfiguration(ps As PlotSettings, validator As PlotSettingsValidator) As Boolean
            Dim devices As String() = {
                "ARCXIS - DWG To PDF - Brics.pc3",
                "ARCXIS - DWG To PDF.pc3",
                "DWG To PDF.pc3"
            }

            Dim mediaNames As String() = {
                "ANSI_full_bleed_B_(11.00_x_17.00_Inches)",
                "ANSI_full_bleed_B_(17.00_x_11.00_Inches)",
                "ANSI B (11.00 x 17.00 Inches)",
                "ANSI_B_(11.00_x_17.00_Inches)",
                "Tabloid (11 x 17 in)",
                "11x17"
            }

            Dim availableDevices As StringCollection = Nothing
            Try
                availableDevices = validator.GetPlotDeviceList()
            Catch
            End Try

            For Each device In devices
                If availableDevices IsNot Nothing AndAlso
                   Not availableDevices.Cast(Of String)().Any(Function(d) d.Equals(device, StringComparison.OrdinalIgnoreCase)) Then
                    Continue For
                End If

                For Each mediaName In mediaNames
                    Try
                        validator.SetPlotConfigurationName(ps, device, mediaName)
                        Return True
                    Catch
                    End Try
                Next

                Try
                    validator.SetPlotConfigurationName(ps, device, Nothing)
                    Dim canonicalMedia As StringCollection = validator.GetCanonicalMediaNameList(ps)
                    If canonicalMedia IsNot Nothing AndAlso canonicalMedia.Count > 0 Then
                        validator.SetPlotConfigurationName(ps, device, canonicalMedia(0))
                        Return True
                    End If
                Catch
                End Try
            Next

            Return False
        End Function

        Public Shared Function ResolvePlotConfig() As PlotConfig
            Dim candidates As String() = {
                "ARCXIS - DWG To PDF - Brics.pc3",
                "ARCXIS - DWG To PDF.pc3",
                "DWG To PDF.pc3"
            }

            For Each name In candidates
                Try
                    Return PlotConfigManager.SetCurrentConfig(name)
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Public Shared Sub PageSetUp11x17(layoutname As String, Optional ByVal CTBFile As String = "")
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim acLayoutMgr As LayoutManager = LayoutManager.Current
            acLayoutMgr.CurrentLayout = layoutname

            If CTBFile = "" Or CTBFile = "ARCXIS" Then
                CTBFile = "ARCXIS.ctb"
            ElseIf CTBFile = "DPIS" Then
                CTBFile = "DPIS-11x17.ctb"
            ElseIf CTBFile = "PTS" Then
                CTBFile = "PTS.ctb"
            ElseIf CTBFile = "ARCXIS - Mechanical" Then
                CTBFile = "ARCXIS - Mechanical.ctb"
            End If

            Using acLckDoc As DocumentLock = acDoc.LockDocument()
                ' TRANSACTION 1: Create the named plot settings if it doesn't exist
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim plSets As DBDictionary = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead)
                    Dim layoutId As ObjectId = acLayoutMgr.GetLayoutId(layoutname)
                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                    If Not plSets.Contains("Arcxis") Then
                        Dim acPlSet As New PlotSettings(False) ' False = PaperSpace
                        acPlSet.CopyFrom(acLayout)
                        acPlSet.PlotSettingsName = "Arcxis"
                        plSets.UpgradeOpen()
                        plSets.SetAt("Arcxis", acPlSet)
                        acTrans.AddNewlyCreatedDBObject(acPlSet, True)
                    End If

                    acTrans.Commit()
                End Using

                ' TRANSACTION 2: Now configure the plot settings
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim plSets As DBDictionary = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead)
                    Dim layoutId As ObjectId = acLayoutMgr.GetLayoutId(layoutname)
                    Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                    Dim acPlSet As PlotSettings = CType(acTrans.GetObject(plSets.GetAt("Arcxis"), OpenMode.ForWrite), PlotSettings)

                    Try
                        Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current
                        If Not TryApply11x17PlotConfiguration(acPlSet, acPlSetVdr) Then
                            Throw New InvalidOperationException("Could not find a valid PC3/paper size combination for 11x17 plotting.")
                        End If
                        acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)
                        acPlSetVdr.SetPlotType(acPlSet, Teigha.DatabaseServices.PlotType.Extents)
                        Dim lowerLeft As New Point2d(0, 0)
                        Dim upperRight As New Point2d(17, 11)
                        Dim extents As New Extents2d(lowerLeft, upperRight)
                        acPlSetVdr.SetPlotWindowArea(acPlSet, extents)
                        acPlSetVdr.SetPlotOrigin(acPlSet, New Point2d(0, 0))
                        acPlSetVdr.SetPlotCentered(acPlSet, True)
                        acPlSetVdr.SetUseStandardScale(acPlSet, True)
                        acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit)
                        acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches)
                        acPlSet.ScaleLineweights = True
                        acPlSet.ShowPlotStyles = False
                        acPlSetVdr.RefreshLists(acPlSet)
                        acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed
                        acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal
                        acPlSet.PrintLineweights = True
                        acPlSet.PlotTransparency = False
                        acPlSet.PlotPlotStyles = True
                        acPlSet.DrawViewportsFirst = False
                        acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees270)
                        Try
                            acPlSetVdr.SetCurrentStyleSheet(acPlSet, CTBFile)
                        Catch
                            acPlSetVdr.SetCurrentStyleSheet(acPlSet, "ARCXIS.ctb")
                        End Try

                        ' Copy all settings to the layout
                        acLayout.CopyFrom(acPlSet)
                    Catch es As Exception
                        MsgBox(es.Message & vbCrLf & "Error setting plot configuration. Check PC3 and paper size name.")
                    End Try

                    acTrans.Commit()
                End Using
            End Using
        End Sub

        Private Shared Sub ReloadXrefsInBatches(db As Database, xrefIds As List(Of ObjectId), Optional batchSize As Integer = 20)
            If db Is Nothing OrElse xrefIds Is Nothing OrElse xrefIds.Count = 0 Then Return
            If batchSize < 1 Then batchSize = 20

            Dim i As Integer = 0
            While i < xrefIds.Count
                Dim chunk As New ObjectIdCollection()
                Dim jLimit As Integer = Math.Min(i + batchSize, xrefIds.Count)
                Dim j As Integer = i
                While j < jLimit
                    chunk.Add(xrefIds(j))
                    j += 1
                End While

                Try
                    db.ReloadXrefs(chunk)
                Catch
                    ' Ignore reload errors and continue with next chunk.
                End Try

                i = jLimit
            End While
        End Sub

        Public Shared Sub ReloadNestedXrefs(db As Database, tr As Transaction)
            If db Is Nothing Then Return

            ' ReloadXrefs is unsafe with active transactions in some CAD runtimes.
            ' Keep this signature for compatibility, but do not reload while transaction is active.
            If tr IsNot Nothing Then Return

            ReloadNestedXrefs(db)
        End Sub

        Public Shared Sub ReloadNestedXrefs(db As Database)
            If db Is Nothing Then Return

            Dim xrefIds As New List(Of ObjectId)()
            Using tr As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)

                For Each id As ObjectId In bt
                    Dim btr As BlockTableRecord = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                    If btr Is Nothing Then Continue For

                    If btr.IsFromExternalReference Then
                        If btr.Name = "Master Seal File" OrElse btr.XrefStatus = XrefStatus.Unresolved Then
                            xrefIds.Add(btr.ObjectId)
                        End If
                    End If
                Next
            End Using

            ReloadXrefsInBatches(db, xrefIds)
        End Sub

        Public Shared Sub ReloadNestedXrefsAuto()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            If acDoc Is Nothing Then Return
            Dim acCurDb As Database = acDoc.Database
            Using acDoc.LockDocument()
                ReloadNestedXrefs(acCurDb)
            End Using

        End Sub

        Public Shared Sub DetachXrefs(db As Database, tr As Transaction, xref As String)

            ' Also handle raster image definitions that reference the given xref string
            Try
                Dim imgDictId As ObjectId = RasterImageDef.GetImageDictionary(db)
                If Not imgDictId.IsNull Then
                    Dim imgDict As DBDictionary = tr.GetObject(imgDictId, OpenMode.ForRead)

                    ' Collect keys to remove to avoid modifying the dictionary while iterating
                    Dim keysToRemove As New List(Of String)()
                    For Each entry As DBDictionaryEntry In imgDict
                        Dim key As String = entry.Key
                        Dim defId As ObjectId = entry.Value
                        Dim defObj As RasterImageDef = TryCast(tr.GetObject(defId, OpenMode.ForRead), RasterImageDef)
                        If defObj IsNot Nothing Then
                            Dim src As String = If(defObj.SourceFileName, String.Empty)
                            If String.Equals(key, xref, StringComparison.OrdinalIgnoreCase) OrElse
                               (Not String.IsNullOrWhiteSpace(src) AndAlso src.IndexOf(xref, StringComparison.OrdinalIgnoreCase) >= 0) Then
                                keysToRemove.Add(key)
                            End If
                        End If
                    Next

                    If keysToRemove.Count > 0 Then
                        ' We'll need the block table to search for RasterImage entities
                        Dim btRead As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)

                        imgDict.UpgradeOpen()
                        For Each key In keysToRemove
                            If Not imgDict.Contains(key) Then Continue For
                            Dim defId As ObjectId = imgDict.GetAt(key)

                            ' Erase any RasterImage entities that reference this definition
                            For Each btrId As ObjectId In btRead
                                Dim btrRec As BlockTableRecord = TryCast(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                                If btrRec Is Nothing Then Continue For
                                For Each objId As ObjectId In btrRec
                                    If objId.ObjectClass.DxfName = "RASTERIMAGE" Then
                                        Dim ri As RasterImage = TryCast(tr.GetObject(objId, OpenMode.ForWrite), RasterImage)
                                        If ri IsNot Nothing AndAlso ri.ImageDefId = defId Then
                                            ri.Erase()
                                        End If
                                    End If
                                Next
                            Next

                            ' Erase the RasterImageDef itself and remove from dictionary
                            If Not defId.IsNull Then
                                Dim def As DBObject = tr.GetObject(defId, OpenMode.ForWrite)
                                If def IsNot Nothing Then
                                    def.Erase()
                                End If
                            End If

                            Try
                                imgDict.Remove(key)
                            Catch
                                ' ignore remove errors
                            End Try
                        Next
                        imgDict.DowngradeOpen()
                    End If
                End If
            Catch
                ' Best-effort: ignore any failures while cleaning raster defs
            End Try
        End Sub

        Private Function GetCurrentDwgFolder() As String
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            If Not String.IsNullOrWhiteSpace(db.Filename) Then
                Return Path.GetDirectoryName(db.Filename)
            End If
            ' Unsaved drawing fallback
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

        Private NotInheritable Class WindowWrapper
            Implements IWin32Window

            Private ReadOnly _handle As IntPtr

            Public Sub New(handle As IntPtr)
                _handle = handle
            End Sub

            Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
                Get
                    Return _handle
                End Get
            End Property
        End Class

        ' ===== Helpers you can put anywhere in your module =====

        ' Snapshot of a few system variables you change
        Private Structure SysVarSnapshot
            Public BackGroundPlot As Object
            Public CMDDIA As Object
            Public FILEDIA As Object
        End Structure

        Private Function SnapSysVars() As SysVarSnapshot
            Return New SysVarSnapshot With {
        .BackGroundPlot = Application.GetSystemVariable("BackGroundPlot"),
        .CMDDIA = Application.GetSystemVariable("CMDDIA"),
        .FILEDIA = Application.GetSystemVariable("FILEDIA")
    }
        End Function

        Private Sub RestoreSysVars(s As SysVarSnapshot)
            Try
                Application.SetSystemVariable("BackGroundPlot", s.BackGroundPlot)
                Application.SetSystemVariable("CMDDIA", s.CMDDIA)
                Application.SetSystemVariable("FILEDIA", s.FILEDIA)
            Catch
                ' swallow – we’re restoring best-effort
            End Try
        End Sub

        ' Track layer flags you change so you can restore exactly
        Private Class LayerFlags
            Public Name As String
            Public IsOff As Boolean
            Public IsFrozen As Boolean
            Public IsPlottable As Boolean
        End Class


        ' Set or update a custom DWG property

        Public Shared Sub SetCustomDwgProp(propName As String, propValue As String)
            Dim db = Application.DocumentManager.MdiActiveDocument.Database
            Dim b As New DatabaseSummaryInfoBuilder(db.SummaryInfo)

            Dim tbl = TryCast(b.CustomPropertyTable, System.Collections.IDictionary)
            If tbl Is Nothing Then Throw New InvalidOperationException("CustomPropertyTable is not editable on this version.")
            If tbl.Contains(propName) Then
                tbl(propName) = propValue
            Else
                tbl.Add(propName, propValue)
            End If
            db.SummaryInfo = b.ToDatabaseSummaryInfo()
        End Sub



        ' <CommandMethod("TransformDpis")>



        <CommandMethod("overnightprinting")>
        Sub OverNightPrinting()

            Dim acDoc = Application.DocumentManager.MdiActiveDocument
            Dim accurdb = acDoc.Database
            Dim content As New System.Text.StringBuilder()

            Using acLckDoc As DocumentLock = acDoc.LockDocument()
                Using acTrans As Transaction = accurdb.TransactionManager.StartTransaction()
                    Dim PropertiesBuilder = GetCustomDwgPropForDoc(accurdb, "BUILDER")
                    Dim PropertiesPlan = GetCustomDwgPropForDoc(accurdb, "PLAN")
                    Dim PropertiesPlanType = If(GetCustomDwgPropReliable("PLAN TYPE"), "").Trim()

                    Dim missingProps As New List(Of String)()
                    If String.IsNullOrWhiteSpace(PropertiesBuilder) Then missingProps.Add("BUILDER")
                    If String.IsNullOrWhiteSpace(PropertiesPlan) Then missingProps.Add("PLAN")
                    If String.IsNullOrWhiteSpace(PropertiesPlanType) Then missingProps.Add("PLAN TYPE")

                    If missingProps.Count > 0 Then
                        content.AppendLine("=== MISSING DRAWING PROPERTIES ===")
                        content.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                        content.AppendLine($"DWG File: {Path.GetFileName(accurdb.Filename)}")
                        content.AppendLine()
                        content.AppendLine("The following required drawing properties are missing or empty:")
                        content.AppendLine()
                        For Each prop In missingProps
                            content.AppendLine($"  - {prop}")
                        Next
                        content.AppendLine()
                        content.AppendLine("Please set these properties before submitting the drawing.")
                        CreateTextFileInDwgLocation("ReasonForFailure.log", content.ToString())
                        Exit Sub
                    End If

                    If String.Equals(PropertiesPlanType, "MECHANICAL", StringComparison.OrdinalIgnoreCase) Then
                        MechanicalHeadlessPrinting.PrintMechanicalPlans()
                    Else
                        content.AppendLine("=== UNSUPPORTED PLAN TYPE ===")
                        content.AppendLine($"PLAN TYPE must be MECHANICAL. Found: {PropertiesPlanType}")
                        CreateTextFileInDwgLocation("ReasonForFailure.log", content.ToString())
                    End If
                End Using
            End Using

        End Sub

        Public Shared Sub QueueLayoutsForCsv(layoutList As List(Of List(Of String)), pdfName As String, builder As String, planName As String, projectnumber As String, Optional ByVal IRC As String = "", Optional ByVal IECC As String = "", Optional ByVal MechCounty As String = "",
                                             Optional ByVal MechMan As String = "", Optional ByVal MechFuel As String = "", Optional ByVal MechPlan As Boolean = False, Optional ByVal DocType As String = "", Optional ByVal BuilderDivision As String = "")
            If String.IsNullOrWhiteSpace(pdfName) Then Exit Sub
            If layoutList Is Nothing OrElse layoutList.Count = 0 Then Exit Sub

            ' Build identifier ("Builder - PdfName")
            Dim pdfKey As String = If(pdfName, "").Trim()
            If pdfKey.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                pdfKey = pdfKey.Substring(0, pdfKey.Length - 4)
            End If

            Dim identifier As String = $"{If(builder, "").Trim()} - {pdfKey}".Trim()

            ' Prevent duplicate queuing for same identifier (already queued)
            Dim alreadyQueued As Boolean = _pendingRows.Any(Function(r) r.Count >= 1 AndAlso String.Equals(r(0), identifier, StringComparison.OrdinalIgnoreCase))
            If alreadyQueued Then Exit Sub

            ' Extract meta from first layout row
            Dim first = layoutList(0)
            Dim modName As String = ""
            Dim swingRaw As String = ""
            Dim elevation As String = ""
            Dim venttype As String = ""
            Dim attictype As String = ""
            Dim Community As String = ""
            Dim ClimateZone As String = ""
            Dim IeccValue As String = ""
            Dim InsulationValue As String = ""
            Dim FuelValue As String = ""
            Dim VersionValue As String = ""
            ' NEW: collect all layoutList(6) values, unique, comma-separated
            Dim planOption As String = String.Join(", ", layoutList.Where(Function(r) r IsNot Nothing AndAlso r.Count > 6).Select(Function(r) If(r(6), "").Trim()).
            Where(Function(v) v <> "").Distinct(StringComparer.OrdinalIgnoreCase))

            If first.Count > 2 Then modName = If(first(2), "").Trim()
            If first.Count > 3 Then swingRaw = If(first(3), "").Trim()
            If first.Count > 8 AndAlso Not String.IsNullOrWhiteSpace(first(8)) Then
                elevation = first(8).Trim()
            ElseIf first.Count > 7 Then
                elevation = If(first(7), "").Trim()
            End If

            If String.Equals(modName, "ATTIC VENT", StringComparison.OrdinalIgnoreCase) Then
                venttype = If(first(9), "").Trim()
                attictype = If(first(10), "").Trim()
                ' Optional: don't overwrite the aggregated planOption from all rows
                ' If first.Count > 6 Then planOption = If(first(6), "").Trim()
            End If

            If String.Equals(DocType, "ENERGY", StringComparison.OrdinalIgnoreCase) Then
                If first.Count > 9 Then Community = If(first(9), "").Trim()
                If first.Count > 10 Then ClimateZone = If(first(10), "").Trim()
                If first.Count > 11 Then IeccValue = If(first(11), "").Trim()
                If first.Count > 12 Then InsulationValue = If(first(12), "").Trim()
                If first.Count > 13 Then FuelValue = If(first(13), "").Trim()
                If first.Count > 14 Then VersionValue = If(first(14), "").Trim()
            End If


            swingRaw = NormalizeSwing(swingRaw)

            ' Store four logical rows: PLAN / ELEV / SW / MOD
            _pendingRows.Add(New List(Of String) From {identifier, "PLAN", planName})
            _pendingRows.Add(New List(Of String) From {identifier, "ELEVATION", elevation})
            _pendingRows.Add(New List(Of String) From {identifier, "SWING", swingRaw})
            If BuilderDivision <> "" Then
                _pendingRows.Add(New List(Of String) From {identifier, "BUILDER DIVISION", BuilderDivision})
            End If
            If modName.Equals("BRACING", StringComparison.OrdinalIgnoreCase) Or modName.Equals("WINDSTORM", StringComparison.OrdinalIgnoreCase) Then
                If Not String.IsNullOrWhiteSpace(planOption) Then
                    Dim planOptionFirstPart As String = If(String.IsNullOrWhiteSpace(planOption), "", planOption.Split(" "c)(0))
                    _pendingRows.Add(New List(Of String) From {identifier, "WINDSPEED", planOptionFirstPart})
                End If
            End If
            If MechPlan Then
                _pendingRows.Add(New List(Of String) From {identifier, "PLAN TYPE", "MECHANICAL"})
                _pendingRows.Add(New List(Of String) From {identifier, "COUNTY", MechCounty})
                _pendingRows.Add(New List(Of String) From {identifier, "MANUFACTURER", MechMan})
                _pendingRows.Add(New List(Of String) From {identifier, "FUEL TYPE", MechFuel})
            Else
                _pendingRows.Add(New List(Of String) From {identifier, "PLAN TYPE", modName})
            End If
            If modName.Equals("ATTIC VENT", StringComparison.OrdinalIgnoreCase) Then
                _pendingRows.Add(New List(Of String) From {identifier, "VENT TYPE", venttype})
                _pendingRows.Add(New List(Of String) From {identifier, "ATTIC TYPE", attictype})
                If Not String.IsNullOrWhiteSpace(planOption) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "OPTION", planOption})
                End If
            End If
            If DocType.Equals("ENERGY", StringComparison.OrdinalIgnoreCase) Then
                _pendingRows.Add(New List(Of String) From {identifier, "PLAN TYPE", "ENERGY"})
                If Not String.IsNullOrWhiteSpace(Community) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "COMMUNITY", Community})
                End If
                If Not String.IsNullOrWhiteSpace(ClimateZone) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "CLIMATE ZONE", ClimateZone})
                End If
                If Not String.IsNullOrWhiteSpace(IeccValue) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "IECC", IeccValue})
                End If
                If Not String.IsNullOrWhiteSpace(InsulationValue) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "INSULATION TYPE", InsulationValue})
                End If
                If Not String.IsNullOrWhiteSpace(FuelValue) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "FUEL TYPE", FuelValue})
                End If
                If Not String.IsNullOrWhiteSpace(VersionValue) Then
                    _pendingRows.Add(New List(Of String) From {identifier, "PLAN VERSION", VersionValue})
                End If
            End If
        End Sub

        Public Shared Function FlushQueuedCsv(Optional builder As String = Nothing, Optional planName As String = Nothing, Optional FileName As String = "", Optional plantype As String = "") As String
            If _pendingRows Is Nothing OrElse _pendingRows.Count = 0 Then Return Nothing

            Dim pendingDir As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\fs2\k\DPIS Drawings\PDF File Data\Pending"
            If Not Directory.Exists(pendingDir) Then Directory.CreateDirectory(pendingDir)

            ' Attempt to derive builder/plan from first PLAN row if parameters not supplied.
            If String.IsNullOrWhiteSpace(builder) OrElse String.IsNullOrWhiteSpace(planName) Then
                Dim planRow = _pendingRows.FirstOrDefault(Function(r) r.Count >= 3 AndAlso r(1) = "PLAN")
                If planRow IsNot Nothing Then
                    If String.IsNullOrWhiteSpace(builder) Then
                        ' Builder is the part before " - " in Identifier
                        Dim ident = planRow(0)
                        Dim dashIdx = ident.IndexOf(" - ", StringComparison.Ordinal)
                        If dashIdx >= 0 Then builder = ident.Substring(0, dashIdx).Trim()
                    End If
                    If String.IsNullOrWhiteSpace(planName) Then planName = planRow(2)
                End If
            End If

            Dim fileBase As String = $"{SafeSegment(If(builder, ""))}_{SafeSegment(If(planName, ""))}" & If(Not String.IsNullOrWhiteSpace(FileName), "_" & SafeSegment(FileName), "") & If(Not String.IsNullOrWhiteSpace(plantype), "_" & SafeSegment(plantype), "")

            If String.IsNullOrWhiteSpace(fileBase.Replace("_", "")) Then fileBase = "Combined"

            Dim csvPath As String = Path.Combine(pendingDir, fileBase & ".csv")
            Dim counter As Integer = 1
            While File.Exists(csvPath)
                csvPath = Path.Combine(pendingDir, $"{fileBase}_{counter}.csv")
                counter += 1
            End While

            Using sw As New StreamWriter(csvPath, False, System.Text.Encoding.UTF8)
                sw.WriteLine("Identifier,Key,Value")
                For Each row In _pendingRows
                    ' Row integrity: expect exactly 3 columns now
                    Dim ident As String = If(row.ElementAtOrDefault(0), "")
                    Dim key As String = If(row.ElementAtOrDefault(1), "")
                    Dim valueStr As String = If(row.ElementAtOrDefault(2), "")
                    sw.WriteLine($"{CsvQuote(ident)},{CsvQuote(key)},{CsvQuote(valueStr)}")
                Next
            End Using

            _pendingRows.Clear()
            Return csvPath

        End Function

        Private Shared Function SafeSegment(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Dim cleaned As String = New String(s.Trim().Select(Function(ch) If(Char.IsLetterOrDigit(ch) Or ch = "-"c Or ch = "_"c, ch, "_"c)).ToArray())
            While cleaned.Contains("__")
                cleaned = cleaned.Replace("__", "_")
            End While
            Return cleaned.Trim("_"c)
        End Function

        Private Shared Function CsvQuote(s As String) As String
            Return """" & s.Replace("""", """""") & """"
        End Function



        ' --- PDF publish tracking and combine helpers ---

        ''' <summary>
        ''' Wait up to timeoutMs for the file to exist. Best-effort helper.
        ''' </summary>

        ''' <summary>
        ''' Register a published PDF path (duplicates ignored).
        ''' Call this after publish succeeds.
        ''' </summary>

        ' --- Updated combine + watermark support ---
        ' Replace the existing CombineRegisteredPdfs, CombinePdfsCommand and AutoCombinePdfsCommand with these.

        ' --- Per-PDF metadata captured at publish time (no string parsing) ---

        ' Simple combiner for an explicit list (does not touch the global registry)

        ' Build a combined FR+WB list from the PDFs published for the current stamp (preserving publish order)

        ' Normalize swing values to LEFT/RIGHT
        Private Shared Function NormalizeSwing(s As String) As String
            Dim t As String = If(s, "").Trim().ToUpperInvariant()
            If t = "L" OrElse t = "LEFT" Then Return "LEFT"
            If t = "R" OrElse t = "RIGHT" Then Return "RIGHT"
            Return t
        End Function

        ' Overload that lets callers customize filename and (optionally) subfolder.


        ' === Added: Excel driven layer import & merge ===

        ' Parse color specifications: ACI number, name, or R,G,B

        Public Shared Sub EnsureLayoutCount(requiredCount As Integer, Optional customCTB As String = "")
            If requiredCount <= 0 Then Return

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            If acDoc Is Nothing Then Return

            Dim db As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor

            Try
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    ' Count existing paperspace layouts (excluding Model)
                    Dim layoutDict As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                    If layoutDict Is Nothing Then
                        tr.Commit()
                        Return
                    End If

                    Dim existingLayoutCount As Integer = 0
                    For Each entry As DBDictionaryEntry In layoutDict
                        Dim layout As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                        If layout IsNot Nothing AndAlso Not layout.ModelType Then
                            existingLayoutCount += 1
                        End If
                    Next

                    tr.Commit()

                    ' If counts don't match, create layouts
                    If requiredCount > existingLayoutCount Then

                        ' Determine CTB file if not provided
                        If String.IsNullOrEmpty(customCTB) Then
                            Dim planType As String = GetCustomDwgPropReliable("PLAN TYPE")
                            If String.Equals(planType, "Mechanical", StringComparison.OrdinalIgnoreCase) Then
                                customCTB = "ARCXIS - Mechanical"
                            Else
                                customCTB = "ARCXIS"
                            End If
                        End If

                        ' Create the required number of layouts
                        CreateLayoutsWithTitleblock(requiredCount, True, customCTB)

                    End If
                End Using

            Catch ex As Exception
                ed.WriteMessage(vbLf & $"Error ensuring layout count: {ex.Message}")
            End Try
        End Sub

        Public Shared Function CreateTextFileInDwgLocation(Optional fileName As String = "", Optional content As String = "", Optional append As Boolean = False) As String
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return Nothing

            Dim db As Database = doc.Database

            ' Check if the DWG is saved
            If String.IsNullOrWhiteSpace(db.Filename) Then
                MessageBox.Show("The current drawing has not been saved yet. Please save the drawing first.",
                               "Unsaved Drawing", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If

            Try
                ' Get the directory where the DWG is located
                Dim dwgFolder As String = Path.GetDirectoryName(db.Filename)

                ' If no filename provided, use DWG name with .txt extension
                If String.IsNullOrWhiteSpace(fileName) Then
                    Dim dwgName As String = Path.GetFileNameWithoutExtension(db.Filename)
                    fileName = dwgName & ".log"
                ElseIf Not fileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase) Then
                    ' Ensure .log extension
                    fileName &= ".log"
                End If

                ' Combine to get full path
                Dim txtFilePath As String = Path.Combine(dwgFolder, fileName)

                ' Write the content to the file
                If append Then
                    File.AppendAllText(txtFilePath, content & Environment.NewLine, System.Text.Encoding.UTF8)
                Else
                    File.WriteAllText(txtFilePath, content, System.Text.Encoding.UTF8)
                End If

                Return txtFilePath

            Catch ex As Exception
                MessageBox.Show($"Error creating text file: {ex.Message}",
                               "File Creation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function

        ' Map mm (or string) to closest LineWeight enum
        ' === End Added ===
    End Class
End Namespace