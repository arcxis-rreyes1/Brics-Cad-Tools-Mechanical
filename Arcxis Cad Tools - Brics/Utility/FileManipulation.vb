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
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports DocumentFormat.OpenXml.Drawing.Diagrams
Imports DocumentFormat.OpenXml.Office2010.Drawing
Imports DocumentFormat.OpenXml.Office2010.Excel
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop
Imports Microsoft.SqlServer.Server
Imports PdfSharp.Drawing
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports Path = System.IO.Path
Imports Bricscad.ApplicationServices
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document
Imports Exception = Teigha.Runtime.Exception
Imports Layout = Teigha.DatabaseServices.Layout
Imports Color = Teigha.Colors.Color
Imports System.Collections.Specialized
Imports Arcxis_Cad_Tools.Utility.PageBlockDivisionCommands


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
        ' Temporary: installable first-chance hook to break in VS when Teigha throws
        <CommandMethod("InstallFirstChanceHook")>
        Public Sub InstallFirstChanceHook()
            Try
                AddHandler AppDomain.CurrentDomain.FirstChanceException, AddressOf OnFirstChance
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "FirstChance hook installed")
            Catch ex As System.Exception
                ' avoid throwing during install
            End Try
        End Sub

        Private Shared _lastTeigha As Teigha.Runtime.Exception

        <CommandMethod("RemoveFirstChanceHook")>
        Public Sub RemoveFirstChanceHook()
            Try
                RemoveHandler AppDomain.CurrentDomain.FirstChanceException, AddressOf OnFirstChance
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "FirstChance hook removed")
            Catch
            End Try
        End Sub

        Private Sub OnFirstChance(sender As Object, e As System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs)
            Try
                Dim odEx = TryCast(e.Exception, Teigha.Runtime.Exception)
                If odEx Is Nothing Then Return
                _lastTeigha = odEx ' inspect this anytime in debugger
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
            vbLf & $"Teigha thrown: {odEx.ErrorStatus} - {odEx.Message}")
                Debugger.Break()
            Catch
            End Try
        End Sub


        <CommandMethod("SDFT")>
        Sub SetupDrawings()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database
            Dim SealLoop As New List(Of String)

            Dim StampLayers As New List(Of String)
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim blockName As String = "Arcxis Title Block"
                Dim sourceDwgPath As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\Arcxis\Engineering\Drafting Standards\CAD Blocks\Arcxis Title Block - Block.dwg"
                BlockImport(acCurDb, blockName, sourceDwgPath, acTrans)
                'ReloadNestedXrefs(acDb, acTrans)

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
                frm.SealsList.Items.Add(seal)
            Next

            Dim Builder = GetCustomDwgPropReliable("BUILDER")
            Dim plan = GetCustomDwgPropReliable("PLAN")
            Dim Stamps = GetCustomDwgPropReliable("STAMPS")

            If Builder <> "" Then
                frm.TextBox1.Text = Builder
            End If

            If plan <> "" Then
                frm.TextBox2.Text = plan
            End If

            If Not String.IsNullOrWhiteSpace(Stamps) Then
                Dim tokens = Stamps.Split(","c).
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

            frm.ComboBox1.SelectedIndex = 0 'Default to ARCXIS

            frm.ShowDialog()

            CleanUpLayoutsAndBlocks()

        End Sub

        <CommandMethod("RemoveXrefs")>
        Public Shared Sub RemoveSealXrefs()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using acLck As DocumentLock = doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Try
                        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)

                        Dim removed As Integer = 0

                        For Each id As ObjectId In bt
                            Dim btr As BlockTableRecord = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                            If btr Is Nothing Then Continue For
                            If Not btr.IsFromExternalReference Then Continue For

                            Dim path As String = String.Empty
                            Try
                                ' BlockTableRecord.PathName exists in AutoCAD .NET API and holds the external reference path
                                path = btr.PathName
                            Catch
                                ' ignore if property unavailable
                            End Try

                            If Not String.IsNullOrWhiteSpace(path) AndAlso path.IndexOf("seal$", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                Try
                                    db.DetachXref(btr.ObjectId)
                                    removed += 1
                                Catch ex As Exception
                                    ed.WriteMessage(vbLf & "Failed to detach xref: " & path & " - " & ex.Message)
                                End Try
                            End If
                        Next

                        ' Also remove raster images that reference files under a \seal$ share
                        Try
                            Dim imgDictId As ObjectId = RasterImageDef.GetImageDictionary(db)
                            If Not imgDictId.IsNull Then
                                Dim imgDict As DBDictionary = tr.GetObject(imgDictId, OpenMode.ForRead)
                                Dim keysToRemove As New List(Of String)()
                                For Each entry As DBDictionaryEntry In imgDict
                                    Dim defId As ObjectId = entry.Value
                                    Dim defObj As RasterImageDef = TryCast(tr.GetObject(defId, OpenMode.ForRead), RasterImageDef)
                                    If defObj IsNot Nothing Then
                                        Dim src As String = If(defObj.SourceFileName, String.Empty)
                                        If Not String.IsNullOrWhiteSpace(src) AndAlso src.IndexOf("seal$", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                            keysToRemove.Add(entry.Key)
                                        End If
                                    End If
                                Next

                                If keysToRemove.Count > 0 Then
                                    Dim btRead As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                                    imgDict.UpgradeOpen()
                                    For Each key In keysToRemove
                                        If Not imgDict.Contains(key) Then Continue For
                                        Dim defId As ObjectId = imgDict.GetAt(key)
                                        ' erase referencing RasterImage entities
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

                                        ' erase the definition
                                        If Not defId.IsNull Then
                                            Dim defObj As DBObject = tr.GetObject(defId, OpenMode.ForWrite)
                                            If defObj IsNot Nothing Then defObj.Erase()
                                        End If

                                        Try
                                            imgDict.Remove(key)
                                        Catch
                                        End Try
                                    Next
                                    imgDict.DowngradeOpen()
                                End If
                            End If
                        Catch
                            ' ignore image cleanup errors
                        End Try

                        tr.Commit()
                        ed.WriteMessage(vbLf & $"RemoveSealXrefs: detached {removed} xref(s) and cleaned matching raster images.")
                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "RemoveSealXrefs error: " & ex.Message)
                    End Try
                End Using
            End Using
        End Sub

        Private Shared Sub DeleteStrayDsdFiles(rootFolder As String)
            If String.IsNullOrWhiteSpace(rootFolder) OrElse Not Directory.Exists(rootFolder) Then Exit Sub

            Try
                ' Delete all .dsd files in this folder
                For Each dsdFile In Directory.GetFiles(rootFolder, "*.dsd", SearchOption.TopDirectoryOnly)
                    Try
                        File.Delete(dsdFile)
                    Catch ex As Exception
                        ' Optionally log or ignore
                    End Try
                Next

                ' Recurse into subfolders
                For Each subDir In Directory.GetDirectories(rootFolder)
                    DeleteStrayDsdFiles(subDir)
                Next
            Catch ex As Exception
                ' Optionally log or ignore
            End Try
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

        <CommandMethod("SCP")>
        Sub CallDrawingProps()

            Dim frm As New Form_FramingProperties

            frm.ShowDialog()

        End Sub


        <CommandMethod("AVS")>
        Sub CallAtticVentProps()
            Dim frm As New Form_AtticVentPrinting
            frm.ShowDialog()
        End Sub

        <CommandMethod("AMP")>
        Sub CallMechanicalProps()
            Dim frm As New Form_MechanicalPrinting
            frm.ShowDialog()
        End Sub

        <CommandMethod("AFP")>
        Sub AutoFramingPrint() '(ByVal BUilder As String)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database
            Dim PlanType As New List(Of String)
            Dim Swings As New List(Of String)
            Dim Elevations As New List(Of String)
            Dim Options As New List(Of String)
            Dim AllValues As New List(Of List(Of String))
            Dim insertionPoint2 As Point3d
            Dim pdfname As String = ""
            Dim planname As String = ""
            Dim CurrentLayerID As ObjectId = acCurDb.Clayer
            Dim Builder As String = ""
            Dim ProjectNumber As String = ""
            Dim frm As New Form_Automated_Printing_Info
            Dim CustomPrinting As Boolean = False
            Dim DPISCTB As Boolean = False

            Dim SealLoop As New List(Of String)

            Dim StampLayers As New List(Of String)

            ReloadNestedXrefsAuto()


            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim lytab As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim cleanlayerstring As String

                'ReloadNestedXrefs(acCurDb, acTrans)

                For Each layer In lytab

                    Dim lytr As LayerTableRecord = acTrans.GetObject(layer, OpenMode.ForRead)

                    If lytr.Name Like "Master Seal File|S-SEAL-*" Then

                        cleanlayerstring = lytr.Name.Substring(24, lytr.Name.Length - 24)

                        StampLayers.Add(cleanlayerstring)

                    End If

                Next

            End Using

            PageBlockDivisionCommands.EnsurePageDivisions()

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
                frm.SealsList.Items.Add(seal)
            Next

            SetCustomDwgPropReliable("PLAN TYPE", "FRAMING")

            Dim FormBuilder = GetCustomDwgPropReliable("BUILDER")
            Dim Formplan = GetCustomDwgPropReliable("PLAN")
            Dim FormStamps = GetCustomDwgPropReliable("STAMPS")
            Dim FormProject = GetCustomDwgPropReliable("PROJECT NUMBER")
            Dim FormSheetLabels = GetCustomDwgPropReliable("SHEET LABELING")
            Dim packageFRWB As String = GetCustomDwgPropReliable("PackageFRWB")

            Dim pageDivisions As List(Of String) =
    PageBlockDivisionCommands.GetPageBlockDivisions().
        Select(Function(s) If(s, "").Trim()).
        Where(Function(s) s <> "").
        Distinct(StringComparer.OrdinalIgnoreCase).
        ToList()

            ' CheckedListBox5 = exactly PAGE divisions
            PageBlockDivisionCommands.SyncCheckedListBoxExact(frm.CheckedListBox2, pageDivisions)

            Dim divisionsValue As String = String.Join(
        ", ",
        pageDivisions.
        Select(Function(s) If(s, "").Trim()).
        Where(Function(s) s <> "").
        Distinct(StringComparer.OrdinalIgnoreCase)
)

            SetCustomDwgPropReliable("DIVISIONS", divisionsValue)

            Dim doWSFW As Boolean = String.Equals(packageFRWB, "WSFW", StringComparison.OrdinalIgnoreCase)
            Dim doRFR As Boolean = String.Equals(packageFRWB, "RFR", StringComparison.OrdinalIgnoreCase)

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


            If FormSheetLabels <> "" Then
                If FormSheetLabels = "S" Then
                    frm.SheetLabels.SelectedIndex = 0
                ElseIf FormSheetLabels = "FR" Then
                    frm.SheetLabels.SelectedIndex = 1
                End If
            Else
                frm.SheetLabels.BackColor = System.Drawing.Color.Red
                frm.Button1.Visible = False
            End If

            frm.TextBox1.Text = FormBuilder
            frm.TextBox2.Text = Formplan
            frm.ShowDialog()

            Dim StickAndTruss As Boolean = False

            If frm.DialogResult = DialogResult.Cancel Then Exit Sub

            Builder = UCase(frm.TextBox1.Text)

            Dim FRsheets As Boolean = False

            If frm.SheetLabels.SelectedItem = "FR" Then
                FRsheets = True
            End If

            If Builder = "MERITAGE HOMES" Then

                Dim result2 As DialogResult = MessageBox.Show("Loop Trust and Stick Framing?", "caption", MessageBoxButtons.YesNoCancel)
                If result2 = DialogResult.Cancel Then
                    StickAndTruss = False
                ElseIf result2 = DialogResult.No Then
                    StickAndTruss = False
                ElseIf result2 = DialogResult.Yes Then
                    StickAndTruss = True
                End If

            ElseIf Builder = "LENNAR HOMES" Then
                DPISCTB = True
            End If

            Dim LayerLoop As New List(Of String)

            If StickAndTruss Then

                LayerLoop.Add("S-FRM-STICK")
                LayerLoop.Add("S-FRM-TRUSS")

            Else

                LayerLoop.Add("")

            End If

            planname = UCase(frm.TextBox2.Text)

            If Builder <> FormBuilder Then

                SetCustomDwgPropReliable("BUILDER", Builder)

            End If

            If planname <> Formplan Then
                SetCustomDwgPropReliable("PLAN", planname)

            End If

            If planname <> Formplan Then
                SetCustomDwgPropReliable("PROJECT NUMBER", ProjectNumber)

            End If

            If FormSheetLabels <> frm.SheetLabels.SelectedItem Then
                SetCustomDwgPropReliable("SHEET LABELING", frm.SheetLabels.SelectedItem)
            End If

            For Each selecteditem In frm.SealsList.CheckedItems

                SealLoop.Add(selecteditem)

            Next

            Dim parts1 As New List(Of String)()

            For Each item As String In SealLoop
                If Not String.IsNullOrWhiteSpace(item) Then
                    parts1.Add(item.Trim())
                End If
            Next
            Dim result1 As String = String.Join(", ", parts1)

            If result1 <> FormStamps Then

                SetCustomDwgPropReliable("STAMPS", result1)

            End If

            Dim BuilderDivisionsList As New List(Of String)

            For Each selecteditem In frm.CheckedListBox2.CheckedItems

                BuilderDivisionsList.Add(selecteditem)

            Next

            If BuilderDivisionsList.Count = 0 Then
                BuilderDivisionsList.Add("")
            End If

            If Not frm.CheckedListBox1.CheckedItems.Contains("Both") Then
                CustomPrinting = True
            End If
            If Not frm.CheckedListBox4.CheckedItems.Contains("All") Then
                CustomPrinting = True
            End If
            If Not frm.CheckedListBox5.CheckedItems.Contains("All") Then
                CustomPrinting = True
            End If

            Dim TempElevs As New List(Of String)

            acEd.Regen()

            Using acTrans3 As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim lytab As LayerTable = acTrans3.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim alllayers As New ArrayList
                Dim fulllayerstring As String
                Dim CurrentLayer As LayerTableRecord = acTrans3.GetObject(CurrentLayerID, OpenMode.ForRead)

                acCurDb.Clayer = lytab("0")


                For Each layer In lytab

                    Dim lytr As LayerTableRecord = acTrans3.GetObject(layer, OpenMode.ForWrite)

                    If lytr.Name Like "S-ANNO-AUTOMATION" Then

                        lytr.IsPlottable = False

                    End If

                    fulllayerstring = "*S-SEAL-*"

                    If lytr.Name Like fulllayerstring Then
                        lytab.UpgradeOpen()
                        lytr.IsFrozen = True
                        lytr.IsOff = True
                    End If


                Next
                acTrans3.Commit()
                acDoc.Editor.Regen()
            End Using

            ' Start a transaction
            Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                ' Open the Block table for read
                Dim blkTable As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)

                ' Open the BlockTableRecord (ModelSpace) for read
                Dim blkTableRec As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                ' --- define desired positions ONCE before the attribute loop ---
                Dim order As String() = {Nothing, Nothing, "MOD", "SW", "PRNT", "SIZE", "OPTIONS", "ELEV", "", "LXS", "TND", "DIVISIONS"}

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

                            insertionPoint2 = blkRef.Position

                            Dim insertionPointStr As String = insertionPoint2.ToString()

                            valuesList.Add(blockobjecthandle)

                            Dim insertionExists As Boolean = AllValues.Any(Function(values) values(0) = insertionPointStr)

                            ' Only proceed if the insertion point is unique
                            'If Not insertionExists Then

                            valuesList.Add(insertionPoint2.ToString())

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

                                    ' Pick ONE: first-wins or last-wins
                                    ' First non-empty wins:
                                    If Not tagValues.ContainsKey(tag) Then tagValues(tag) = txt
                                    ' Last wins (use this instead of the line above):
                                    ' tagValues(tag) = txt

                                    ' Your extra per-tag lists:
                                    Select Case tag.ToUpperInvariant()
                                        Case "MOD"
                                            If Not PlanType.Contains(txt) Then PlanType.Add(txt)
                                        Case "ELEV"
                                            For Each v In txt.Split(","c).Select(Function(s) s.Trim()).Where(Function(s) s <> "")
                                                If Not Elevations.Contains(v) Then Elevations.Add(v)
                                            Next

                                        Case "OPTIONS"
                                            If Not Options.Contains(txt) Then Options.Add(txt)
                                    End Select
                                Next
                            End If

                            ' --- Merge into valuesList WITHOUT touching other indices ---
                            ' Make sure valuesList is big enough
                            While valuesList.Count < order.Length
                                valuesList.Add(String.Empty)
                            End While

                            ' Only write the indices managed here; leave others as-is
                            For i As Integer = 0 To order.Length - 1
                                Dim key = order(i)
                                If key IsNot Nothing AndAlso tagValues.ContainsKey(key) Then
                                    valuesList(i) = tagValues(key)
                                End If
                            Next
                            ' Add the individual values to the combined list
                            If valuesList.Count >= 7 Then
                                AllValues.Add(valuesList)
                            End If
                        End If
                        'End If
                    End If
                Next

                ' Commit the transaction
                acTrans.Commit()

            End Using

            Elevations.Sort()
            Dim NewInsertPoint As New List(Of List(Of String))
            Dim NewInsertPointList As New List(Of String)

            Dim insertionPoint1 As Point3d
            Dim LeftLayoutList As New List(Of List(Of String))
            Dim RightLayoutList As New List(Of List(Of String))
            Dim FinalLeftList As New List(Of List(Of String))
            Dim FinalRightList As New List(Of List(Of String))
            Dim counter As Integer
            counter = 1

            Dim NewFolderLocation As String

            Options.Add("")

            If CustomPrinting Then

                PlanType.Clear()
                Elevations.Clear()
                Swings.Clear()
                SealLoop.Clear()

                If frm.CheckedListBox1.SelectedItem = "Both" Then

                    For Each selecteditem In frm.CheckedListBox1.Items
                        If selecteditem <> "Both" Then
                            Swings.Add(selecteditem)
                        End If
                    Next

                Else

                    For Each selecteditem In frm.CheckedListBox1.CheckedItems
                        Swings.Add(selecteditem)
                    Next

                End If

                If frm.CheckedListBox4.SelectedItem = "All" Then

                    For Each selecteditem In frm.CheckedListBox4.Items
                        If selecteditem <> "All" Then
                            PlanType.Add(selecteditem)
                        End If
                    Next

                Else

                    For Each selecteditem In frm.CheckedListBox4.CheckedItems
                        PlanType.Add(selecteditem)
                    Next

                End If

                If frm.CheckedListBox5.SelectedItem = "All" Then

                    For Each selecteditem In frm.CheckedListBox5.Items
                        If selecteditem <> "All" Then
                            Elevations.Add(selecteditem)
                        End If

                    Next

                Else

                    For Each selecteditem In frm.CheckedListBox5.CheckedItems
                        Elevations.Add(selecteditem)
                    Next

                End If

                For Each selecteditem In frm.SealsList.CheckedItems

                    SealLoop.Add(selecteditem)

                Next

            End If

            Dim maxVal As Integer =
    AllValues.
        Where(Function(r) r IsNot Nothing AndAlso r.Count > 0).
        Select(Function(r)
                   Dim n As Integer
                   Return If(Integer.TryParse(r(4), n), n, Integer.MinValue)
               End Function).
        Max()

            If maxVal = Integer.MinValue Then
                maxVal = 0 ' no valid numbers found
            End If

            EnsureLayoutCount(maxVal)

            Dim SelectedFolder As String
            Dim isFile As Boolean
            SelectedFolder = PromptForPathOrFolder(isFile)
            If String.IsNullOrEmpty(SelectedFolder) Then Return

            If Options.Count = 1 Then Options.Add("")

            Dim SealToStamp As String

            Dim TodaysDate As String = Date.Today.ToString("MM dd yy", CultureInfo.InvariantCulture)
            For Each division In BuilderDivisionsList
                Dim divisionfolder As String
                If division <> "" Then
                    divisionfolder = "\" & division
                Else
                    divisionfolder = ""
                End If

                Dim pdfdivision As String
                If division <> "" Then
                    pdfdivision = " (" & division & ")"
                Else
                    pdfdivision = ""
                End If

                For Each stamp In SealLoop

                    SealToStamp = "Master Seal File|S-SEAL-" & stamp
                    TurnOnOrOffLayer(SealToStamp, True)

                    For Each layer In LayerLoop

                        If layer = "S-FRM-TRUSS" Then
                            TurnOnOrOffLayer(layer, True)
                            TurnOnOrOffLayer("S-FRM-STICK", False)
                        ElseIf layer = "S-FRM-STICK" Then
                            TurnOnOrOffLayer(layer, True)
                            TurnOnOrOffLayer("S-FRM-TRUSS", False)
                        End If

                        For Each type In PlanType
                            For Each ElevValue In Elevations
                                For Each opt In Options
                                    For Each valueList In AllValues

                                        'valueList(0) contains blockID
                                        'valueList(1) contains InsertionPoint
                                        'valueList(2) contains PlanType
                                        'valueList(3) contains Swing
                                        'valueList(4) contains Sequence
                                        'valueList(5) contains Scale
                                        'valueList(6) contains Option
                                        'valueList(7) contains Elevation
                                        'valuelist(8) inst set yet but is set as the single elevations when multiple


                                        If valueList(2) = type AndAlso valueList(7).Contains(ElevValue) AndAlso valueList(6) = opt And valueList(11).Contains(division) Then

                                            If valueList(7).Contains(",") Then
                                                Dim result As New List(Of String)
                                                Dim parts() As String = valueList(7).Split(","c)
                                                For Each part As String In parts
                                                    result.Add(part.Trim())
                                                Next

                                                If Not result.Contains(ElevValue.Trim(), StringComparer.OrdinalIgnoreCase) Then
                                                    Continue For
                                                End If

                                            ElseIf valueList(7).Length > ElevValue.Length Then

                                                Continue For

                                            End If

                                            'Perform the operation when both BeamLayout And FNDElev match
                                            Dim insertionPointStr As String = valueList(1) ' Example string from valueList

                                            insertionPoint1 = CreatePoint3dFromString(insertionPointStr)

                                            Dim NewHandle As String = valueList(0)
                                            Dim long1 As Long = NewHandle
                                            Dim hand As Handle = New Handle(long1)
                                            Dim objID As ObjectId = acDb.GetObjectId(False, hand, 0)

                                            If CustomPrinting Then

                                                If UCase(valueList(3)) = "L" Or UCase(valueList(3)) = "LEFT" And Swings.Contains("Left") Then

                                                    LeftLayoutList.Add(valueList)

                                                ElseIf UCase(valueList(3)) = "R" Or UCase(valueList(3)) = "RIGHT" And Swings.Contains("Right") Then

                                                    RightLayoutList.Add(valueList)

                                                End If
                                            Else

                                                If UCase(valueList(3)) = "L" Or UCase(valueList(3)) = "LEFT" Then
                                                    LeftLayoutList.Add(valueList)
                                                Else
                                                    RightLayoutList.Add(valueList)
                                                End If

                                            End If
                                        End If

                                    Next
                                    counter = 1

                                    For Each RightEntry In RightLayoutList
                                        If RightEntry(4) = counter Then
                                            RightEntry(8) = ElevValue
                                            FinalRightList.Add(RightEntry)
                                            counter += 1
                                            ZoomObjectsInViewport(RightEntry, planname, Builder, FRsheets, division)
                                        End If
                                    Next

                                    If layer = "S-FRM-STICK" Then
                                        NewFolderLocation = SelectedFolder & "\" & TodaysDate & "\" & planname & divisionfolder & "\" & stamp & "\STICK\" & type
                                    ElseIf layer = "S-FRM-TRUSS" Then
                                        NewFolderLocation = SelectedFolder & "\" & TodaysDate & "\" & planname & divisionfolder & "\" & stamp & "\TRUSS\" & type
                                    Else
                                        NewFolderLocation = SelectedFolder & "\" & TodaysDate & "\" & planname & divisionfolder & "\" & stamp & "\" & type
                                    End If


                                    If FinalRightList.Count <> 0 Then

                                        If Not Directory.Exists(NewFolderLocation) Then

                                            Directory.CreateDirectory(NewFolderLocation)

                                        End If

                                        If UCase(type) = "FRAMING" Then

                                            pdfname = UCase("RIGHT FR " & planname & " " & ElevValue)

                                        ElseIf UCase(type) = "BRACING" Then

                                            pdfname = UCase("RIGHT WB " & planname & " " & ElevValue & " (WALLBRACING)")

                                            If opt = "115 MPH" Then
                                                pdfname = pdfname & " 115 MPH"
                                            ElseIf opt = "130 MPH" Then
                                                pdfname = pdfname & " 130 MPH"
                                            ElseIf opt = "142 MPH" Then
                                                pdfname = pdfname & " 142 MPH"
                                            End If

                                        ElseIf UCase(type) = "WINDSTORM" Then

                                            pdfname = UCase("RIGHT WS " & planname & " " & ElevValue & " (150 MPH)")

                                        End If


                                    End If

                                    ' Right side
                                    If FinalRightList.Count <> 0 Then
                                        PlotTAutomatedTabs(FinalRightList, (pdfname & pdfdivision), NewFolderLocation, DPISCTB)
                                        If division <> "" Then
                                            QueueLayoutsForCsv(FinalRightList, (pdfname & pdfdivision), Builder, planname, ProjectNumber, BuilderDivision:=division)
                                        Else
                                            QueueLayoutsForCsv(FinalRightList, (pdfname & pdfdivision), Builder, planname, ProjectNumber, BuilderDivision:=division)
                                        End If
                                    End If


                                    counter = 1
                                    For Each LeftEntry In LeftLayoutList
                                        If LeftEntry(4) = counter Then
                                            LeftEntry(8) = ElevValue
                                            FinalLeftList.Add(LeftEntry)
                                            counter += 1

                                            ZoomObjectsInViewport(LeftEntry, planname, Builder, FRsheets, division)

                                        End If
                                    Next

                                    If FinalLeftList.Count <> 0 Then

                                        If Not Directory.Exists(NewFolderLocation) Then

                                            Directory.CreateDirectory(NewFolderLocation)

                                        End If

                                        If UCase(type) = "FRAMING" Then

                                            pdfname = UCase("LEFT FR " & planname & " " & ElevValue)

                                        ElseIf UCase(type) = "BRACING" Then

                                            pdfname = UCase("LEFT WB " & planname & " " & ElevValue & " (WALLBRACING)")

                                            If opt = "115 MPH" Then
                                                pdfname = pdfname & " 115 MPH"
                                            ElseIf opt = "130 MPH" Then
                                                pdfname = pdfname & " 130 MPH"
                                            ElseIf opt = "142 MPH" Then
                                                pdfname = pdfname & " 142 MPH"
                                            End If

                                        ElseIf UCase(type) = "WINDSTORM" Then

                                            pdfname = UCase("LEFT WS " & planname & " " & ElevValue & " (150 MPH)")

                                        End If

                                    End If

                                    ' Left side
                                    If FinalLeftList.Count <> 0 Then
                                        PlotTAutomatedTabs(FinalLeftList, (pdfname & pdfdivision), NewFolderLocation, DPISCTB)
                                        If division <> "" Then
                                            QueueLayoutsForCsv(FinalLeftList, (pdfname & pdfdivision), Builder, planname, ProjectNumber, BuilderDivision:=division)
                                        Else
                                            QueueLayoutsForCsv(FinalLeftList, (pdfname & pdfdivision), Builder, planname, ProjectNumber)
                                        End If

                                    End If

                                    FinalLeftList.Clear()
                                    FinalRightList.Clear()
                                    LeftLayoutList.Clear()
                                    RightLayoutList.Clear()
                                Next
                            Next

                        Next

                    Next

                    TurnOnOrOffLayer(SealToStamp, False)

                    ' ... inside the per-stamp loop in AutoFramingPrint(), after TurnOnOrOffLayer(SealToStamp, False)
                    Dim masterfolderpath As String = SelectedFolder & "\" & TodaysDate & "\" & planname & divisionfolder & "\" & stamp

                    ' Package per-elevation FR+WB first (these do NOT register to the global list)
                    If doWSFW Then
                        Dim outs As List(Of String) =
            PackageFrWbPerElevationFromCurrentPublished(
                masterfolderpath,
                planname,
                nameTemplate:="{SWING} WSFW {PLAN} {ELEV} (WSFW WALLBRACE)",
                subfolderName:="WSFW"
            )
                        For Each o In outs
                            acEd.WriteMessage(vbLf & "WSFW combined PDF written: " & o)
                        Next
                    ElseIf doRFR Then
                        Dim outs As List(Of String) =
            PackageFrWbPerElevationFromCurrentPublished(
                masterfolderpath,
                planname,
                nameTemplate:="{SWING} RFR {PLAN} {ELEV} - READY FRAME",
                subfolderName:="RFR"
            )
                        For Each o In outs
                            acEd.WriteMessage(vbLf & "RFR combined PDF written: " & o)
                        Next
                    End If

                    ' If WSFW, snapshot the FR/WB originals BEFORE MASTER (CombineRegisteredPdfs clears the registry)
                    Dim originalsToDelete As List(Of String) = Nothing
                    If doWSFW Or doRFR Then
                        originalsToDelete = New List(Of String)()
                        Dim snap = SnapshotCurrentPublished()
                        For Each p In snap
                            Dim meta As Object = Nothing
                            SyncLock _publishedPdfsLock
                                _pdfMetaByPath.TryGetValue(p, meta)
                            End SyncLock
                            Dim m = TryCast(meta, Object)
                            If meta IsNot Nothing Then
                                Dim pm = DirectCast(meta, PdfMeta)
                                If String.Equals(pm.ModName, "FRAMING", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(pm.ModName, "BRACING", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(pm.ModName, "WINDSTORM", StringComparison.OrdinalIgnoreCase) Then
                                    originalsToDelete.Add(p)
                                End If
                            End If
                        Next
                    End If

                    ' Build MASTER (this clears the registry)
                    Dim out = Path.Combine(masterfolderpath, planname & " MASTER.pdf")
                    Dim res = CombineRegisteredPdfs(out, "For Review",, 240, 0.06, 330.35)

                    If res IsNot Nothing Then
                        acEd.WriteMessage(vbLf & "Combined PDF written: " & res)

                        ' After MASTER succeeds, delete the original FR/WB PDFs if WSFW selected
                        If doWSFW Or doRFR Then
                            If originalsToDelete IsNot Nothing AndAlso originalsToDelete.Count > 0 Then
                                DeleteFilesSafe(originalsToDelete)
                                acEd.WriteMessage(vbLf & $"Deleted {originalsToDelete.Count} original FR/WB PDFs.")
                            End If
                        End If
                    Else
                        acEd.WriteMessage(vbLf & "Failed to create combined PDF.")
                    End If

                    out = Path.Combine(masterfolderpath, planname & " PERMIT.pdf")
                    res = CombineRegisteredPdfs(out, "For Permit Only",, 150, 0.06, 330.35)
                    _publishedPdfs.Clear()

                Next
            Next

            Dim lm As LayoutManager = LayoutManager.Current

            lm.CurrentLayout = "Model"
            SealLoop.Clear()
            LayerLoop.Clear()
            PlanType.Clear()
            Elevations.Clear()
            Options.Clear()

            Dim emittedCsv As String = FlushQueuedCsv(Builder, planname)
            If Not String.IsNullOrEmpty(emittedCsv) Then
                acEd.WriteMessage(vbLf & "CSV written: " & emittedCsv)
            End If

            If Not String.IsNullOrWhiteSpace(acDb.Filename) Then
                DeleteStrayDsdFiles(SelectedFolder & "\" & TodaysDate & "\")
            End If


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

        Private Sub ZoomObjectsInViewport(Layout As List(Of String), plannumber As String, Builder As String, FRSheets As Boolean, Optional division As String = "")

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim PlanString As String
            Dim Framing As Boolean = False

            If Layout(2) = "Framing" Or Layout(2) = "Bracing" Or Layout(2) = "Windstorm" Then
                Framing = True
            End If

            PlanString = "PLAN " & plannumber

            Dim Swing As String

            Dim layoutId As ObjectId

            If Layout(3) = "L" Or UCase(Layout(3)) = "LEFT" Then
                Swing = "LEFT"
            Else
                Swing = "RIGHT"
            End If

            Dim buildername As String
            If division <> "" Then
                buildername = Builder & " - " & division.ToUpper()
            Else
                buildername = Builder
            End If

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                ' Reference the Layout Manager
                Dim acLayoutMgr As LayoutManager = LayoutManager.Current
                Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

                layoutId = acLayoutMgr.GetLayoutId(Layout(4))

                Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                Try
                    Dim Layid As ObjectId

                    Layid = layouts.GetAt(Layout(4))

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

                                    If tagvalue.Contains("PLAN ") Then

                                        attref.TextString = PlanString

                                    ElseIf tagvalue.Contains("OPTIONAL") Then

                                        If UCase(Layout(2)) = "FRAMING" Then

                                            attref.TextString = "FRAMING PLAN"

                                        ElseIf UCase(Layout(2)) = "BRACING" Then

                                            If Layout(6) = "115 MPH" Then

                                                attref.TextString = "WALL BRACING PLAN - 115 MPH ULTIMATE 3-SECOND GUST"

                                            ElseIf Layout(6) = "130 MPH" Then

                                                attref.TextString = "WALL BRACING PLAN - 130 MPH ULTIMATE 3-SECOND GUST"

                                            ElseIf Layout(6) = "142 MPH" Then

                                                attref.TextString = "WALL BRACING PLAN - 142 MPH ULTIMATE 3-SECOND GUST"

                                            End If

                                        ElseIf UCase(Layout(2)) = "WINDSTORM" Then

                                            attref.TextString = "WINDSTORM PLAN - 150 MPH ULTIMATE 3-SECOND GUST"
                                        End If

                                    ElseIf tagvalue.Contains("ELEVATION") Then

                                        attref.TextString = "ELEVATION " & Layout(8) & " - " & Swing & " SWING"

                                    ElseIf tagvalue.Contains("CUSTOMER'S NAME") Then

                                        attref.TextString = UCase(buildername)

                                    ElseIf tagvalue.Contains("PLANDATE") Then

                                        attref.TextString = Date.Today.ToString("d")

                                    ElseIf tagvalue.Contains("FR-1") Then

                                        If FRSheets Then

                                            If UCase(Layout(2)) = "FRAMING" Then

                                                attref.TextString = "FR-" & Layout(4)

                                            ElseIf UCase(Layout(2)) = "BRACING" Then

                                                attref.TextString = "WB-" & Layout(4)

                                            ElseIf UCase(Layout(2)) = "WINDSTORM" Then

                                                attref.TextString = "W-" & Layout(4)

                                            End If
                                        Else
                                            If UCase(Layout(2)) = "FRAMING" Then

                                                attref.TextString = "S2." & Layout(4)

                                            ElseIf UCase(Layout(2)) = "BRACING" Then

                                                attref.TextString = "S3." & Layout(4)

                                            ElseIf UCase(Layout(2)) = "WINDSTORM" Then

                                                attref.TextString = "S3." & Layout(4) & "T"

                                            End If
                                        End If
                                    ElseIf tagvalue.Contains("1/8"" = 1'-0""") Then
                                        If Layout(5) = "1/8" Then
                                            attref.TextString = "1/8"" = 1'-0"""
                                        ElseIf Layout(5) = "3/32" Then
                                            attref.TextString = "3/32"" = 1'-0"""
                                        End If
                                    End If

                                Next

                            End If


                        End If

                    Next

                    ModifyViewPortCenter(vpIds, Layout(4), Layout(1), Layout(5))

                Catch es As Exception
                    MsgBox(es.Message)
                End Try

                ' Save the changes made
                acTrans.Commit()

            End Using

        End Sub

        <CommandMethod("SyncTBFromModelBlock")>
        Public Shared Sub SyncAttributesFromModelBlockToTitleblocks()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database

            ' Prompt user to select a block reference in model space
            Dim peo As New PromptEntityOptions(vbLf & "Select source block in Model space to copy attributes from:")
            peo.SetRejectMessage(vbLf & "Only block references are allowed.")
            peo.AddAllowedClass(GetType(BlockReference), False)
            Dim per As PromptEntityResult = ed.GetEntity(peo)
            If per.Status <> PromptStatus.OK Then
                ed.WriteMessage(vbLf & "Selection cancelled.")
                Return
            End If

            Using acLck As DocumentLock = doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Try
                        Dim srcBr As BlockReference = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead), BlockReference)
                        If srcBr Is Nothing Then
                            ed.WriteMessage(vbLf & "Selected entity is not a block reference.")
                            Return
                        End If

                        ' Collect attributes from the source block (tags -> value)
                        Dim attrValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                        For Each aid As ObjectId In srcBr.AttributeCollection
                            Dim a As AttributeReference = TryCast(tr.GetObject(aid, OpenMode.ForRead), AttributeReference)
                            If a IsNot Nothing Then
                                Dim tag = If(a.Tag, String.Empty).Trim()
                                If tag <> String.Empty Then
                                    attrValues(tag) = If(a.TextString, String.Empty)
                                End If
                            End If
                        Next

                        If attrValues.Count = 0 Then
                            ed.WriteMessage(vbLf & "No attributes found on selected block.")
                            Return
                        End If

                        ' Iterate layouts and update title block instances
                        Dim updatedCount As Integer = 0

                        Dim layoutDict As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                        For Each entry As DBDictionaryEntry In layoutDict
                            Dim layout As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                            If layout Is Nothing OrElse layout.ModelType Then Continue For

                            Dim btrId As ObjectId = layout.BlockTableRecordId
                            Dim btr As BlockTableRecord = TryCast(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                            If btr Is Nothing Then Continue For

                            For Each entId As ObjectId In btr
                                If entId.ObjectClass.DxfName <> "INSERT" Then Continue For
                                Dim tbRef As BlockReference = TryCast(tr.GetObject(entId, OpenMode.ForWrite), BlockReference)
                                If tbRef Is Nothing Then Continue For

                                Dim defRec As BlockTableRecord = TryCast(tr.GetObject(tbRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                If defRec Is Nothing Then Continue For

                                Dim defName = If(defRec.Name, String.Empty)
                                If defName.Equals("Arcxis Title Block", StringComparison.OrdinalIgnoreCase) OrElse defName.Equals("DPIS RevisionBlock", StringComparison.OrdinalIgnoreCase) Then
                                    ' Update matching attributes on the title block instance
                                    For Each attId As ObjectId In tbRef.AttributeCollection
                                        Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                        If attRef Is Nothing Then Continue For
                                        Dim tag = If(attRef.Tag, String.Empty).Trim()
                                        If tag = String.Empty Then Continue For
                                        If attrValues.ContainsKey(tag) Then
                                            attRef.TextString = attrValues(tag)
                                            If attRef.TextString.Length > 36 Then
                                                attRef.Height = 1 / 32
                                            End If
                                            updatedCount += 1
                                        End If
                                    Next
                                End If
                            Next
                        Next

                        tr.Commit()
                        ed.WriteMessage(vbLf & $"Updated {updatedCount} titleblock attribute values from selected block.")
                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "Error: " & ex.Message)
                    End Try
                End Using
            End Using
        End Sub

        Private Sub ModifyViewPortCenter(VPIDS As ObjectIdCollection, LAYOUT As String, BASEPOINT As String, Scale As String)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = acDoc.Database
            Dim basePt As Point3d = CreatePoint3dFromString(BASEPOINT)

            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim lm As LayoutManager = LayoutManager.Current
                Dim layoutId As ObjectId = lm.GetLayoutId(LAYOUT)
                Dim layout1 As Layout = tr.GetObject(layoutId, OpenMode.ForRead)

                ' Keep the viewport frame fixed in paperspace
                Dim paperCenter As New Point3d(7.706, 5.5, 0)

                ' Offsets in MODEL units for where you want to look
                Dim dx As Double = 0.0, dy As Double = 0.0

                ' AutoCAD API: CustomScale = model units per paper unit (e.g., 1/8"=1' -> 96)
                Dim customScale As Double
                Select Case Scale
                    Case "1/8" : customScale = 96.0 : dx = 709.92 : dy = -504.0
                    Case "3/32" : customScale = 128.0 : dx = 946.558 : dy = -672.0
                    Case Else : customScale = 1.0
                End Select

                ' Touch ONLY the requested viewports
                For Each vpId As ObjectId In VPIDS
                    Dim vp = TryCast(tr.GetObject(vpId, OpenMode.ForRead), Viewport)
                    If vp Is Nothing Then Continue For
                    'If vp.Number = 1 Then Continue For ' never touch overall PS viewport


                    vp.UpgradeOpen()
                    Dim wasLocked = vp.Locked
                    vp.Locked = False


                    ' Do not change vp.Width / vp.Height (paper units)

                    ' Camera straight down, no twist
                    vp.ViewDirection = Teigha.Geometry.Vector3d.ZAxis
                    vp.TwistAngle = 0.0

                    ' Deterministic zoom: modelHeight = paperHeight * CustomScale
                    If vp.Height > 10.5 Then Continue For
                    If vp.CustomScale <> 1 / customScale Then
                        vp.CustomScale = 1 / customScale
                    End If

                    vp.ViewHeight = vp.Height * customScale
                    ' (If you ever need width-based checks: modelWidth = vp.Width * customScale)

                    ' Aim at a model point; use BOTH ViewTarget and ViewCenter to the same XY
                    Dim cx As Double = basePt.X + dx
                    Dim cy As Double = basePt.Y + dy
                    'vp.ViewTarget = New Point3d(cx, cy, 0.0)
                    vp.ViewCenter = New Point2d(cx, cy)

                    ' Keep the paperspace frame put
                    vp.CenterPoint = paperCenter

                    ' No UpdateDisplay needed; Commit will flush
                    vp.Locked = wasLocked
                    vp.DowngradeOpen()
                Next

                tr.Commit()
            End Using
        End Sub

        Private Sub PutAt(ByRef list As List(Of String), index As Integer, value As String)
            If list Is Nothing Then list = New List(Of String)()
            While list.Count <= index
                list.Add(String.Empty)
            End While
            list(index) = value
        End Sub

        Public Function CreatePoint3dFromString(pointStr As String) As Point3d
            ' Split the string by commas or spaces, depending on the format

            pointStr = pointStr.Trim("(", ")")

            Dim coords() As String = pointStr.Split(New Char() {","c, " "c}, StringSplitOptions.RemoveEmptyEntries)

            ' Ensure we have three coordinates
            If coords.Length <> 3 Then
                Throw New ArgumentException("The point String must contain exactly three coordinates.")
            End If

            ' Convert the coordinates from string to double
            Dim x As Double = Double.Parse(coords(0).Trim())
            Dim y As Double = Double.Parse(coords(1).Trim())
            Dim z As Double = Double.Parse(coords(2).Trim())

            ' Return a new Point3d object
            Return New Point3d(x, y, z)
        End Function

        ' Reworked: write ONE logical set per PDF as Identifier/Key/Value rows.
        ' Column 1: Identifier = "Builder - PdfName"
        ' Column 2: Key        = PLAN | ELEV | SW | MOD
        ' Column 3: Value
        ' Creates a small CSV per PDF (4 data rows + header).
        Private Sub WriteLayoutsToCsv(layoutList As List(Of List(Of String)), pdfName As String, builder As String, planName As String)
            Dim pendingDir As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\fs2\k\DPIS Drawings\PDF File Data\Pending"
            If Not Directory.Exists(pendingDir) Then Directory.CreateDirectory(pendingDir)

            ' Derive values from first layout
            Dim modName As String = ""
            Dim swingRaw As String = ""
            Dim elevation As String = ""

            If layoutList IsNot Nothing AndAlso layoutList.Count > 0 Then
                Dim first = layoutList(0)
                If first.Count > 2 Then modName = If(first(2), "").Trim()
                If first.Count > 3 Then swingRaw = If(first(3), "").Trim()
                ' Prefer single elevation (index 8), fallback to index 7 (all elevs)
                If first.Count > 8 AndAlso Not String.IsNullOrWhiteSpace(first(8)) Then
                    elevation = first(8).Trim()
                ElseIf first.Count > 7 Then
                    elevation = If(first(7), "").Trim()
                End If
            End If

            swingRaw = NormalizeSwing(swingRaw)

            Dim identifier As String = If(builder, "").Trim() & " - " & If(pdfName, "").Trim()

            ' Unique file name per PDF (keep user/timestamp pattern)
            Dim user As String = Environment.UserName
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim safePdfName As String = Path.GetFileNameWithoutExtension(pdfName).Replace(" ", "_")
            Dim csvPath As String = Path.Combine(pendingDir, $"{user}_{safePdfName}_{timestamp}.csv")

            Using sw As New StreamWriter(csvPath, False, System.Text.Encoding.UTF8)
                sw.WriteLine("Identifier,Key,Value")
                WriteKv(sw, identifier, "PLAN", planName)
                WriteKv(sw, identifier, "ELEV", elevation)
                WriteKv(sw, identifier, "SW", swingRaw)
                WriteKv(sw, identifier, "MOD", modName)
            End Using
        End Sub

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
        Private Sub AppendLayoutsToMainExcel(layoutList As List(Of List(Of String)), pdfName As String, builder As String, planName As String)
            Dim mainExcelPath As String = Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\fs2\k\DPIS Drawings\PDF File Data\PDF Info.xlsx"
            Dim xlApp As Excel.Application = Nothing
            Dim xlWb As Excel.Workbook = Nothing
            Dim xlWs As Excel.Worksheet = Nothing

            Try
                Dim modName As String = ""
                Dim swingRaw As String = ""
                Dim elevation As String = ""

                If layoutList IsNot Nothing AndAlso layoutList.Count > 0 Then
                    Dim first = layoutList(0)
                    If first.Count > 2 Then modName = If(first(2), "").Trim()
                    If first.Count > 3 Then swingRaw = If(first(3), "").Trim()
                    If first.Count > 8 AndAlso Not String.IsNullOrWhiteSpace(first(8)) Then
                        elevation = first(8).Trim()
                    ElseIf first.Count > 7 Then
                        elevation = If(first(7), "").Trim()
                    End If
                End If

                swingRaw = NormalizeSwing(swingRaw)

                ' Replace C# ?? with VB If(...)
                Dim identifier As String = If(builder, "").Trim() & " - " & If(pdfName, "").Trim()

                xlApp = New Excel.Application()
                xlApp.DisplayAlerts = False

                Dim startRow As Integer
                If IO.File.Exists(mainExcelPath) Then
                    xlWb = xlApp.Workbooks.Open(mainExcelPath)
                    xlWs = CType(xlWb.Sheets(1), Excel.Worksheet)
                    startRow = xlWs.Cells(xlWs.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
                Else
                    xlWb = xlApp.Workbooks.Add()
                    xlWs = CType(xlWb.Sheets(1), Excel.Worksheet)
                    xlWs.Cells(1, 1).Value = "Identifier"
                    xlWs.Cells(1, 2).Value = "Key"
                    xlWs.Cells(1, 3).Value = "Value"
                    startRow = 2
                End If

                ' Explicit System.Tuple to avoid ambiguity
                Dim kv As New List(Of System.Tuple(Of String, String)) From {
                    New System.Tuple(Of String, String)("PLAN", planName),
                    New System.Tuple(Of String, String)("ELEV", elevation),
                    New System.Tuple(Of String, String)("SW", swingRaw),
                    New System.Tuple(Of String, String)("MOD", modName)
                }

                Dim r As Integer = startRow
                For Each pair In kv
                    xlWs.Cells(r, 1).Value = identifier
                    xlWs.Cells(r, 2).Value = pair.Item1
                    xlWs.Cells(r, 3).Value = pair.Item2
                    r += 1
                Next

                xlWb.SaveAs(mainExcelPath)
                xlWb.Close()
                xlApp.Quit()
            Catch ex As Exception
                MessageBox.Show("Error writing to main Excel file: " & ex.Message)
            Finally
                If xlWs IsNot Nothing Then Marshal.ReleaseComObject(xlWs)
                If xlWb IsNot Nothing Then Marshal.ReleaseComObject(xlWb)
                If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub PlotTAutomatedTabs(lAYOUTLIST As List(Of List(Of String)), Pdfname As String, NEWFOLDERLOCATION As String, Optional DPISCTB As Boolean = False)
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acDoc.LockDocument()

                ' then set de.Nps = targetNps for each DsdEntry

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    ' Save & force no UI prompts while publishing
                    Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                    Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                    Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                    Application.SetSystemVariable("BackGroundPlot", 0)
                    Application.SetSystemVariable("CMDDIA", 0)
                    Application.SetSystemVariable("FILEDIA", 0)

                    Try
                        ' 1) Ensure output folder exists
                        Dim outputDir As String = NEWFOLDERLOCATION
                        If Not Directory.Exists(outputDir) Then Directory.CreateDirectory(outputDir)

                        ' 2) Current DWG path
                        Dim dwgprefix As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
                        Dim DWGnm As String = CStr(Application.GetSystemVariable("DWGNAME"))
                        Dim dwgFile As String = Path.Combine(dwgprefix, DWGnm)

                        ' 3) Build target file paths safely
                        Dim pdfFile As String = Path.Combine(outputDir, Pdfname & ".pdf")
                        Dim dsdFile As String = Path.Combine(outputDir, Pdfname & ".dsd")

                        If File.Exists(pdfFile) Then
                            acTrans.Commit()
                            Return
                        End If

                        ' 4) Build DSD entries
                        Dim dsd As New DsdData()
                        Dim dsdEntries As New DsdEntryCollection()

                        For Each lay As List(Of String) In lAYOUTLIST
                            Dim title As String = Path.GetFileNameWithoutExtension(DWGnm) & "-" &
                                      lay(2) & "-" & lay(3) & "-" & lay(7) & "-" & lay(4)

                            Dim de As New DsdEntry()
                            de.DwgName = dwgFile
                            de.Layout = lay(4)           ' layout name
                            de.Title = title

                            ' when creating each DsdEntry:
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
                "IncludeLayer=FALSE",
                "Type=6",
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

                        ' wait briefly for the PDF to appear then register it
                        If WaitForFileExists(pdfFile, 30000) Then
                            RegisterPublishedPdf(pdfFile)
                            ' Register MOD/Elevation/Swing using the layouts that produced this PDF
                            RegisterPublishedPdfMeta(pdfFile, lAYOUTLIST)
                        End If

                        ' Cleanup
                        If File.Exists(dsdFile) Then File.Delete(dsdFile)

                        acTrans.Commit()

                    Finally
                        ' Restore system vars
                        Application.SetSystemVariable("BackGroundPlot", bgPrev)
                        Application.SetSystemVariable("CMDDIA", cmdPrev)
                        Application.SetSystemVariable("FILEDIA", fileDiaPrev)
                    End Try

                End Using
            End Using
        End Sub

        Private Sub PlotAutomatedTabs(lAYOUTLIST As List(Of List(Of String)), Pdfname As String, NEWFOLDERLOCATION As String)
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

            Using acDoc.LockDocument()
                ' Force no UI
                Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                Application.SetSystemVariable("BackGroundPlot", 0)
                Application.SetSystemVariable("CMDDIA", 0)
                Application.SetSystemVariable("FILEDIA", 0)

                Try
                    ' Paths
                    Dim dwgprefix As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
                    Dim dwgName As String = CStr(Application.GetSystemVariable("DWGNAME"))
                    Dim dwgFile As String = Path.Combine(dwgprefix, dwgName)

                    Dim outputDir As String = NEWFOLDERLOCATION
                    If Not Directory.Exists(outputDir) Then Directory.CreateDirectory(outputDir)

                    Dim pdfFile As String = Path.Combine(outputDir, Pdfname & ".pdf")
                    Dim dsdFile As String = Path.Combine(outputDir, Pdfname & ".dsd")

                    ' If already published, bail fast
                    If File.Exists(pdfFile) Then Exit Sub

                    Dim dsd As New DsdData()
                    Dim dsdEntries As New DsdEntryCollection()

                    ' Build entries (NEW entry per layout)
                    For Each lay As List(Of String) In lAYOUTLIST
                        Dim title As String = Path.GetFileNameWithoutExtension(dwgName) & "-" &
                                      lay(2) & "-" & lay(8) & "-" & lay(3) & "-" & lay(4)

                        Dim de As New DsdEntry()
                        de.DwgName = dwgFile
                        de.Layout = lay(4)
                        de.Title = title
                        de.Nps = "Arcxis"            ' use your named page setup if it exists
                        de.NpsSourceDwg = dwgFile
                        dsdEntries.Add(de)
                    Next

                    ' Set once (not inside the loop)
                    dsd.SetDsdEntryCollection(dsdEntries)
                    dsd.SheetType = SheetType.MultiPdf
                    dsd.NoOfCopies = 1
                    dsd.IsHomogeneous = True
                    dsd.ProjectPath = outputDir
                    dsd.DestinationName = pdfFile

                    ' Optional: turn off heavy features unless you need them
                    dsd.Dwf3dOptions.PublishWithMaterials = False
                    dsd.Dwf3dOptions.GroupByXrefHierarchy = False

                    ' Suppress prompts
                    dsd.SetUnrecognizedData("PromptForDwfName", "FALSE")
                    dsd.SetUnrecognizedData("PromptForPwd", "FALSE")

                    ' Write DSD, normalize flags, and re-read
                    If File.Exists(dsdFile) Then File.Delete(dsdFile)
                    dsd.WriteDsd(dsdFile)

                    Dim text As String = File.ReadAllText(dsdFile)
                    text = text.Replace("PromptForDwfName=TRUE", "PromptForDwfName=FALSE")
                    text = text.Replace("PromptForName=TRUE", "PromptForName=FALSE")
                    text = text.Replace("Setup=", "Setup=Arcxis")
                    ' Ensure destination fields exist
                    If Not text.Contains(vbCrLf & "OutDir=") Then text &= vbCrLf & "OutDir=" & outputDir
                    If Not text.Contains(vbCrLf & "Dst=") Then text &= vbCrLf & "Dst=" & pdfFile
                    File.WriteAllText(dsdFile, text)

                    dsd.ReadDsd(dsdFile)

                    ' Use PC3 (or pass Nothing)
                    Dim pc As PlotConfig = Nothing
                    Try : pc = PlotConfigManager.SetCurrentConfig("ARCXIS - DWG To PDF - Brics.pc3") : Catch : End Try

                    ' Go!
                    Application.Publisher.PublishExecute(dsd, pc)

                    If File.Exists(dsdFile) Then File.Delete(dsdFile)

                Finally
                    ' Restore system vars
                    Application.SetSystemVariable("BackGroundPlot", bgPrev)
                    Application.SetSystemVariable("CMDDIA", cmdPrev)
                    Application.SetSystemVariable("FILEDIA", fileDiaPrev)
                End Try
            End Using
        End Sub

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

        <CommandMethod("CleanDPIS")>
        Public Sub CleanUpLayoutsAndBlocks()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor

            Using acDoc.LockDocument()
                Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()

                    ' === 1. Delete all layouts except Modelspace ===
                    Dim lm As LayoutManager = LayoutManager.Current
                    Dim bt As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                    Dim dictLayouts As DBDictionary = acTrans.GetObject(acDb.LayoutDictionaryId, OpenMode.ForRead)

                    ' === 2. Erase all instances of specified blocks ===
                    Dim blkNames As String() = {"Seal-Tim-Texas", "TML TX"}
                    For Each blkName In blkNames
                        If bt.Has(blkName) Then
                            Dim btr As BlockTableRecord = acTrans.GetObject(bt(blkName), OpenMode.ForRead)
                            Dim idsToErase As New List(Of ObjectId)

                            ' Collect all references to this block
                            For Each id As ObjectId In btr.GetBlockReferenceIds(True, True)
                                idsToErase.Add(id)
                            Next

                            ' Erase them
                            For Each id As ObjectId In idsToErase
                                Dim ent As Entity = TryCast(acTrans.GetObject(id, OpenMode.ForWrite), Entity)
                                If ent IsNot Nothing Then
                                    ent.Erase()
                                End If
                            Next
                        End If
                    Next

                    Dim xrefs As String() = {"ArcxisLogo"}
                    For Each xref In xrefs
                        DetachXrefs(acDb, acTrans, xref)
                    Next

                    acTrans.Commit()
                End Using
            End Using

            ' === 3. Purge block definitions ===
            Using acDoc.LockDocument()
                Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                    Dim purgeNames As String() = {"Seal-Tim-Texas", "TML TX"}

                    For Each blkName In purgeNames
                        If bt.Has(blkName) Then
                            Dim id As ObjectId = bt(blkName)
                            Dim ids As New ObjectIdCollection()
                            ids.Add(id)

                            acDb.Purge(ids)
                            If ids.Count > 0 Then
                                Dim btr As BlockTableRecord = acTrans.GetObject(ids(0), OpenMode.ForWrite)
                                Try
                                    btr.Erase()
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                vbLf & $"Purged block: {blkName}")
                                Catch ex As Exception
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                vbLf & $"Could not purge block {blkName}: {ex.Message}")
                                End Try
                            End If
                        End If
                    Next

                    acTrans.Commit()
                End Using
            End Using
        End Sub



        <CommandMethod("CleanUpFile")>
        Sub CleanUpFile()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    For Each objId As ObjectId In ms
                        Dim ent As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                        If TypeOf ent Is DBText Then
                            Dim txt As DBText = CType(ent, DBText)
                            Dim textStr As String = txt.TextString.Trim().ToUpper()


                            'for explanation
                            '(?i) means Case-insensitive
                            '\s*[^-]+ means Some text (not a dash), optional spaces

                            If System.Text.RegularExpressions.Regex.IsMatch(textStr, "^(?i)\s*[^-]+\s*-\s*[^-]+\s*-\s*[^-]+\s*$") Then
                                txt.UpgradeOpen()
                                txt.Erase()
                                Continue For
                            End If

                        ElseIf TypeOf ent Is BlockReference Then

                            Dim blkRef As BlockReference = CType(ent, BlockReference)
                            Dim btr As BlockTableRecord = acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead)
                            Dim blockName As String = btr.Name

                            ' === CASE 1: Named "TB-INFO" ===
                            If blockName.Equals("TB-INFO", StringComparison.OrdinalIgnoreCase) Then
                                blkRef.UpgradeOpen()
                                blkRef.Erase()
                                Continue For
                            End If

                            ' === CASE 2: Anonymous A$ blocks ===
                            If blockName.StartsWith("A$", StringComparison.OrdinalIgnoreCase) Then
                                Dim onlyText As Boolean = True
                                Dim foundPattern As Boolean = False

                                For Each subEntId As ObjectId In btr
                                    Dim subEnt As Entity = acTrans.GetObject(subEntId, OpenMode.ForRead)
                                    If TypeOf subEnt Is DBText Then
                                        Dim txt As DBText = CType(subEnt, DBText)
                                        Dim textStr As String = txt.TextString.Trim().ToUpper()

                                        ' Match "text - text - LEFT/RIGHT"
                                        If System.Text.RegularExpressions.Regex.IsMatch(textStr, "^(?i)\s*[^-]+\s*-\s*[^-]+\s*-\s*[^-]+\s*$") Then
                                            foundPattern = True
                                        End If
                                    Else
                                        onlyText = False
                                        Exit For
                                    End If
                                Next

                                If onlyText AndAlso foundPattern Then
                                    blkRef.UpgradeOpen()
                                    blkRef.Erase()
                                End If

                            End If

                        End If

                    Next


                    acTrans.Commit()

                End Using

            End Using

            CleanUpLayoutsAndBlocks()

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

        <CommandMethod("CreateLayoutsWithTitleblock")>
        Public Shared Sub CallLayoutCreater()
            Dim frm As New Form_AddLayouts()
            Dim NumberOfLayouts As Integer = 0
            Dim CTBName As String = ""
            frm.ComboBox1.SelectedIndex = 0
            frm.ShowDialog()
            If frm.DialogResult <> DialogResult.OK Then
                Exit Sub
            End If
            NumberOfLayouts = frm.TextBox1.Text
            CTBName = frm.ComboBox1.SelectedItem
            CreateLayoutsWithTitleblock(NumberOfLayouts, , CTBName)
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
                        acPlSetVdr.SetPlotConfigurationName(acPlSet, "ARCXIS - DWG To PDF - Brics.pc3", "ANSI_full_bleed_B_(11.00_x_17.00_Inches)")
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
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, CTBFile)

                        ' Copy all settings to the layout
                        acLayout.CopyFrom(acPlSet)
                    Catch es As Exception
                        MsgBox(es.Message & vbCrLf & "Error setting plot configuration. Check PC3 and paper size name.")
                    End Try

                    acTrans.Commit()
                End Using
            End Using
        End Sub

        <CommandMethod("ListPaperSizes")>
        Public Shared Sub ListAvailablePaperSizes()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Try
                    ' Get the current layout to copy settings from
                    Dim lm As LayoutManager = LayoutManager.Current
                    Dim layoutId As ObjectId = lm.GetLayoutId(lm.CurrentLayout)
                    Dim layout As Layout = acTrans.GetObject(layoutId, OpenMode.ForRead)

                    ' Create a temporary PlotSettings and copy from layout
                    Dim tempPlotSettings As New PlotSettings(False)
                    tempPlotSettings.CopyFrom(layout)

                    Dim validator As PlotSettingsValidator = PlotSettingsValidator.Current

                    ' Get list of available plotters first
                    Dim devices As StringCollection = validator.GetPlotDeviceList()

                    Dim targetDevice As String = "ARCXIS - DWG To PDF - Brics.pc3"
                    Dim foundDevice As Boolean = devices.Cast(Of String)().Any(Function(d) d.Equals(targetDevice, StringComparison.OrdinalIgnoreCase))

                    If Not foundDevice Then
                        ed.WriteMessage(vbLf & "ERROR: PC3 file '" & targetDevice & "' not found!")
                        acTrans.Commit()
                        Return
                    End If

                    ' **KEY FIX: Set the device with a known-good paper size first**
                    ' Try common 11x17 paper size names that typically exist
                    Dim tryMediaNames As String() = {
                "ANSI B (11.00 x 17.00 Inches)",
                "ANSI_B_(11.00_x_17.00_Inches)",
                "Tabloid (11 x 17 in)",
                "11x17"
            }

                    Dim successMedia As String = Nothing
                    For Each mediaName In tryMediaNames
                        Try
                            validator.SetPlotConfigurationName(tempPlotSettings, targetDevice, mediaName)
                            successMedia = mediaName
                            Exit For
                        Catch
                            ' Try next one
                        End Try
                    Next

                    If successMedia Is Nothing Then
                        ' Last resort: try with first available media from a default plotter
                        Try
                            ' Use DWG To PDF.pc3 (built-in) to get a valid media name
                            Dim tempPs As New PlotSettings(False)
                            tempPs.CopyFrom(layout)
                            validator.SetPlotConfigurationName(tempPs, "DWG To PDF.pc3", Nothing)
                            Dim defaultMedia As StringCollection = validator.GetCanonicalMediaNameList(tempPs)
                            If defaultMedia.Count > 0 Then
                                ' Try the first media from built-in plotter
                                validator.SetPlotConfigurationName(tempPlotSettings, targetDevice, defaultMedia(0))
                                successMedia = defaultMedia(0)
                            End If
                        Catch
                            ed.WriteMessage(vbLf & "ERROR: Could not initialize plot device with any paper size.")
                            ed.WriteMessage(vbLf & "Your PC3 file may be corrupted or incompatible.")
                            acTrans.Commit()
                            Return
                        End Try
                    End If

                    ' Now get the full list of available media
                    Dim mediaNames As StringCollection = validator.GetCanonicalMediaNameList(tempPlotSettings)

                    ed.WriteMessage(vbLf & vbLf & "=== Available Paper Sizes for '" & targetDevice & "' ===" & vbLf)
                    ed.WriteMessage(vbLf & "(Successfully initialized with: " & successMedia & ")" & vbLf)

                    Dim count As Integer = 0
                    For Each mediaName As String In mediaNames
                        Dim localName As String = validator.GetLocaleMediaName(tempPlotSettings, mediaName)
                        ed.WriteMessage(vbLf & $"  [{count}] Canonical: '{mediaName}'")
                        ed.WriteMessage(vbLf & $"      Display:   '{localName}'")
                        count += 1
                    Next

                    ed.WriteMessage(vbLf & vbLf & $"Total: {count} paper sizes found.")
                    ed.WriteMessage(vbLf & "Use the CANONICAL name in SetPlotConfigurationName()." & vbLf)

                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message)
                    ed.WriteMessage(vbLf & "Stack: " & ex.StackTrace)
                End Try

                acTrans.Commit()
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

        Public Function PromptForPathOrFolder(Optional ByRef userChoseFilePath As Boolean = False) As String
            Dim startPath As String = GetCurrentDwgFolder()
            Dim owner As New WindowWrapper(Application.MainWindow.Handle)
            Dim chosen As String = Nothing

            Using dlg As New FolderBrowserDialog()
                dlg.Description = "Choose a folder (starts at current DWG folder)"
                If Directory.Exists(startPath) Then dlg.SelectedPath = startPath
                dlg.ShowNewFolderButton = True

                Dim result = dlg.ShowDialog(owner)
                If result = DialogResult.OK AndAlso Directory.Exists(dlg.SelectedPath) Then
                    chosen = dlg.SelectedPath
                    userChoseFilePath = False

                    ' Set focus back to AutoCAD and optionally to the DWG folder
                    Application.MainWindow.Focus()
                    Try
                        Environment.CurrentDirectory = startPath
                    Catch
                        ' Ignore if fails
                    End Try

                    Return chosen
                End If
            End Using

            userChoseFilePath = False
            Return Nothing
        End Function

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


        <CommandMethod("overnightprinting")>
        Sub OverNightPrinting()

            Dim acDoc = Application.DocumentManager.MdiActiveDocument
            Dim acDb = acDoc.Database
            Dim accurdb = acDoc.Database
            Dim ed = acDoc.Editor
            Dim acEd As Editor = acDoc.Editor
            Dim BlockBuilder As String = ""
            Dim BlockPlan As String = ""
            Dim BlockStamps As String = ""
            Dim PropertiesBuilder As String
            Dim PropertiesPlan As String
            Dim PropertiesStamps As String
            Dim PropertiesProject As String
            Dim PropertiesSheetLabeling As String
            Dim PropertiesPlanType As String
            Dim PropertiesPlanDivisions As String
            Dim Builder As String
            Dim Planname As String
            Dim Stamps As String
            Dim wrong As Boolean = False
            Dim PlanType As New List(Of String)
            Dim Swings As New List(Of String)
            Dim Elevations As New List(Of String)
            Dim Options As New List(Of String)
            Dim AllValues As New List(Of List(Of String))
            Dim insertionPoint2 As Point3d
            Dim pdfname As String = ""
            Dim CurrentLayerID As ObjectId = accurdb.Clayer
            Dim frm As New Form_Automated_Printing_Info
            Dim CustomPrinting As Boolean = False
            Dim content As New System.Text.StringBuilder()

            Dim FRSheets As Boolean = False


            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = accurdb.TransactionManager.StartTransaction()

                    Dim bt As BlockTable = acTrans.GetObject(accurdb.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    PropertiesBuilder = GetCustomDwgPropForDoc(accurdb, "BUILDER")
                    PropertiesPlan = GetCustomDwgPropForDoc(accurdb, "PLAN")
                    PropertiesStamps = GetCustomDwgPropForDoc(accurdb, "STAMPS")
                    PropertiesProject = GetCustomDwgPropReliable("PROJECT NUMBER")
                    PropertiesSheetLabeling = GetCustomDwgPropReliable("SHEET LABELING")
                    PropertiesPlanType = If(GetCustomDwgPropReliable("PLAN TYPE"), "").Trim()
                    PropertiesPlanDivisions = GetCustomDwgPropReliable("DIVISIONS")
                    ' Validate required properties
                    Dim missingProps As New List(Of String)()

                    If String.IsNullOrWhiteSpace(PropertiesBuilder) Then missingProps.Add("BUILDER")
                    If String.IsNullOrWhiteSpace(PropertiesPlan) Then missingProps.Add("PLAN")
                    If String.IsNullOrWhiteSpace(PropertiesPlanType) Then
                        missingProps.Add("PLAN TYPE")
                    ElseIf String.Equals(PropertiesPlanType, "FRAMING", StringComparison.OrdinalIgnoreCase) Then
                        If String.IsNullOrWhiteSpace(PropertiesSheetLabeling) Then missingProps.Add("SHEET LABELING")
                        If String.IsNullOrWhiteSpace(PropertiesStamps) Then missingProps.Add("STAMPS")
                    End If


                    If missingProps.Count > 0 Then
                        ' Build the content string for the text file
                        Content.AppendLine("=== MISSING DRAWING PROPERTIES ===")
                        Content.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                        Content.AppendLine($"DWG File: {Path.GetFileName(accurdb.Filename)}")
                        Content.AppendLine()
                        Content.AppendLine("The following required drawing properties are missing or empty:")
                        Content.AppendLine()
                        For Each prop In missingProps
                            Content.AppendLine($"  - {prop}")
                        Next
                        Content.AppendLine()
                        Content.AppendLine("Please set these properties before submitting the drawing.")

                        ' Create the text file with the missing properties (no MessageBox for accore)
                        Dim txtPath As String = CreateTextFileInDwgLocation("ReasonForFailure.log", Content.ToString())

                        Exit Sub
                    End If

                    If String.Equals(PropertiesSheetLabeling, "FR", StringComparison.OrdinalIgnoreCase) Then FRSheets = True

                    If String.Equals(PropertiesPlanType, "ATTIC VENT", StringComparison.OrdinalIgnoreCase) Then

                        AtticVentHeadlessPrinting.PrintAtticVent()

                        Exit Sub
                    ElseIf String.Equals(PropertiesPlanType, "MECHANICAL", StringComparison.OrdinalIgnoreCase) Then

                        MechanicalHeadlessPrinting.PrintMechanicalPlans()

                    End If

                End Using

            End Using

            Dim input As String = Stamps

            Dim SealLoop As List(Of String) = input.Split(","c).ToList()

            Dim StickAndTruss As Boolean = False

            Dim LayerLoop As New List(Of String)

            LayerLoop.Add("")

            Dim BuilderDivisionsList As List(Of String)
            If String.IsNullOrWhiteSpace(PropertiesPlanDivisions) Then
                BuilderDivisionsList = New List(Of String)()
            Else
                BuilderDivisionsList = PropertiesPlanDivisions.Split(","c).
        Select(Function(s) s.Trim()).
        Where(Function(s) s <> "").
        Distinct(StringComparer.OrdinalIgnoreCase).
        ToList()
            End If

            Dim TempElevs As New List(Of String)

            If BuilderDivisionsList.Count = 0 Then
                BuilderDivisionsList.Add("")
            End If

            Try

                Using acTrans3 As Transaction = accurdb.TransactionManager.StartTransaction()

                    Dim lytab As LayerTable = acTrans3.GetObject(accurdb.LayerTableId, OpenMode.ForRead)
                    Dim alllayers As New ArrayList
                    Dim fulllayerstring As String
                    Dim CurrentLayer As LayerTableRecord = acTrans3.GetObject(CurrentLayerID, OpenMode.ForRead)

                    accurdb.Clayer = lytab("0")


                    For Each layer In lytab

                        Dim lytr As LayerTableRecord = acTrans3.GetObject(layer, OpenMode.ForWrite)

                        If lytr.Name Like "S-ANNO-AUTOMATION" Then

                            lytr.IsPlottable = False

                        End If

                        fulllayerstring = "*S-SEAL-*"

                        If lytr.Name Like fulllayerstring Then
                            lytab.UpgradeOpen()
                            lytr.IsFrozen = True
                            lytr.IsOff = True
                        End If

                    Next

                    acTrans3.Commit()
                    acDoc.Editor.Regen()

                End Using

                ' Start a transaction
                Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                    ' Open the Block table for read
                    Dim blkTable As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)

                    ' Open the BlockTableRecord (ModelSpace) for read
                    Dim blkTableRec As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                    Dim order As String() = {Nothing, Nothing, "MOD", "SW", "PRNT", "SIZE", "OPTIONS", "ELEV", "", "LXS", "TND", "DIVISIONS"}

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

                                insertionPoint2 = blkRef.Position

                                Dim insertionPointStr As String = insertionPoint2.ToString()

                                valuesList.Add(blockobjecthandle)

                                Dim insertionExists As Boolean = AllValues.Any(Function(values) values(0) = insertionPointStr)

                                ' Only proceed if the insertion point is unique
                                'If Not insertionExists Then

                                valuesList.Add(insertionPoint2.ToString())

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

                                        ' Pick ONE: first-wins or last-wins
                                        ' First non-empty wins:
                                        If Not tagValues.ContainsKey(tag) Then tagValues(tag) = txt
                                        ' Last wins (use this instead of the line above):
                                        ' tagValues(tag) = txt

                                        ' Your extra per-tag lists:
                                        Select Case tag.ToUpperInvariant()
                                            Case "MOD"
                                                If Not PlanType.Contains(txt) Then PlanType.Add(txt)
                                            Case "ELEV"
                                                For Each v In txt.Split(","c).Select(Function(s) s.Trim()).Where(Function(s) s <> "")
                                                    If Not Elevations.Contains(v) Then Elevations.Add(v)
                                                Next

                                            Case "OPTIONS"
                                                If Not Options.Contains(txt) Then Options.Add(txt)
                                        End Select
                                    Next
                                End If

                                ' --- Merge into valuesList WITHOUT touching other indices ---
                                ' Make sure valuesList is big enough
                                While valuesList.Count < order.Length
                                    valuesList.Add(String.Empty)
                                End While

                                ' Only write the indices managed here; leave others as-is
                                For i As Integer = 0 To order.Length - 1
                                    Dim key = order(i)
                                    If key IsNot Nothing AndAlso tagValues.ContainsKey(key) Then
                                        valuesList(i) = tagValues(key)
                                    End If
                                Next
                                ' Add the individual values to the combined list
                                If valuesList.Count >= 7 Then
                                    AllValues.Add(valuesList)
                                End If
                            End If
                            'End If
                        End If
                    Next

                    ' Commit the transaction
                    acTrans.Commit()

                End Using

                Elevations.Sort()
                Dim NewInsertPoint As New List(Of List(Of String))
                Dim NewInsertPointList As New List(Of String)

                Dim insertionPoint1 As Point3d
                Dim LeftLayoutList As New List(Of List(Of String))
                Dim RightLayoutList As New List(Of List(Of String))
                Dim FinalLeftList As New List(Of List(Of String))
                Dim FinalRightList As New List(Of List(Of String))
                Dim counter As Integer
                counter = 1

                Dim NewFolderLocation As String

                Options.Add("")
                Swings.Add("Right")
                Swings.Add("Left")

                Dim SelectedFolder As String
                SelectedFolder = GetCurrentDwgFolder()
                If String.IsNullOrEmpty(SelectedFolder) Then Return

                If Options.Count = 1 Then Options.Add("")

                Dim SealToStamp As String

                For Each division In BuilderDivisionsList
                    Dim divisionfolder As String
                    If division <> "" Then
                        divisionfolder = "\" & division
                    Else
                        divisionfolder = ""

                    End If

                    Dim pdfdivision As String
                    If division <> "" Then
                        pdfdivision = " (" & division & ")"
                    Else
                        pdfdivision = ""
                    End If

                    For Each stamp In SealLoop

                        SealToStamp = "Master Seal File|S-SEAL-" & stamp
                        TurnOnOrOffLayer(SealToStamp, True)

                        For Each layer In LayerLoop

                            If layer = "S-FRM-TRUSS" Then
                                TurnOnOrOffLayer(layer, True)
                                TurnOnOrOffLayer("S-FRM-STICK", False)
                            ElseIf layer = "S-FRM-STICK" Then
                                TurnOnOrOffLayer(layer, True)
                                TurnOnOrOffLayer("S-FRM-TRUSS", False)
                            End If

                            For Each type In PlanType
                                For Each ElevValue In Elevations
                                    For Each opt In Options
                                        For Each valueList In AllValues

                                            'valueList(0) contains blockID
                                            'valueList(1) contains InsertionPoint
                                            'valueList(2) contains PlanType
                                            'valueList(3) contains Swing
                                            'valueList(4) contains Sequence
                                            'valueList(5) contains Scale
                                            'valueList(6) contains Option
                                            'valueList(7) contains Elevation
                                            'valuelist(8) inst set yet but is set as the single elevations when multiple


                                            If valueList(2) = type AndAlso valueList(7).Contains(ElevValue) AndAlso valueList(6) = opt Then

                                                If valueList(7).Contains(",") Then
                                                    Dim result As New List(Of String)
                                                    Dim parts() As String = valueList(7).Split(","c)
                                                    For Each part As String In parts
                                                        result.Add(part.Trim())
                                                    Next

                                                    If Not result.Contains(ElevValue.Trim(), StringComparer.OrdinalIgnoreCase) Then
                                                        Continue For
                                                    End If

                                                ElseIf valueList(7).Length > ElevValue.Length Then

                                                    Continue For

                                                End If

                                                'Perform the operation when both BeamLayout And FNDElev match
                                                Dim insertionPointStr As String = valueList(1) ' Example string from valueList

                                                insertionPoint1 = CreatePoint3dFromString(insertionPointStr)

                                                Dim NewHandle As String = valueList(0)
                                                Dim long1 As Long = NewHandle
                                                Dim hand As Handle = New Handle(long1)
                                                Dim objID As ObjectId = acDb.GetObjectId(False, hand, 0)



                                                If UCase(valueList(3)) = "L" Or UCase(valueList(3)) = "LEFT" Then
                                                    LeftLayoutList.Add(valueList)
                                                Else
                                                    RightLayoutList.Add(valueList)
                                                End If

                                            End If

                                        Next
                                        counter = 1

                                        For Each RightEntry In RightLayoutList
                                            If RightEntry(4) = counter Then
                                                If RightEntry.Count = 8 Then
                                                    RightEntry.Add(ElevValue)
                                                Else
                                                    RightEntry(8) = ElevValue
                                                End If
                                                FinalRightList.Add(RightEntry)
                                                counter += 1
                                                ZoomObjectsInViewport(RightEntry, Planname, Builder, FRSheets, division)
                                            End If
                                        Next

                                        Dim TodaysDate As String = Date.Today.ToString("MM dd yy", CultureInfo.InvariantCulture)

                                        NewFolderLocation = SelectedFolder & divisionfolder & "\" & stamp & "\" & type


                                        If FinalRightList.Count <> 0 Then


                                            If String.IsNullOrEmpty(SelectedFolder) Then Return

                                            If Not NetworkHelpers.IsNetworkPathAccessible(SelectedFolder) Then
                                                Return
                                            End If

                                            If Not NetworkHelpers.CreateDirectoryWithRetry(NewFolderLocation) Then
                                                Continue For ' Skip this iteration instead of crashing
                                            End If

                                            If UCase(type) = "FRAMING" Then

                                                pdfname = UCase("RIGHT FR " & Planname & " " & ElevValue)

                                            ElseIf UCase(type) = "BRACING" Then

                                                pdfname = UCase("RIGHT WB " & Planname & " " & ElevValue & " (WALLBRACING)")

                                                If opt = "115 MPH" Then
                                                    pdfname = pdfname & " 115 MPH"
                                                ElseIf opt = "130 MPH" Then
                                                    pdfname = pdfname & " 130 MPH"
                                                ElseIf opt = "142 MPH" Then
                                                    pdfname = pdfname & " 142 MPH"
                                                End If

                                            ElseIf UCase(type) = "WINDSTORM" Then

                                                pdfname = UCase("RIGHT WS " & Planname & " " & ElevValue & " (150 MPH)")

                                            End If

                                        End If

                                        If FinalRightList.Count <> 0 Then
                                            PlotTAutomatedTabs(FinalRightList, (pdfname & pdfdivision), NewFolderLocation)
                                            QueueLayoutsForCsv(FinalRightList, (pdfname & pdfdivision), Builder, Planname, PropertiesProject, BuilderDivision:=division)
                                        End If


                                        counter = 1
                                        For Each LeftEntry In LeftLayoutList
                                            If LeftEntry(4) = counter Then
                                                If LeftEntry.Count = 8 Then
                                                    LeftEntry.Add(ElevValue)
                                                Else
                                                    LeftEntry(8) = ElevValue
                                                End If
                                                FinalLeftList.Add(LeftEntry)
                                                counter += 1

                                                ZoomObjectsInViewport(LeftEntry, Planname, Builder, FRSheets, division)

                                            End If
                                        Next

                                        If FinalLeftList.Count <> 0 Then


                                            If String.IsNullOrEmpty(SelectedFolder) Then Return

                                            If Not NetworkHelpers.IsNetworkPathAccessible(SelectedFolder) Then
                                                Return
                                            End If

                                            If Not NetworkHelpers.CreateDirectoryWithRetry(NewFolderLocation) Then
                                                Continue For ' Skip this iteration instead of crashing
                                            End If

                                            If UCase(type) = "FRAMING" Then

                                                pdfname = UCase("LEFT FR " & Planname & " " & ElevValue)

                                            ElseIf UCase(type) = "BRACING" Then

                                                pdfname = UCase("LEFT WB " & Planname & " " & ElevValue & " (WALLBRACING)")

                                                If opt = "115 MPH" Then
                                                    pdfname = pdfname & " 115 MPH"
                                                ElseIf opt = "130 MPH" Then
                                                    pdfname = pdfname & " 130 MPH"
                                                ElseIf opt = "142 MPH" Then
                                                    pdfname = pdfname & " 142 MPH"
                                                End If

                                            ElseIf UCase(type) = "WINDSTORM" Then

                                                pdfname = UCase("LEFT WS " & Planname & " " & ElevValue & " (150 MPH)")

                                            End If

                                        End If

                                        If FinalLeftList.Count <> 0 Then
                                            PlotTAutomatedTabs(FinalLeftList, (pdfname & pdfdivision), NewFolderLocation)
                                            QueueLayoutsForCsv(FinalLeftList, (pdfname & pdfdivision), Builder, Planname, PropertiesProject, BuilderDivision:=division)
                                        End If

                                        FinalLeftList.Clear()
                                        FinalRightList.Clear()
                                        LeftLayoutList.Clear()
                                        RightLayoutList.Clear()
                                    Next
                                Next

                            Next



                        Next

                        TurnOnOrOffLayer(SealToStamp, False)

                    Next
                Next

                Dim lm As LayoutManager = LayoutManager.Current

                lm.CurrentLayout = "Model"
                SealLoop.Clear()
                LayerLoop.Clear()
                PlanType.Clear()
                Elevations.Clear()
                Options.Clear()

                Dim emittedCsv As String = FlushQueuedCsv(Builder, Planname)
                If Not String.IsNullOrEmpty(emittedCsv) Then
                    acEd.WriteMessage(vbLf & "CSV written: " & emittedCsv)
                End If


            Catch ex As System.Exception
                ' One place to report/log any failure
                LogError(ex, "TopLevel")
                ' (Optional) show a gentle UI message for overnight unattended runs

            End Try


        End Sub

        Public Sub SyncCustomPropsFromAutoPlanPrintInfo()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            If acDoc Is Nothing Then Exit Sub

            Dim db As Database = acDoc.Database
            Dim updatedProps As New List(Of String)
            Dim blockFound As Boolean = False

            Using acDoc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()

                    Dim bt As BlockTable = TryCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    If bt Is Nothing OrElse Not bt.Has(BlockTableRecord.ModelSpace) Then
                        tr.Commit()
                        Exit Sub
                    End If

                    Dim ms As BlockTableRecord = TryCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                    If ms Is Nothing Then
                        tr.Commit()
                        Exit Sub
                    End If

                    For Each id As ObjectId In ms
                        Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                        If ent Is Nothing Then Continue For

                        If TypeOf ent Is BlockReference Then
                            Dim br As BlockReference = DirectCast(ent, BlockReference)
                            Dim btr As BlockTableRecord = TryCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            If btr Is Nothing Then Continue For

                            If btr.Name.Equals("AutoPlanPrintInfo", StringComparison.OrdinalIgnoreCase) Then
                                blockFound = True

                                ' Collect attribute values by tag
                                Dim attrValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                For Each attId As ObjectId In br.AttributeCollection
                                    Dim ar As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                                    If ar Is Nothing Then Continue For
                                    Dim tag As String = If(ar.Tag, String.Empty).Trim()
                                    Dim val As String = If(ar.TextString, String.Empty).Trim()
                                    If tag.Length > 0 Then
                                        attrValues(tag) = val
                                    End If
                                Next

                                ' Helper to map attribute tags to custom property names we maintain
                                ' Uses "contains" to be tolerant of tag variants in the block
                                Dim tryUpdate As Action(Of String, String) =
                                Sub(tagContains As String, propName As String)
                                    Dim hit = attrValues.FirstOrDefault(Function(kv) kv.Key.IndexOf(tagContains, StringComparison.OrdinalIgnoreCase) >= 0)
                                    If Not String.IsNullOrEmpty(hit.Key) Then
                                        Dim attrVal As String = hit.Value
                                        Dim curVal As String = GetCustomDwgPropReliable(propName)
                                        If Not String.Equals(If(curVal, String.Empty).Trim(),
                                                             If(attrVal, String.Empty).Trim(),
                                                             StringComparison.Ordinal) Then
                                            SetCustomDwgPropReliable(propName, attrVal)
                                            updatedProps.Add(propName)
                                        End If
                                    End If
                                End Sub

                                ' Compare and update the standard set we use elsewhere in this form
                                tryUpdate("BUILDER", "BUILDER")
                                tryUpdate("PLAN", "PLAN")
                                tryUpdate("STAMPS", "STAMPS")
                                tryUpdate("PROJECT", "PROJECT NUMBER")
                                tryUpdate("SHEET", "SHEET LABELING")
                                tryUpdate("PACKAGE", "PackageFRWB")
                                tryUpdate("FRWB", "PackageFRWB")

                                Exit For ' Only the first matching block
                            End If
                        End If
                    Next

                    tr.Commit()
                End Using
            End Using

            ' Optional: uncomment to see a quick summary
            ' If blockFound AndAlso updatedProps.Count > 0 Then
            '     MessageBox.Show($"Updated properties: {String.Join(", ", updatedProps)}",
            '                     "AutoPlanPrintInfo Sync",
            '                     MessageBoxButtons.OK, MessageBoxIcon.Information)
            ' End If
        End Sub

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


        ' Minimal logger
        Private Sub LogError(ex As Exception, Optional whereTag As String = "")
            Dim tag = If(String.IsNullOrEmpty(whereTag), "", $" [{whereTag}]")
            'acEd.WriteMessage(vbLf & $"[OverNightPrinting ERROR]{tag}: {ex.Message}" & vbLf)
            Try
                Dim p = IO.Path.Combine(Environment.GetFolderPath(Module_Arcxis_TB.NetworkUNCPathForEgnyte & "\FS2\K\DPIS Drawings\ToPrint"), "OverNightPrinting.log")
                IO.File.AppendAllText(p, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {tag} {ex.ToString()}{Environment.NewLine}")
            Catch
                ' ignore
            End Try
        End Sub

        ' Set or update a custom DWG property
        Public Shared Sub ListPropsWithEnumerator()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim propsObj As Object = doc.Database.SummaryInfo.CustomProperties
            If propsObj Is Nothing Then
                ed.WriteMessage(vbLf & "No custom properties.")
                Return
            End If

            Dim en As IEnumerator = DirectCast(propsObj, IEnumerable).GetEnumerator()
            While en.MoveNext()
                Dim cur = en.Current
                If TypeOf cur Is DictionaryEntry Then
                    Dim de = DirectCast(cur, DictionaryEntry)
                    ed.WriteMessage(vbLf & $"{de.Key}: {de.Value}")
                Else
                    ' KeyValuePair(Of String,String) case
                    Dim t = cur.GetType()
                    Dim k As String = CStr(t.GetProperty("Key").GetValue(cur, Nothing))
                    Dim v As String = CStr(t.GetProperty("Value").GetValue(cur, Nothing))
                    ed.WriteMessage(vbLf & $"{k}: {v}")
                End If
            End While
        End Sub

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
        Private Sub ChanceBlocksAndLayers()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor

            Using acDoc.LockDocument()
                Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()

                    ' === 1. Delete all layouts except Modelspace ===
                    Dim lm As LayoutManager = LayoutManager.Current
                    Dim bt As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                    Dim dictLayouts As DBDictionary = acTrans.GetObject(acDb.LayoutDictionaryId, OpenMode.ForRead)

                    Dim layoutsToDelete As New List(Of String)
                    For Each entry As DBDictionaryEntry In dictLayouts
                        Dim loId As ObjectId = entry.Value
                        Dim lo As Layout = TryCast(acTrans.GetObject(loId, OpenMode.ForRead), Layout)
                        If lo IsNot Nothing AndAlso Not lo.ModelType Then
                            layoutsToDelete.Add(lo.LayoutName)
                        End If
                    Next

                    For Each layName In layoutsToDelete
                        Try
                            lm.DeleteLayout(layName)
                            ed.WriteMessage(vbLf & $"Deleted layout: {layName}")
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & $"Failed to delete layout {layName}: {ex.Message}")
                        End Try
                    Next

                    ' === 2. Erase all instances of specified blocks ===
                    Dim blkNames As String() = {"Arcxis Title Block", "DPIS RevisionBlock"}
                    For Each blkName In blkNames
                        If bt.Has(blkName) Then
                            Dim btr As BlockTableRecord = acTrans.GetObject(bt(blkName), OpenMode.ForRead)
                            Dim idsToErase As New List(Of ObjectId)

                            ' Collect all references to this block
                            For Each id As ObjectId In btr.GetBlockReferenceIds(True, True)
                                idsToErase.Add(id)
                            Next

                            ' Erase them
                            For Each id As ObjectId In idsToErase
                                Dim ent As Entity = TryCast(acTrans.GetObject(id, OpenMode.ForWrite), Entity)
                                If ent IsNot Nothing Then
                                    ent.Erase()
                                End If
                            Next
                        End If
                    Next

                    acTrans.Commit()
                End Using
            End Using

            ' === 3. Purge block definitions ===
            Using acDoc.LockDocument()
                Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                    Dim purgeNames As String() = {"Arcxis Title Block", "DPIS RevisionBlock"}

                    For Each blkName In purgeNames
                        If bt.Has(blkName) Then
                            Dim id As ObjectId = bt(blkName)
                            Dim ids As New ObjectIdCollection()
                            ids.Add(id)

                            acDb.Purge(ids)
                            If ids.Count > 0 Then
                                Dim btr As BlockTableRecord = acTrans.GetObject(ids(0), OpenMode.ForWrite)
                                Try
                                    btr.Erase()
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                vbLf & $"Purged block: {blkName}")
                                Catch ex As Exception
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                vbLf & $"Could not purge block {blkName}: {ex.Message}")
                                End Try
                            End If
                        End If
                    Next

                    acTrans.Commit()
                End Using
            End Using
        End Sub


        Private Sub CallCustomProps()

            Dim frm As New Form_FramingProperties
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database
            Dim SealLoop As New List(Of String)

            Dim StampLayers As New List(Of String)

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim lytab As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim cleanlayerstring As String

                For Each layer In lytab

                    Dim lytr As LayerTableRecord = acTrans.GetObject(layer, OpenMode.ForRead)

                    If lytr.Name Like "Master Seal File|S-SEAL-*" Then

                        cleanlayerstring = lytr.Name.Substring(24, lytr.Name.Length - 24)

                        StampLayers.Add(cleanlayerstring)

                    End If
                Next

                acTrans.Dispose()

            End Using


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
                frm.SealsList.Items.Add(seal)
            Next

            Dim Builder = GetCustomDwgPropReliable("BUILDER")
            Dim plan = GetCustomDwgPropReliable("PLAN")
            Dim Stamps = GetCustomDwgPropReliable("STAMPS")

            If Builder <> "" Then
                frm.TextBox1.Text = Builder
            End If

            If plan <> "" Then
                frm.TextBox2.Text = plan
            End If

            If Not String.IsNullOrWhiteSpace(Stamps) Then
                Dim tokens = Stamps.Split(","c).
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

            frm.TextBox1.Text = Builder
            frm.TextBox2.Text = plan


            frm.ShowDialog()

        End Sub

        <CommandMethod("CheckRectCrossings")>
        Public Shared Sub CheckRectCrossings()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim badLayouts As New List(Of String)()

            Using acLck As DocumentLock = doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Try
                        ' Operate in ModelSpace (user requested modelspace check)
                        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                        Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                        Dim rectIds As New List(Of ObjectId)()

                        ' Find candidate rectangles: closed polylines with ~17x11 size (paperspace working rectangle)
                        For Each id As ObjectId In btr
                            Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                            If ent Is Nothing Then Continue For
                            If TypeOf ent Is Polyline Then
                                Dim pl As Polyline = CType(ent, Polyline)
                                If Not pl.Closed Then Continue For
                                Try
                                    Dim ext As Extents3d = pl.GeometricExtents
                                    Dim w As Double = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X)
                                    Dim h As Double = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y)

                                    ' User-provided working rectangle: 84' x 118'-3 13/16" (~118.3177083333 ft)
                                    Dim expectedFeetW As Double = 118.3177083333
                                    Dim expectedFeetH As Double = 84.0

                                    ' Also consider model units may be inches (feet * 12)
                                    Dim expectedInchesW As Double = expectedFeetW * 12.0
                                    Dim expectedInchesH As Double = expectedFeetH * 12.0

                                    Dim matched As Boolean = False

                                    ' Tolerances: 1.0 ft for feet-based drawings, 6.0 in for inch-based drawings
                                    If (Math.Abs(w - expectedFeetW) < 1.0 AndAlso Math.Abs(h - expectedFeetH) < 1.0) OrElse
                               (Math.Abs(w - expectedFeetH) < 1.0 AndAlso Math.Abs(h - expectedFeetW) < 1.0) Then
                                        matched = True
                                    End If

                                    If Not matched Then
                                        If (Math.Abs(w - expectedInchesW) < 6.0 AndAlso Math.Abs(h - expectedInchesH) < 6.0) OrElse
                                   (Math.Abs(w - expectedInchesH) < 6.0 AndAlso Math.Abs(h - expectedInchesW) < 6.0) Then
                                            matched = True
                                        End If
                                    End If

                                    If matched Then
                                        rectIds.Add(id)
                                    End If
                                Catch
                                    ' ignore entities without extents
                                End Try
                            End If
                        Next

                        If rectIds.Count = 0 Then
                            ed.WriteMessage(vbLf & "CheckRectCrossings: no working rectangle found in ModelSpace.")
                            tr.Commit()
                            Return
                        End If

                        For Each rectId In rectIds
                            Dim rect As Polyline = TryCast(tr.GetObject(rectId, OpenMode.ForWrite), Polyline)
                            If rect Is Nothing Then Continue For

                            Dim crossingFound As Boolean = False

                            For Each otherId As ObjectId In btr
                                If otherId = rectId Then Continue For
                                Dim otherEnt As Entity = TryCast(tr.GetObject(otherId, OpenMode.ForRead), Entity)
                                If otherEnt Is Nothing Then Continue For

                                ' Exclude page block references
                                If otherId.ObjectClass.DxfName = "INSERT" Then
                                    Try
                                        Dim br As BlockReference = CType(otherEnt, BlockReference)
                                        Dim defRec As BlockTableRecord = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                        If defRec.Name.Equals("PAGE", StringComparison.OrdinalIgnoreCase) Then
                                            Continue For
                                        End If
                                    Catch
                                    End Try
                                End If

                                ' If both are curves, use geometric intersection
                                If TypeOf otherEnt Is Curve Then
                                    Try
                                        Dim curveRect As Curve = CType(rect, Curve)
                                        Dim curveOther As Curve = CType(otherEnt, Curve)
                                        Dim pts As New Point3dCollection()
                                        curveRect.IntersectWith(curveOther, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero)
                                        If pts.Count > 0 Then
                                            crossingFound = True
                                            Exit For
                                        End If
                                    Catch
                                        ' fall through to extents
                                    End Try
                                End If

                                ' Fallback: sample points on the other entity's extents and test containment
                                Try
                                    Dim ext1 As Extents3d = rect.GeometricExtents
                                    Dim ext2 As Extents3d = otherEnt.GeometricExtents

                                    ' Quick rejection: disjoint extents
                                    If (ext2.MaxPoint.X < ext1.MinPoint.X) OrElse (ext2.MinPoint.X > ext1.MaxPoint.X) OrElse (ext2.MaxPoint.Y < ext1.MinPoint.Y) OrElse (ext2.MinPoint.Y > ext1.MaxPoint.Y) Then
                                        ' disjoint, continue
                                    Else
                                        ' Build sample points from the other's extents: 4 corners, center, and mid-edges
                                        Dim samples As New List(Of Point3d)()
                                        Dim minP = ext2.MinPoint
                                        Dim maxP = ext2.MaxPoint
                                        samples.Add(New Point3d(minP.X, minP.Y, 0))
                                        samples.Add(New Point3d(minP.X, maxP.Y, 0))
                                        samples.Add(New Point3d(maxP.X, minP.Y, 0))
                                        samples.Add(New Point3d(maxP.X, maxP.Y, 0))
                                        samples.Add(New Point3d((minP.X + maxP.X) / 2.0, (minP.Y + maxP.Y) / 2.0, 0))
                                        samples.Add(New Point3d((minP.X + maxP.X) / 2.0, minP.Y, 0))
                                        samples.Add(New Point3d((minP.X + maxP.X) / 2.0, maxP.Y, 0))
                                        samples.Add(New Point3d(minP.X, (minP.Y + maxP.Y) / 2.0, 0))
                                        samples.Add(New Point3d(maxP.X, (minP.Y + maxP.Y) / 2.0, 0))

                                        Dim insideCount As Integer = 0
                                        For Each sPt In samples
                                            If IsPointInPolyline(rect, sPt) Then insideCount += 1
                                        Next

                                        If insideCount = 0 Then
                                            ' All sample points outside -> treat as outside (no crossing)
                                        ElseIf insideCount = samples.Count Then
                                            ' All sample points inside -> fully inside (no crossing)
                                        Else
                                            ' Some inside and some outside -> likely crossing
                                            crossingFound = True
                                            Exit For
                                        End If
                                    End If
                                Catch
                                    ' ignore entities without extents
                                End Try
                            Next

                            If crossingFound Then
                                rect.Color = Color.FromColorIndex(ColorMethod.ByAci, 1) ' Red
                            Else
                                rect.Color = Color.FromColorIndex(ColorMethod.ByLayer, 0) ' ByLayer
                            End If
                        Next

                        ' --- check each paperspace layout for exactly 2 model viewports (Number > 1) ---
                        Try
                            Dim layoutDict As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                            If layoutDict IsNot Nothing Then
                                For Each entry As DBDictionaryEntry In layoutDict
                                    Dim layout As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                                    If layout Is Nothing OrElse layout.ModelType Then Continue For

                                    Dim btrId As ObjectId = layout.BlockTableRecordId
                                    Dim btrRec As BlockTableRecord = TryCast(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                                    If btrRec Is Nothing Then Continue For

                                    Dim vpCount As Integer = 0
                                    For Each entId As ObjectId In btrRec
                                        If entId.ObjectClass.DxfName = "VIEWPORT" Then
                                            Dim vp As Viewport = TryCast(tr.GetObject(entId, OpenMode.ForRead), Viewport)
                                            If vp IsNot Nothing AndAlso vp.Number > 1 Then vpCount += 1
                                        End If
                                    Next

                                    If vpCount <> 2 Then
                                        badLayouts.Add(layout.LayoutName & " (" & vpCount.ToString() & " viewports)")
                                    End If
                                Next
                            End If
                        Catch
                            ' ignore layout inspection errors (best-effort)
                        End Try

                        tr.Commit()
                    Catch ex As Exception
                        ed.WriteMessage(vbLf & "CheckRectCrossings error: " & ex.Message)
                    End Try
                End Using
            End Using

            ' Show dialog if any paperspaces don't meet the viewport requirement
            If badLayouts.Count > 0 Then
                Dim msg As String = "Paperspaces missing exactly two model viewports:" & vbCrLf & String.Join(vbCrLf, badLayouts)
                MessageBox.Show(msg, "Viewport check", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ed.WriteMessage(vbLf & "CheckRectCrossings: completed")
        End Sub

        Private Shared Function IsPointInPolyline(poly As Polyline, pt As Point3d) As Boolean
            If poly Is Nothing OrElse Not poly.Closed Then Return False
            Dim count As Integer = poly.NumberOfVertices
            If count < 3 Then Return False

            Dim x As Double = pt.X
            Dim y As Double = pt.Y
            Dim inside As Boolean = False

            For i As Integer = 0 To count - 1
                Dim p1 As Point2d = poly.GetPoint2dAt(i)
                Dim j As Integer = (i + 1) Mod count
                Dim p2 As Point2d = poly.GetPoint2dAt(j)

                Dim xi As Double = p1.X
                Dim yi As Double = p1.Y
                Dim xj As Double = p2.X
                Dim yj As Double = p2.Y

                Dim intersect As Boolean = ((yi > y) <> (yj > y)) AndAlso (x < (xj - xi) * (y - yi) / (yj - yi + 0.0) + xi)
                If intersect Then inside = Not inside
            Next

            Return inside
        End Function
        ' ================== NEW QUEUING MODEL (one logical set per PDF) ==================
        ' We now capture exactly one metadata set per PDF (PLAN, ELEV, SW, MOD) and store
        ' rows as triples: Identifier | Key | Value.
        ' Identifier format: "Builder - PdfName"
        ' Keys: PLAN, ELEV, SW, MOD
        ' Value: extracted or empty.
        '
        ' This replaces the previous wide layout-row accumulation. Multiple calls for the
        ' same PDF will be ignored (first wins) to prevent duplicates.
        Public Shared Sub QueueLayoutsForCsv(layoutList As List(Of List(Of String)), pdfName As String, builder As String, planName As String, projectnumber As String, Optional ByVal IRC As String = "", Optional ByVal IECC As String = "", Optional ByVal MechCounty As String = "",
                                             Optional ByVal MechMan As String = "", Optional ByVal MechFuel As String = "", Optional ByVal MechPlan As Boolean = False, Optional ByVal DocType As String = "", Optional ByVal BuilderDivision As String = "")
            If String.IsNullOrWhiteSpace(pdfName) Then Exit Sub
            If layoutList Is Nothing OrElse layoutList.Count = 0 Then Exit Sub

            ' Build identifier ("Builder - PdfName")
            Dim identifier As String = (If(builder, "").Trim() & " - " & Path.GetFileNameWithoutExtension(If(pdfName, "").Trim())).Trim()

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
            End If

            If String.Equals(DocType, "ENERGY", StringComparison.OrdinalIgnoreCase) Then
                Community = If(first(9), "").Trim()
            End If


            swingRaw = NormalizeSwing(swingRaw)

            ' Store four logical rows: PLAN / ELEV / SW / MOD
            _pendingRows.Add(New List(Of String) From {identifier, "PLAN", planName})
            _pendingRows.Add(New List(Of String) From {identifier, "ELEV", elevation})
            _pendingRows.Add(New List(Of String) From {identifier, "SW", swingRaw})
            If BuilderDivision <> "" Then
                _pendingRows.Add(New List(Of String) From {identifier, "BUILDER DIVISION", BuilderDivision})
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
                _pendingRows.Add(New List(Of String) From {identifier, "IRC", IRC})
                _pendingRows.Add(New List(Of String) From {identifier, "IECC", IECC})
            End If
            If DocType.Equals("ENERGY", StringComparison.OrdinalIgnoreCase) Then
                _pendingRows.Add(New List(Of String) From {identifier, "COMMUNITY", Community})
            End If
        End Sub

        ' ================== FLUSH UPDATED FORMAT ==================
        ' Emits a single CSV aggregating all queued PDFs:
        ' Header: Identifier,Key,Value
        ' Each queued PDF contributes exactly four rows (PLAN/ELEV/SW/MOD).
        ' File name derived from first builder/plan found (sanitized); falls back to Combined.csv.

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
        Private Shared ReadOnly _publishedPdfs As New List(Of String)()
        Private Shared ReadOnly _publishedPdfsLock As New Object()

        ''' <summary>
        ''' Wait up to timeoutMs for the file to exist. Best-effort helper.
        ''' </summary>
        Private Shared Function WaitForFileExists(filePath As String, timeoutMs As Integer) As Boolean
            Try
                Dim sw As New System.Diagnostics.Stopwatch()
                sw.Start()
                While sw.ElapsedMilliseconds < timeoutMs
                    If File.Exists(filePath) Then
                        Return True
                    End If
                    System.Threading.Thread.Sleep(200)
                End While
                Return File.Exists(filePath)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Register a published PDF path (duplicates ignored).
        ''' Call this after publish succeeds.
        ''' </summary>
        Public Shared Sub RegisterPublishedPdf(pdfPath As String)
            If String.IsNullOrWhiteSpace(pdfPath) Then Return
            Try
                If Not File.Exists(pdfPath) Then Return
                SyncLock _publishedPdfsLock
                    If Not _publishedPdfs.Any(Function(p) String.Equals(p, pdfPath, StringComparison.OrdinalIgnoreCase)) Then
                        _publishedPdfs.Add(pdfPath)
                    End If
                End SyncLock
            Catch
                ' best-effort: swallow
            End Try
        End Sub

        ' --- Updated combine + watermark support ---
        ' Replace the existing CombineRegisteredPdfs, CombinePdfsCommand and AutoCombinePdfsCommand with these.
        Public Shared Function CombineRegisteredPdfs(outputPath As String, Optional watermark As String = Nothing, Optional fontName As String = "Arial", Optional fontSize As Double = 144, Optional opacity As Double = 0.15, Optional angle As Double = 52.35) As String
            ' Initialize font resolver if needed
            EnsureFontResolver()

            SyncLock _publishedPdfsLock
                If _publishedPdfs Is Nothing OrElse _publishedPdfs.Count = 0 Then Return Nothing
                Try
                    Dim outDoc As New PdfSharp.Pdf.PdfDocument()

                    For Each src In _publishedPdfs
                        Try
                            If Not File.Exists(src) Then Continue For
                            Using inp As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(src, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                                For Each pg As PdfSharp.Pdf.PdfPage In inp.Pages
                                    outDoc.AddPage(pg)
                                    ' If watermark requested, draw it onto the newly added page
                                    If Not String.IsNullOrWhiteSpace(watermark) Then
                                        Dim added As PdfSharp.Pdf.PdfPage = outDoc.Pages(outDoc.PageCount - 1)
                                        Using gfx As XGraphics = XGraphics.FromPdfPage(added, XGraphicsPdfPageOptions.Append)
                                            Dim pageW = added.Width.Point
                                            Dim pageH = added.Height.Point
                                            'angle = 240
                                            'fontSize = 240
                                            ' Prepare font and brush with alpha
                                            Dim font As New XFont(fontName, fontSize, XFontStyleEx.Regular)
                                            Dim col As XColor = XColor.FromArgb(CInt(255.0 * Math.Max(0.0, Math.Min(1.0, opacity))), XColors.Black)
                                            Dim brush As New XSolidBrush(col)

                                            ' Draw rotated centered watermark
                                            gfx.TranslateTransform(pageW / 2.0, pageH / 2.0)
                                            gfx.RotateTransform(angle)
                                            gfx.DrawString(watermark, font, brush, New XPoint(0, 0), XStringFormats.Center)
                                            gfx.RotateTransform(-angle)
                                            gfx.TranslateTransform(-pageW / 2.0, -pageH / 2.0)
                                        End Using
                                    End If
                                Next
                            End Using
                        Catch ex As System.InvalidOperationException
                            ' Log which file failed
                            Try
                                Dim doc = Application.DocumentManager.MdiActiveDocument
                                If doc IsNot Nothing Then
                                    doc.Editor.WriteMessage(vbLf & $"ERROR: Failed to process PDF: {src}")
                                    doc.Editor.WriteMessage(vbLf & $"  Exception: {ex.Message}")
                                End If
                            Catch
                            End Try
                            ' Continue with other files instead of failing entirely
                            Continue For
                        Catch ex As Exception
                            ' Log other errors too
                            Try
                                Dim doc = Application.DocumentManager.MdiActiveDocument
                                If doc IsNot Nothing Then
                                    doc.Editor.WriteMessage(vbLf & $"ERROR: Unexpected error with PDF: {src}")
                                    doc.Editor.WriteMessage(vbLf & $"  Exception: {ex.GetType().Name} - {ex.Message}")
                                End If
                            Catch
                            End Try
                            Continue For
                        Catch

                            ' skip single-source failures and continue
                        End Try
                        Continue For
                    Next

                    Dim dir = Path.GetDirectoryName(outputPath)
                    If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

                    outDoc.Save(outputPath)

                    ' Clear registered list after successful combine
                    ' _publishedPdfs.Clear()

                    Return outputPath
                Catch
                    Return Nothing
                End Try
            End SyncLock
        End Function

        <CommandMethod("CombinePdfs")>
        Public Sub CombinePdfsCommand(Optional filename As String = "")
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim ed As Editor = doc.Editor

            SyncLock _publishedPdfsLock
                If _publishedPdfs Is Nothing OrElse _publishedPdfs.Count = 0 Then
                    ed.WriteMessage(vbLf & "No published PDFs registered to combine.")
                    Return
                End If
            End SyncLock

            Dim owner As New WindowWrapper(Application.MainWindow.Handle)
            Using dlg As New SaveFileDialog()
                dlg.Filter = "PDF files (*.pdf)|*.pdf"
                dlg.Title = "Save combined PDF as"
                dlg.FileName = "Combined.pdf"

                Dim res = dlg.ShowDialog(owner)

                If res = DialogResult.OK Then
                    Dim out As String = dlg.FileName
                    Dim result As String = CombineRegisteredPdfs(out)
                    If Not String.IsNullOrEmpty(result) Then
                        ed.WriteMessage(vbLf & "Combined PDF written: " & result)
                    Else
                        ed.WriteMessage(vbLf & "Failed to create combined PDF.")
                    End If
                Else
                    ed.WriteMessage(vbLf & "Combine cancelled.")
                End If
            End Using
        End Sub

        Public Sub AutoCombinePdfsCommand(out As String, Optional watermark As String = Nothing)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim ed As Editor = doc.Editor

            SyncLock _publishedPdfsLock
                If _publishedPdfs Is Nothing OrElse _publishedPdfs.Count = 0 Then
                    ed.WriteMessage(vbLf & "No published PDFs registered to combine.")
                    Return
                End If
            End SyncLock

            watermark = "MASTER"
            Dim result As String = CombineRegisteredPdfs(out, "FOR REVIEW")
            If Not String.IsNullOrEmpty(result) Then
                ed.WriteMessage(vbLf & "Combined PDF written: " & result)
            Else
                ed.WriteMessage(vbLf & "Failed to create combined PDF.")
            End If
        End Sub

        ' --- Per-PDF metadata captured at publish time (no string parsing) ---
        Private Class PdfMeta
            Public Property Path As String
            Public Property ModName As String     ' layout row index 2 (e.g., FRAMING / BRACING / WINDSTORM)
            Public Property Elevation As String   ' single elevation (row index 8 preferred, else 7)
            Public Property Swing As String       ' row index 3 (L/R)
        End Class

        Private Shared ReadOnly _pdfMetaByPath As New Dictionary(Of String, PdfMeta)(StringComparer.OrdinalIgnoreCase)

        Private Shared Function ExtractSingleElevation(layouts As List(Of List(Of String))) As String
            If layouts Is Nothing OrElse layouts.Count = 0 Then Return Nothing

            ' Prefer single elevation at index 8
            Dim singles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row In layouts
                If row IsNot Nothing AndAlso row.Count > 8 Then
                    Dim v = If(row(8), Nothing)
                    If Not String.IsNullOrWhiteSpace(v) Then singles.Add(v.Trim())
                End If
            Next
            If singles.Count = 1 Then Return singles.First()
            If singles.Count > 1 Then Return String.Join(", ", singles)

            ' Fallback: ELEV(All) at index 7
            Dim alls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row In layouts
                If row IsNot Nothing AndAlso row.Count > 7 Then
                    Dim v = If(row(7), Nothing)
                    If Not String.IsNullOrWhiteSpace(v) Then alls.Add(v.Trim())
                End If
            Next
            If alls.Count = 1 Then Return alls.First()
            If alls.Count > 1 Then Return String.Join(", ", alls)

            Return Nothing
        End Function

        Public Shared Sub RegisterPublishedPdfMeta(pdfPath As String, layouts As List(Of List(Of String)))
            If String.IsNullOrWhiteSpace(pdfPath) OrElse Not IO.File.Exists(pdfPath) Then Return

            Dim modName As String = ""
            Dim swing As String = ""
            If layouts IsNot Nothing AndAlso layouts.Count > 0 Then
                Dim first = layouts(0)
                If first IsNot Nothing AndAlso first.Count > 3 Then
                    modName = If(first(2), "")
                    swing = If(first(3), "")
                End If
            End If
            Dim elev As String = ExtractSingleElevation(layouts)

            Dim meta As New PdfMeta With {
        .Path = pdfPath,
        .ModName = If(modName, ""),
        .Swing = If(swing, ""),
        .Elevation = If(elev, "")
    }

            SyncLock _publishedPdfsLock
                _pdfMetaByPath(pdfPath) = meta
            End SyncLock
        End Sub

        Private Shared Function SnapshotCurrentPublished() As List(Of String)
            SyncLock _publishedPdfsLock
                Return _publishedPdfs.ToList()
            End SyncLock
        End Function

        ' Simple combiner for an explicit list (does not touch the global registry)
        Private Shared Function CombineSpecificPdfs(sources As IEnumerable(Of String),
                                           outputPath As String,
                                           Optional watermark As String = Nothing,
                                           Optional fontName As String = "Arial",
                                           Optional fontSize As Double = 144,
                                           Optional opacity As Double = 0.15,
                                           Optional angle As Double = 232.35) As String
            Try
                Dim files = sources.
            Where(Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso IO.File.Exists(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
                If files.Count = 0 Then Return Nothing

                Dim outDoc As New PdfSharp.Pdf.PdfDocument()
                For Each src In files
                    Using inp As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(src, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                        For Each pg As PdfSharp.Pdf.PdfPage In inp.Pages
                            outDoc.AddPage(pg)

                            If Not String.IsNullOrWhiteSpace(watermark) Then
                                Dim added As PdfSharp.Pdf.PdfPage = outDoc.Pages(outDoc.PageCount - 1)
                                Using gfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(added, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                                    Dim pageW = added.Width.Point
                                    Dim pageH = added.Height.Point
                                    Dim font As New PdfSharp.Drawing.XFont(fontName, fontSize, PdfSharp.Drawing.XFontStyleEx.Bold)
                                    Dim col As PdfSharp.Drawing.XColor = PdfSharp.Drawing.XColor.FromArgb(CInt(255.0 * Math.Max(0.0, Math.Min(1.0, opacity))), PdfSharp.Drawing.XColors.Black)
                                    Dim brush As New PdfSharp.Drawing.XSolidBrush(col)

                                    gfx.TranslateTransform(pageW / 2.0, pageH / 2.0)
                                    gfx.RotateTransform(angle)
                                    gfx.DrawString(watermark, font, brush, New PdfSharp.Drawing.XPoint(0, 0), PdfSharp.Drawing.XStringFormats.Center)
                                    gfx.RotateTransform(-angle)
                                    gfx.TranslateTransform(-pageW / 2.0, -pageH / 2.0)
                                End Using
                            End If
                        Next
                    End Using
                Next

                Dim dir = IO.Path.GetDirectoryName(outputPath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
                outDoc.Save(outputPath)
                Return outputPath
            Catch
                Return Nothing
            End Try
        End Function

        ' Build a combined FR+WB list from the PDFs published for the current stamp (preserving publish order)
        Public Shared Function PackageFrWbFromCurrentPublished(masterFolderPath As String, planName As String) As String
            Dim current = SnapshotCurrentPublished()
            If current Is Nothing OrElse current.Count = 0 Then Return Nothing

            Dim frwb As New List(Of String)
            For Each p In current
                Dim meta As PdfMeta = Nothing
                SyncLock _publishedPdfsLock
                    _pdfMetaByPath.TryGetValue(p, meta)
                End SyncLock
                If meta Is Nothing Then Continue For
                If String.Equals(meta.ModName, "FRAMING", StringComparison.OrdinalIgnoreCase) _
           OrElse String.Equals(meta.ModName, "BRACING", StringComparison.OrdinalIgnoreCase) Then
                    frwb.Add(p)
                End If
            Next

            If frwb.Count = 0 Then Return Nothing
            Dim out = Path.Combine(masterFolderPath, planName & " FR+WB MASTER.pdf")
            Return CombineSpecificPdfs(frwb, out, "For Review")
        End Function

        ' Normalize swing values to LEFT/RIGHT
        Private Shared Function NormalizeSwing(s As String) As String
            Dim t As String = If(s, "").Trim().ToUpperInvariant()
            If t = "L" OrElse t = "LEFT" Then Return "LEFT"
            If t = "R" OrElse t = "RIGHT" Then Return "RIGHT"
            Return t
        End Function

        ' Overload that lets callers customize filename and (optionally) subfolder.
        Public Shared Function PackageFrWbPerElevationFromCurrentPublished(
            masterFolderPath As String,
            planName As String,
            nameTemplate As String,
            Optional subfolderName As String = Nothing
        ) As List(Of String)

            Dim outputs As New List(Of String)()
            Dim current As List(Of String) = SnapshotCurrentPublished()
            If current Is Nothing OrElse current.Count = 0 Then Return outputs

            ' Choose target output folder
            Dim targetFolder As String = If(String.IsNullOrWhiteSpace(subfolderName),
                                            masterFolderPath,
                                            Path.Combine(masterFolderPath, subfolderName))
            If Not Directory.Exists(targetFolder) Then Directory.CreateDirectory(targetFolder)

            ' Group in publish order by (Elevation, Swing)
            Dim groups As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

            For Each p In current
                Dim meta As PdfMeta = Nothing
                SyncLock _publishedPdfsLock
                    _pdfMetaByPath.TryGetValue(p, meta)
                End SyncLock
                If meta Is Nothing Then Continue For

                ' Added WINDSTORM to accepted module names
                If Not (String.Equals(meta.ModName, "FRAMING", StringComparison.OrdinalIgnoreCase) OrElse
                        String.Equals(meta.ModName, "BRACING", StringComparison.OrdinalIgnoreCase) OrElse
                        String.Equals(meta.ModName, "WINDSTORM", StringComparison.OrdinalIgnoreCase)) Then
                    Continue For
                End If

                Dim elev As String = If(meta.Elevation, "").Trim()
                Dim swing As String = NormalizeSwing(meta.Swing)
                If String.IsNullOrWhiteSpace(elev) OrElse String.IsNullOrWhiteSpace(swing) Then Continue For

                Dim key As String = elev & "||" & swing
                If Not groups.ContainsKey(key) Then groups(key) = New List(Of String)()
                groups(key).Add(p)
            Next

            For Each kvp In groups
                Dim key As String = kvp.Key
                Dim files As List(Of String) = kvp.Value

                Dim hasFr As Boolean = False
                Dim hasWb As Boolean = False
                Dim hasWs As Boolean = False ' new flag

                For Each p In files
                    Dim meta As PdfMeta = Nothing
                    SyncLock _publishedPdfsLock
                        _pdfMetaByPath.TryGetValue(p, meta)
                    End SyncLock
                    If meta Is Nothing Then Continue For
                    If String.Equals(meta.ModName, "FRAMING", StringComparison.OrdinalIgnoreCase) Then hasFr = True
                    If String.Equals(meta.ModName, "BRACING", StringComparison.OrdinalIgnoreCase) Then hasWb = True
                    If String.Equals(meta.ModName, "WINDSTORM", StringComparison.OrdinalIgnoreCase) Then hasWs = True
                Next

                ' Keep original requirement (must have both FRAMING and BRACING)
                If Not hasFr Then Continue For

                Dim parts = key.Split(New String() {"||"}, StringSplitOptions.None)
                Dim elev As String = parts(0)
                Dim swing As String = parts(1)

                Dim outBase As String =
                    nameTemplate.
                        Replace("{PLAN}", planName).
                        Replace("{ELEV}", SafeSegment(elev)).
                        Replace("{SWING}", swing).
                        ToUpperInvariant()

                Dim outPath As String = Path.Combine(targetFolder, outBase & ".pdf")
                Dim combined As String = CombineSpecificPdfs(files, outPath)
                If Not String.IsNullOrEmpty(combined) Then outputs.Add(combined)
            Next

            Return outputs
        End Function

        <CommandMethod("RefreshXrefLayers")>
        Public Shared Sub RefreshXrefLayers()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using acLck As DocumentLock = doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Try
                        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite)
                        Dim reloadedCount As Integer = 0
                        ' Collect all xrefs that need reloading
                        Dim xrefsToReload As New ObjectIdCollection()
                        For Each id As ObjectId In bt
                            Dim btr As BlockTableRecord = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                            If btr Is Nothing Then Continue For

                            ' Check if it's an xref (not overlay)
                            If btr.IsFromExternalReference Then
                                xrefsToReload.Add(btr.ObjectId)
                            End If
                        Next
                        ' Reload all xrefs
                        If xrefsToReload.Count > 0 Then
                            Try
                                db.ReloadXrefs(xrefsToReload)
                                reloadedCount = xrefsToReload.Count
                            Catch ex As Exception
                                'ed.WriteMessage(vbLf & $"Error reloading xrefs: {ex.Message}")
                            End Try
                        End If
                        ' Force layer table refresh
                        Dim lt As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                        tr.Commit()
                        ' Regen to update display
                        ed.Regen()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
        End Sub


        ' === Added: Excel driven layer import & merge ===
        <CommandMethod("ImportLayersFromExcel")>
        Public Sub ImportLayersFromExcel()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            If acDoc Is Nothing Then Exit Sub
            Dim db As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor

            ' Build the default path dynamically from the network UNC path
            Dim defaultPath As String = IO.Path.Combine(NetworkUNCPathForEgnyte, "Arcxis\Engineering\Drafting Standards\CAD Lisp Routines\BricsCad")

            ' Prompt for Excel file
            Dim ofd As New System.Windows.Forms.OpenFileDialog() With {
            .Title = "Select Excel file with layer definitions",
            .Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*",
            .Multiselect = False,
            .InitialDirectory = If(IO.Directory.Exists(defaultPath), defaultPath, "")
            }
            If ofd.ShowDialog() <> DialogResult.OK Then
                ed.WriteMessage(vbLf & "Import cancelled.")
                Exit Sub
            End If

            Dim excelPath As String = ofd.FileName
            If Not IO.File.Exists(excelPath) Then
                ed.WriteMessage(vbLf & "File not found: " & excelPath)
                Exit Sub
            End If

            Dim xlApp As Excel.Application = Nothing
            Dim xlWb As Excel.Workbook = Nothing
            Dim xlWs As Excel.Worksheet = Nothing

            ' Row data container
            Dim rows As New List(Of Dictionary(Of String, String))()

            Try
                xlApp = New Excel.Application()
                xlApp.DisplayAlerts = False
                xlWb = xlApp.Workbooks.Open(excelPath, [ReadOnly]:=True)
                xlWs = CType(xlWb.Sheets(1), Excel.Worksheet)

                Dim used = xlWs.UsedRange
                Dim rCount As Integer = used.Rows.Count
                Dim cCount As Integer = used.Columns.Count

                ' Expected columns:
                ' 1 LayerName
                ' 2 Color
                ' 3 Linetype
                ' 4 GlobalWidth (treated as lineweight mm)
                ' 5 HatchPattern
                ' 6 HatchScale
                ' 7 Plottable (Yes/No)
                ' 8 (unused/reserved)
                ' 9 MergeFrom (source layer to merge into LayerName if not blank)

                ' Detect header (optional) by first cell text
                Dim startRow As Integer = 1
                Dim firstCellVal = CStr((xlWs.Cells(1, 1).Value))
                If Not String.IsNullOrWhiteSpace(firstCellVal) AndAlso firstCellVal.Trim().ToUpperInvariant().Contains("LAYER") Then
                    startRow = 2
                End If

                For r = startRow To rCount
                    Dim data As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For c = 1 To Math.Min(9, cCount)
                        Dim rawObj = xlWs.Cells(r, c).Value
                        Dim raw As String = If(rawObj Is Nothing, "", CStr(rawObj)).Trim()
                        data("C" & c) = raw
                    Next
                    If data.ContainsKey("C1") AndAlso Not String.IsNullOrWhiteSpace(data("C1")) Then
                        rows.Add(data)
                    End If
                Next

            Catch ex As Exception
                ed.WriteMessage(vbLf & "Excel read error: " & ex.Message)
                Exit Sub
            Finally
                ' Release COM
                If xlWs IsNot Nothing Then Marshal.ReleaseComObject(xlWs)
                If xlWb IsNot Nothing Then
                    xlWb.Close(False)
                    Marshal.ReleaseComObject(xlWb)
                End If
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    Marshal.ReleaseComObject(xlApp)
                End If
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

            If rows.Count = 0 Then
                ed.WriteMessage(vbLf & "No layer rows found.")
                Exit Sub
            End If

            Using acDoc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim layerTable As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
                    Dim ltypeTable As LinetypeTable = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)

                    ' Pass 1: Create/Update layers
                    For Each row In rows
                        Dim layerName = row("C1")
                        If String.IsNullOrWhiteSpace(layerName) Then Continue For

                        Dim colorSpec = If(row.ContainsKey("C2"), row("C2"), "")
                        Dim ltypeName = If(row.ContainsKey("C3"), row("C3"), "")
                        Dim widthSpec = If(row.ContainsKey("C4"), row("C4"), "")
                        Dim hatchPattern = If(row.ContainsKey("C5"), row("C5"), "")
                        Dim hatchScale = If(row.ContainsKey("C6"), row("C6"), "")
                        Dim plottableSpec = If(row.ContainsKey("C7"), row("C7"), "")

                        Dim ltrId As ObjectId
                        Dim ltr As LayerTableRecord

                        If layerTable.Has(layerName) Then
                            ltrId = layerTable(layerName)
                            ltr = CType(tr.GetObject(ltrId, OpenMode.ForWrite), LayerTableRecord)
                        Else
                            layerTable.UpgradeOpen()
                            ltr = New LayerTableRecord() With {.Name = layerName}
                            ltrId = layerTable.Add(ltr)
                            tr.AddNewlyCreatedDBObject(ltr, True)
                            layerTable.DowngradeOpen()
                        End If

                        ' Color
                        If Not String.IsNullOrWhiteSpace(colorSpec) Then
                            Dim acColor As Color = ParseColor(colorSpec)
                            If acColor IsNot Nothing Then
                                ltr.Color = acColor
                            End If
                        End If

                        ' Linetype
                        If Not String.IsNullOrWhiteSpace(ltypeName) Then
                            If Not ltypeTable.Has(ltypeName) Then
                                Try
                                    db.LoadLineTypeFile(ltypeName, "acad.lin")
                                Catch
                                    Try
                                        db.LoadLineTypeFile(ltypeName, "iso.lin")
                                    Catch
                                        ed.WriteMessage(vbLf & "Could not load linetype: " & ltypeName)
                                    End Try
                                End Try
                            End If
                            If ltypeTable.Has(ltypeName) Then
                                ltr.LinetypeObjectId = ltypeTable(ltypeName)
                            End If
                        End If

                        ' Lineweight (Global Width spec mapped)
                        If Not String.IsNullOrWhiteSpace(widthSpec) Then
                            Dim lw = ParseLineweight(widthSpec)
                            If lw.HasValue Then ltr.LineWeight = lw.Value
                        End If

                        ' Plottable
                        If Not String.IsNullOrWhiteSpace(plottableSpec) Then
                            Dim yes = plottableSpec.Trim().ToUpperInvariant()
                            ltr.IsPlottable = (yes = "YES" OrElse yes = "Y" OrElse yes = "TRUE" OrElse yes = "1")
                        End If

                        ' Hatch pattern/scale cannot be stored on a layer; optionally register in XData
                        If Not String.IsNullOrWhiteSpace(hatchPattern) OrElse Not String.IsNullOrWhiteSpace(hatchScale) Then
                            Try
                                ' Simple XData storage under app name "LAYER_HATCH_META"
                                Dim regAppTbl As RegAppTable = CType(tr.GetObject(db.RegAppTableId, OpenMode.ForRead), RegAppTable)
                                Const appName = "LAYER_HATCH_META"
                                If Not regAppTbl.Has(appName) Then
                                    regAppTbl.UpgradeOpen()
                                    Dim appRec As New RegAppTableRecord() With {.Name = appName}
                                    regAppTbl.Add(appRec)
                                    tr.AddNewlyCreatedDBObject(appRec, True)
                                    regAppTbl.DowngradeOpen()
                                End If

                                Dim rb As New ResultBuffer(
                                    New TypedValue(DxfCode.ExtendedDataRegAppName, appName),
                                    New TypedValue(DxfCode.ExtendedDataAsciiString, "PAT=" & hatchPattern),
                                    New TypedValue(DxfCode.ExtendedDataAsciiString, "SCALE=" & hatchScale)
                                )
                                ltr.XData = rb
                            Catch
                                ' Ignore XData errors
                            End Try
                        End If
                    Next

                    ' Pass 2: Merge operations (Column 9 -> Column 1)
                    ' Build layer table fresh after potential additions
                    layerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

                    For Each row In rows
                        Dim targetLayer = row("C1")
                        Dim sourceLayer = If(row.ContainsKey("C9"), row("C9"), "")
                        If String.IsNullOrWhiteSpace(sourceLayer) Then Continue For
                        If String.Equals(sourceLayer, targetLayer, StringComparison.OrdinalIgnoreCase) Then Continue For
                        If Not layerTable.Has(targetLayer) Then
                            ed.WriteMessage(vbLf & $"Target layer missing (skipping merge): {targetLayer}")
                            Continue For
                        End If
                        If Not layerTable.Has(sourceLayer) Then
                            ed.WriteMessage(vbLf & $"Source layer missing (skipping merge): {sourceLayer}")
                            Continue For
                        End If

                        Dim targetId = layerTable(targetLayer)
                        Dim sourceId = layerTable(sourceLayer)

                        ' Reassign entities from source to target
                        Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        For Each btrId As ObjectId In bt
                            Dim btr As BlockTableRecord = CType(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                            For Each entId As ObjectId In btr
                                Try
                                    Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForWrite), Entity)
                                    If ent Is Nothing Then Continue For
                                    If ent.LayerId = sourceId Then
                                        ent.LayerId = targetId
                                    End If
                                Catch ex As Teigha.Runtime.Exception When ex.ErrorStatus = ErrorStatus.OnLockedLayer
                                    Dim innerEnt As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                                    If innerEnt IsNot Nothing Then
                                        Dim ltr As LayerTableRecord = CType(tr.GetObject(innerEnt.LayerId, OpenMode.ForRead), LayerTableRecord)
                                        Dim msg As New System.Text.StringBuilder()
                                        msg.AppendLine($"Layer ""{ltr.Name}"" is preventing this operation.")
                                        msg.AppendLine("To fix this, please do one of the following:")
                                        If ltr.IsLocked Then msg.AppendLine("  • Unlock all layers before retrying.")
                                        If ltr.IsFrozen Then msg.AppendLine("  • Thaw all layers before retrying.")
                                        If ltr.IsOff Then msg.AppendLine("  • Turn all layers on before retrying.")
                                        If ltr.IsDependent Then msg.AppendLine("  • This layer belongs to an external reference (xref) and cannot be modified directly.")
                                        Bricscad.ApplicationServices.Application.ShowAlertDialog(msg.ToString())
                                    End If
                                    Exit Sub
                                Catch ex As System.Exception
                                    Bricscad.ApplicationServices.Application.ShowAlertDialog($"Unexpected error processing entity: {ex.Message}")
                                    Exit Sub
                                End Try
                            Next
                        Next

                        ' Attempt to purge & erase source layer
                        Try
                            layerTable.UpgradeOpen()
                            Dim ids As New ObjectIdCollection()
                            ids.Add(sourceId)
                            db.Purge(ids)
                            If ids.Count > 0 Then
                                Dim srcRec As LayerTableRecord = CType(tr.GetObject(ids(0), OpenMode.ForWrite), LayerTableRecord)
                                srcRec.Erase()
                                ed.WriteMessage(vbLf & $"Merged & removed layer '{sourceLayer}' into '{targetLayer}'.")
                            Else
                                ed.WriteMessage(vbLf & $"Merged layer '{sourceLayer}' into '{targetLayer}' (could not purge).")
                            End If
                            layerTable.DowngradeOpen()
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & $"Merge purge failed for '{sourceLayer}': {ex.Message}")
                        End Try
                    Next

                    tr.Commit()
                End Using
            End Using

            ed.WriteMessage(vbLf & "Layer import & merge complete.")
        End Sub

        ' Parse color specifications: ACI number, name, or R,G,B
        Private Function ParseColor(spec As String) As Color
            Try
                If String.IsNullOrWhiteSpace(spec) Then Return Nothing
                Dim s = spec.Trim()

                ' RGB form
                If s.Contains(",") Then
                    Dim parts = s.Split(","c).Select(Function(x) x.Trim()).ToArray()
                    If parts.Length = 3 Then
                        Dim r = CInt(Val(parts(0)))
                        Dim g = CInt(Val(parts(1)))
                        Dim b = CInt(Val(parts(2)))
                        Return Color.FromRgb(CByte(Math.Max(0, Math.Min(255, r))),
                                             CByte(Math.Max(0, Math.Min(255, g))),
                                             CByte(Math.Max(0, Math.Min(255, b))))
                    End If
                End If

                ' Numeric ACI
                Dim num As Integer
                If Integer.TryParse(s, num) AndAlso num >= 1 AndAlso num <= 255 Then
                    Return Color.FromColorIndex(ColorMethod.ByAci, CShort(num))
                End If

                ' Named basic colors
                Select Case s.ToUpperInvariant()
                    Case "RED" : Return Color.FromColorIndex(ColorMethod.ByAci, 1)
                    Case "YELLOW" : Return Color.FromColorIndex(ColorMethod.ByAci, 2)
                    Case "GREEN" : Return Color.FromColorIndex(ColorMethod.ByAci, 3)
                    Case "CYAN" : Return Color.FromColorIndex(ColorMethod.ByAci, 4)
                    Case "BLUE" : Return Color.FromColorIndex(ColorMethod.ByAci, 5)
                    Case "MAGENTA" : Return Color.FromColorIndex(ColorMethod.ByAci, 6)
                    Case "WHITE" : Return Color.FromColorIndex(ColorMethod.ByAci, 7)
                End Select
            Catch
                ' Ignore
            End Try
            Return Nothing
        End Function

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
                            ElseIf String.Equals(planType, "Attic Vent", StringComparison.OrdinalIgnoreCase) Then
                                customCTB = "ARCXIS"
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
        Private Function ParseLineweight(spec As String) As LineWeight?
            If String.IsNullOrWhiteSpace(spec) Then Return Nothing
            Dim s = spec.Trim().Replace("mm", "").Replace("MM", "")
            Dim val As Double
            If Not Double.TryParse(s, Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, val) Then
                Return Nothing
            End If
            ' AutoCAD LineWeight enum discrete values (hundredths of mm internally)
            Dim candidates As New Dictionary(Of LineWeight, Double) From {
                {LineWeight.LineWeight000, 0.00},
                {LineWeight.LineWeight005, 0.05},
                {LineWeight.LineWeight009, 0.09},
                {LineWeight.LineWeight013, 0.13},
                {LineWeight.LineWeight015, 0.15},
                {LineWeight.LineWeight018, 0.18},
                {LineWeight.LineWeight020, 0.2},
                {LineWeight.LineWeight025, 0.25},
                {LineWeight.LineWeight030, 0.3},
                {LineWeight.LineWeight035, 0.35},
                {LineWeight.LineWeight040, 0.4},
                {LineWeight.LineWeight050, 0.5},
                {LineWeight.LineWeight053, 0.53},
                {LineWeight.LineWeight060, 0.6},
                {LineWeight.LineWeight070, 0.7},
                {LineWeight.LineWeight080, 0.8},
                {LineWeight.LineWeight090, 0.9},
                {LineWeight.LineWeight100, 1.0},
                {LineWeight.LineWeight106, 1.06},
                {LineWeight.LineWeight120, 1.2},
                {LineWeight.LineWeight140, 1.4},
                {LineWeight.LineWeight158, 1.58},
                {LineWeight.LineWeight200, 2.0},
                {LineWeight.LineWeight211, 2.11}
            }
            Dim best = candidates.OrderBy(Function(kv) Math.Abs(kv.Value - val)).First()
            Return best.Key
        End Function
        ' === End Added ===
    End Class
End Namespace