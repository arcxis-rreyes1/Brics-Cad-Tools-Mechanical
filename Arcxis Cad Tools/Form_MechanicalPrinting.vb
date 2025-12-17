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
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
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
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports Document = Autodesk.AutoCAD.ApplicationServices.Document
Imports Excel = Microsoft.Office.Interop.Excel
Imports Exception = Autodesk.AutoCAD.Runtime.Exception
Imports Layout = Autodesk.AutoCAD.DatabaseServices.Layout
Imports Path = System.IO.Path
Imports Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline
Imports Viewport = Autodesk.AutoCAD.DatabaseServices.Viewport

Public Class Form_MechanicalPrinting

    ' === CSV accumulation (list of lists) ===
    Private Shared _pendingRows As New List(Of List(Of String))()
    Private Sub Form_MechanicalPrinting_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Dim FormBuilder = GetCustomDwgPropReliable("BUILDER")
        Dim Formplan = GetCustomDwgPropReliable("PLAN")

        Dim GasType As New List(Of String)
        Dim Manufacturers As New List(Of String)
        Dim Swings As New List(Of String)
        Dim Elevations As New List(Of String)
        Dim Counties As New List(Of String)

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
                    If blkDef.Name = "MechAutoPage" Then

                        Dim blockobjecthandle As String = blkRef.Handle.Value.ToString()


                        ' Check if the block reference has attributes
                        If blkRef.AttributeCollection.Count > 0 Then
                            ' Iterate through the attributes
                            For Each attId As ObjectId In blkRef.AttributeCollection
                                Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForRead), AttributeReference)

                                ' Check the attribute tag and add the value to the appropriate list
                                If attRef IsNot Nothing Then
                                    Select Case attRef.Tag.ToUpper()

                                        Case "GASTYPE"

                                            If Not GasType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                GasType.Add(attRef.TextString)
                                            End If

                                        Case "MANUFACTURER"

                                            If Not Manufacturers.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                Manufacturers.Add(attRef.TextString)
                                            End If

                                        Case "ELEV"

                                            If attRef.TextString.Contains(",") Then
                                                Dim values() As String = attRef.TextString.Split(","c) ' Split the string into an array of values

                                                For Each value As String In values

                                                    If Not Elevations.Contains(value) Then

                                                        Elevations.Add(value)

                                                    End If ' Add each value to the TempElev list, trimming any extra spaces

                                                Next

                                            Else

                                                If Not Elevations.Contains(attRef.TextString) Then

                                                    Elevations.Add(attRef.TextString)

                                                End If

                                            End If

                                        Case "SW"

                                            If Swings.Contains(attRef.TextString) Then

                                                Swings.Add(attRef.TextString)

                                            End If

                                        Case "COUNTY"

                                            If Not Counties.Contains(attRef.TextString) Then

                                                Counties.Add(attRef.TextString)

                                            End If

                                    End Select
                                End If
                            Next
                        End If
                    End If
                    'End If
                End If
            Next


            Elevations.Sort()
            Swings.Sort()
            Counties.Sort()
            GasType.Sort()
            Manufacturers.Sort()
        End Using

        For Each item In GasType
            GasList.Items.Add(item)
        Next

        For Each item In Manufacturers
            ManList.Items.Add(item)
        Next

        For Each item In Counties
            CountyList.Items.Add(item)
        Next

        GasList.SetItemChecked(0, True) '''''Options
        ManList.SetItemChecked(0, True) '''''Manufacturers\
        CountyList.SetItemChecked(0, True) '''''Counties


        TextBox1.Text = FormBuilder
        TextBox2.Text = Formplan

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

    Private Sub CheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs) _
    Handles ManList.ItemCheck, GasList.ItemCheck, CountyList.ItemCheck

        Me.BeginInvoke(Sub()
                           Dim clb = DirectCast(sender, CheckedListBox)
                           If clb.Items.Count = 0 Then Return

                           ' Build post-click states (what the checks will be after this click)
                           Dim states As Boolean() = Enumerable.Range(0, clb.Items.Count) _
            .Select(Function(i As Integer)
                        If i = e.Index Then
                            Return e.NewValue = CheckState.Checked
                        Else
                            Return clb.GetItemChecked(i)
                        End If
                    End Function) _
            .ToArray()

                           ' 1) If user is CHECKING item(0), make it the only checked and select it.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Checked Then
                               For i As Integer = 0 To clb.Items.Count - 1
                                   clb.SetItemChecked(i, i = 0)
                               Next
                               clb.SelectedIndex = 0
                               Return
                           End If

                           ' 2) If user is UNCHECKING item(0), clear its highlight.
                           If e.Index = 0 AndAlso e.NewValue = CheckState.Unchecked Then
                               clb.ClearSelected()
                               ' continue—no other action needed
                           End If

                           ' 3) If a non-zero item is being checked while item(0) is (now) checked, uncheck item(0).
                           If e.Index <> 0 AndAlso e.NewValue = CheckState.Checked AndAlso (states(0) OrElse clb.GetItemChecked(0)) Then
                               clb.SetItemChecked(0, False)
                               clb.SelectedIndex = e.Index  ' highlight the newly-checked item
                               states(0) = False             ' keep local states consistent
                           End If

                           ' 4) If all non-zero items are checked, collapse to only item(0).
                           If clb.Items.Count > 1 Then
                               Dim allNonZeroChecked As Boolean =
                Enumerable.Range(1, clb.Items.Count - 1) _
                          .All(Function(i As Integer)
                                   ' use states (post-click) for accuracy
                                   Return states(i)
                               End Function)

                               If allNonZeroChecked Then
                                   For i As Integer = 0 To clb.Items.Count - 1
                                       clb.SetItemChecked(i, i = 0)
                                   Next
                                   clb.SelectedIndex = 0
                               End If
                           End If
                       End Sub)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Hide()

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Dim GasType As New List(Of String)
        Dim Swings As New List(Of String)
        Dim Elevations As New List(Of String)
        Dim Counties As New List(Of String)
        Dim Manufacturers As New List(Of String)
        Dim AllValues As New List(Of List(Of String))
        Dim insertionPoint2 As Point3d
        Dim pdfname As String = ""
        Dim planname As String = ""
        Dim CurrentLayerID As ObjectId = acCurDb.Clayer
        Dim Builder As String = ""
        Dim ProjectNumber As String = ""
        Dim CustomPrinting As Boolean = False

        Builder = UCase(TextBox1.Text)

        planname = UCase(TextBox2.Text)

        SetCustomDwgPropReliable("BUILDER", Builder)

        SetCustomDwgPropReliable("PLAN", planname)


        If Not ManList.CheckedItems.Contains("All") Then
            CustomPrinting = True
        End If
        If Not GasList.CheckedItems.Contains("All") Then
            CustomPrinting = True
        End If

        Dim TempElevs As New List(Of String)


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
            Dim order As String() = {Nothing, Nothing, "PLANTYPE", "SW", "PRNT", "SIZE", "COUNTY", "ELEV", "", "GASTYPE", "MANUFACTURER"}

            ' Iterate through the ModelSpace block table record
            For Each objId As ObjectId In blkTableRec

                Dim entity As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                ' Check if the entity is a block reference
                If TypeOf entity Is BlockReference Then

                    Dim blkRef As BlockReference = CType(entity, BlockReference)

                    ' Get the block table record for the block reference
                    Dim blkDef As BlockTableRecord = TryCast(acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                    ' Check if the block name is "PAGE"
                    If blkDef.Name = "MechAutoPage" Then

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
                                    Case "GASTYPE"

                                        If Not GasType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                            GasType.Add(attRef.TextString)
                                        End If

                                    Case "MANUFACTURER"

                                        If Not Manufacturers.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                            Manufacturers.Add(attRef.TextString)
                                        End If

                                    Case "ELEV"

                                        If attRef.TextString.Contains(",") Then
                                            Dim values() As String = attRef.TextString.Split(","c) ' Split the string into an array of values

                                            For Each value As String In values

                                                If Not Elevations.Contains(value) Then

                                                    Elevations.Add(value)

                                                End If ' Add each value to the TempElev list, trimming any extra spaces

                                            Next

                                        Else

                                            If Not Elevations.Contains(attRef.TextString) Then

                                                Elevations.Add(attRef.TextString)

                                            End If

                                        End If

                                    Case "SW"

                                        If Swings.Contains(attRef.TextString) Then

                                            Swings.Add(attRef.TextString)

                                        End If

                                    Case "COUNTY"

                                        If Not Counties.Contains(attRef.TextString) Then

                                            Counties.Add(attRef.TextString)

                                        End If
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

        If CustomPrinting Then

            ' Reset
            GasType.Clear()
            Elevations.Clear()
            Swings.Clear()
            Counties.Clear()
            Manufacturers.Clear()

            ' Vent Types from CheckedListBox1 (All = every item except "All")
            If ManList.SelectedItem IsNot Nothing AndAlso ManList.SelectedItem.ToString() = "All" Then
                For Each selecteditem In ManList.Items
                    If Not Object.Equals(selecteditem, "All") Then
                        Manufacturers.Add(selecteditem)
                    End If
                Next
            Else
                For Each selecteditem In ManList.CheckedItems
                    Manufacturers.Add(selecteditem)
                Next
            End If

            ' IRC from CheckedListBox2 (All = every item except "All")
            If CountyList.SelectedItem IsNot Nothing AndAlso CountyList.SelectedItem.ToString() = "All" Then
                For Each selecteditem In CountyList.Items
                    If Not Object.Equals(selecteditem, "All") Then
                        Counties.Add(selecteditem)
                    End If
                Next
            Else
                For Each selecteditem In CountyList.CheckedItems
                    Counties.Add(selecteditem)
                Next
            End If

            ' Options from OptionsList ("All" means no specific option filter -> blank)
            If GasList.CheckedItems.Contains("All") Then
                GasType.Add("") ' blank represents all options
            Else
                For Each selecteditem In GasList.CheckedItems
                    GasType.Add(selecteditem)
                Next
            End If

        End If

        Dim SelectedFolder As String
        Dim isFile As Boolean
        SelectedFolder = PromptForPathOrFolder(isFile)
        If String.IsNullOrEmpty(SelectedFolder) Then Return

        Dim TodaysDate As String = Date.Today.ToString("MM dd yy", CultureInfo.InvariantCulture)
        Dim SealToStamp As String = "Master Seal File|S-SEAL-TML-TX"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''add seal'''''''''''''''''
        Dim basefolderlocation As String = SelectedFolder & "\" & planname
        TurnOnOrOffLayer(SealToStamp, True)
        Dim CounterSkip As Boolean = False
        For Each Manufacturer In Manufacturers

            NewFolderLocation = basefolderlocation & "\" & Manufacturer

            If Not Directory.Exists(NewFolderLocation) Then

                Directory.CreateDirectory(NewFolderLocation)

            End If
            For Each fuel In GasType

                NewFolderLocation = basefolderlocation & "\" & Manufacturer & "\" & fuel

                If Not Directory.Exists(NewFolderLocation) Then

                    Directory.CreateDirectory(NewFolderLocation)

                End If

                For Each county In Counties

                    For Each ElevValue In Elevations
                        For Each valueList In AllValues

                            'valueList(0) contains blockID
                            'valueList(1) contains InsertionPoint
                            'valueList(2) contains PlanType
                            'valueList(3) contains Swing
                            'valueList(4) contains Sequence
                            'valueList(5) contains Scale
                            'valueList(6) contains County
                            'valueList(7) contains Elevation
                            'valuelist(8) inst set yet but is set as the single elevations when multiple
                            'valueList(9) contains GasType
                            'valueList(10) contains Manufacturer


                            If valueList(9) = fuel AndAlso valueList(10) = Manufacturer AndAlso valueList(7).Contains(ElevValue) AndAlso valueList(6) = county Then

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

                                RightLayoutList.Add(valueList)

                            End If

                        Next

                        counter = 1
                        For Each RightEntry In RightLayoutList

                            If RightEntry(4) = counter Then
                                RightEntry(8) = ElevValue
                                FinalRightList.Add(RightEntry)
                                counter += 1
                                ZoomObjectsInViewport(RightEntry, planname, Builder, False)
                            End If

                        Next

                        If FinalRightList.Count <> 0 Then

                            pdfname = UCase("MECH " & fuel & " - " & planname & " - " & county & " - Mechanical Design - " & Manufacturer)

                        End If

                        ' Right side
                        If FinalRightList.Count <> 0 Then
                            PlotTAutomatedTabs(FinalRightList, pdfname, NewFolderLocation)
                            QueueLayoutsForCsv(FinalRightList, pdfname, Builder, planname, ProjectNumber)
                        End If

                        FinalRightList.Clear()
                        RightLayoutList.Clear()
                    Next
                Next
            Next

        Next

        Dim masterfolderpath As String = basefolderlocation
        Dim out = Path.Combine(masterfolderpath, planname & " - HVAC.pdf")
        Dim res = CombineRegisteredPdfs(out, "For Review")
        If res IsNot Nothing Then
            acEd.WriteMessage(vbLf & "Combined PDF written: " & res)
        Else
            acEd.WriteMessage(vbLf & "Failed to create combined PDF.")
        End If

        TurnOnOrOffLayer(SealToStamp, False)
        Dim lm As LayoutManager = LayoutManager.Current

        lm.CurrentLayout = "Model"
        GasType.Clear()
        Elevations.Clear()
        Manufacturers.Clear()
        Counties.Clear()

        Dim emittedCsv As String = FlushQueuedCsv(Builder, planname)
        If Not String.IsNullOrEmpty(emittedCsv) Then
            acEd.WriteMessage(vbLf & "CSV written: " & emittedCsv)
        End If

        If Not String.IsNullOrWhiteSpace(acDb.Filename) Then
            DeleteStrayDsdFiles(SelectedFolder & "\" & TodaysDate & "\")
        End If

        Me.Close()

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

    Private Sub ZoomObjectsInViewport(Layout As List(Of String), plannumber As String, Builder As String, FRSheets As Boolean)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim PlanString As String
        Dim Framing As Boolean = False

        PlanString = "PLAN " & plannumber

        Dim Swing As String

        Dim layoutId As ObjectId

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            ' Reference the Layout Manager
            Dim acLayoutMgr As LayoutManager = LayoutManager.Current
            Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

            layoutId = acLayoutMgr.GetLayoutId(Layout(4))

            Dim acLayout As Autodesk.AutoCAD.DatabaseServices.Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Autodesk.AutoCAD.DatabaseServices.Layout)

            Try
                Dim Layid As ObjectId

                Layid = layouts.GetAt(Layout(4))

                Dim lay As Autodesk.AutoCAD.DatabaseServices.Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Layout)
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
                        Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.DynamicBlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
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

                                ElseIf tagvalue.Contains("ELEVATION") Then

                                    attref.TextString = "HVAC DESIGN REPORT"

                                ElseIf tagvalue.Contains("CUSTOMER'S NAME") Then

                                    attref.TextString = UCase(Builder)

                                ElseIf tagvalue.Contains("PLANDATE") Then

                                    attref.TextString = Date.Today.ToString("d")

                                ElseIf tagvalue.Contains("FR-1") Then

                                    attref.TextString = "AV-" & Layout(4)

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

            Catch es As Autodesk.AutoCAD.Runtime.Exception
                MsgBox(es.Message)
            End Try

            ' Save the changes made
            acTrans.Commit()

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
                Dim vp = TryCast(tr.GetObject(vpId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Viewport)
                If vp Is Nothing Then Continue For
                'If vp.Number = 1 Then Continue For ' never touch overall PS viewport


                vp.UpgradeOpen()
                Dim wasLocked = vp.Locked
                vp.Locked = False


                ' Do not change vp.Width / vp.Height (paper units)

                ' Camera straight down, no twist
                vp.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
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
                        pc = PlotConfigManager.SetCurrentConfig("ARCXIS - DWG To PDF.pc3")
                    Catch
                        ' ignore; Publisher can still use per-layout NPS
                    End Try

                    ' 7) Publish silently
                    ' 7) Publish silently
                    Application.Publisher.PublishExecute(dsd, pc)

                    ' wait briefly for the PDF to appear then register it
                    If WaitForFileExists(pdfFile, 30000) Then
                        RegisterPublishedPdf(pdfFile)
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

    ' ================== NEW QUEUING MODEL (one logical set per PDF) ==================
    ' We now capture exactly one metadata set per PDF (PLAN, ELEV, SW, MOD) and store
    ' rows as triples: Identifier | Key | Value.
    ' Identifier format: "Builder - PdfName"
    ' Keys: PLAN, ELEV, SW, MOD
    ' Value: extracted or empty.
    '
    ' This replaces the previous wide layout-row accumulation. Multiple calls for the
    ' same PDF will be ignored (first wins) to prevent duplicates.
    Public Shared Sub QueueLayoutsForCsv(layoutList As List(Of List(Of String)), pdfName As String, builder As String, planName As String, projectnumber As String)
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
        Dim county As String = first(6).Trim()
        Dim FuelType As String = first(9).Trim()
        Dim Manufacturer As String = first(10).Trim()

        If first.Count > 2 Then modName = If(first(2), "").Trim()
        If Not String.IsNullOrWhiteSpace(first(8)) Then
            elevation = first(8).Trim()
        ElseIf first.Count > 7 Then
            elevation = If(first(7), "").Trim()
        End If


        ' Store four logical rows: PLAN / ELEV / SW / MOD
        _pendingRows.Add(New List(Of String) From {identifier, "PLAN", planName})
        '_pendingRows.Add(New List(Of String) From {identifier, "ELEV", elevation})
        _pendingRows.Add(New List(Of String) From {identifier, "PLAN TYPE", "HVAC"})
        _pendingRows.Add(New List(Of String) From {identifier, "MANUFACTURER", Manufacturer})
        _pendingRows.Add(New List(Of String) From {identifier, "FUEL TYPE", fueltype})
        _pendingRows.Add(New List(Of String) From {identifier, "COUNTY", county})
    End Sub

    ' ================== FLUSH UPDATED FORMAT ==================
    ' Emits a single CSV aggregating all queued PDFs:
    ' Header: Identifier,Key,Value
    ' Each queued PDF contributes exactly four rows (PLAN/ELEV/SW/MOD).
    ' File name derived from first builder/plan found (sanitized); falls back to Combined.csv.
    Public Shared Function FlushQueuedCsv(Optional builder As String = Nothing, Optional planName As String = Nothing) As String
        If _pendingRows Is Nothing OrElse _pendingRows.Count = 0 Then Return Nothing

        Dim pendingDir As String = "\\egnytedrive\energyinspectors\shared\fs2\k\DPIS Drawings\PDF File Data\Pending"
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

        Dim fileBase As String = $"{SafeSegment(If(builder, ""))}_{SafeSegment(If(planName, ""))}"
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

    Public Shared Function CombineRegisteredPdfs(outputPath As String, Optional watermark As String = Nothing, Optional fontName As String = "Arial", Optional fontSize As Double = 144, Optional opacity As Double = 0.15, Optional angle As Double = 232.35) As String
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
                                        angle = 240
                                        fontSize = 240
                                        ' Prepare font and brush with alpha
                                        Dim font As New XFont(fontName, fontSize, XFontStyle.Bold)
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
                    Catch
                        ' skip single-source failures and continue
                    End Try
                Next

                Dim dir = Path.GetDirectoryName(outputPath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

                outDoc.Save(outputPath)

                ' Clear registered list after successful combine
                _publishedPdfs.Clear()

                Return outputPath
            Catch
                Return Nothing
            End Try
        End SyncLock
    End Function

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


End Class