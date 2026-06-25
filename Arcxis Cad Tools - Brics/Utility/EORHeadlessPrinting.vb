Imports System.Collections
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Bricscad.PlottingServices
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document
Imports Exception = Teigha.Runtime.Exception
Imports Layout = Teigha.DatabaseServices.Layout
Imports Path = System.IO.Path

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.EORHeadlessPrinting))>
Namespace Arcxis_Cad_Tools

    Public Class EORHeadlessPrinting

        Public Shared Sub PrintEOR()

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
            Dim CustomPrinting As Boolean = False
            Dim PropertyPlanType As String = ""

            Builder = FileManipulation.GetCustomDwgPropForDoc(acDb, "BUILDER")
            planname = FileManipulation.GetCustomDwgPropForDoc(acDb, "PLAN")
            PropertyPlanType = FileManipulation.GetCustomDwgPropForDoc(acDb, "PLAN TYPE")
            ProjectNumber = If(FileManipulation.GetCustomDwgPropReliable("PROJECT NUMBER"), "")

            Using acTrans3 As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim lytab As LayerTable = acTrans3.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim CurrentLayer As LayerTableRecord = acTrans3.GetObject(CurrentLayerID, OpenMode.ForRead)

                acCurDb.Clayer = lytab("0")

                For Each layer In lytab

                    Dim lytr As LayerTableRecord = acTrans3.GetObject(layer, OpenMode.ForWrite)

                    If lytr.Name Like "S-ANNO-AUTOMATION" Then
                        lytr.IsPlottable = False
                    End If

                    If lytr.Name Like "*SEAL-TML-FL-ES" Or lytr.Name Like "*S-ANNO-90%" Or lytr.Name Like "*S-ANNO-N4CONST" Or lytr.Name Like "*S-ANNO-REV ONLY" Then
                        lytab.UpgradeOpen()
                        lytr.IsFrozen = True
                        lytr.IsOff = True
                    End If

                Next
                acTrans3.Commit()
                acDoc.Editor.Regen()
            End Using

            Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                Dim blkTable As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                Dim blkTableRec As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                Dim order As String() = {Nothing, Nothing, "PLANTYPE", "SW", "PRNT", "SIZE", "OPTIONS", "ELEV", "", "FLR", "MAT"}

                For Each objId As ObjectId In blkTableRec

                    Dim entity As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                    If TypeOf entity Is BlockReference Then

                        Dim blkRef As BlockReference = CType(entity, BlockReference)
                        Dim blkDef As BlockTableRecord = TryCast(acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                        If blkDef.Name = "EOR PAGE" Then

                            Dim blockobjecthandle As String = blkRef.Handle.Value.ToString()
                            Dim valuesList As New List(Of String)

                            insertionPoint2 = blkRef.Position
                            valuesList.Add(blockobjecthandle)
                            valuesList.Add(insertionPoint2.ToString())

                            Dim tagValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                            If blkRef.AttributeCollection.Count > 0 Then
                                For Each attId As ObjectId In blkRef.AttributeCollection
                                    Dim attRef As AttributeReference = TryCast(acTrans.GetObject(attId, OpenMode.ForRead), AttributeReference)
                                    If attRef Is Nothing Then Continue For

                                    Dim tag As String = If(attRef.Tag, "").Trim()
                                    Dim txt As String = If(attRef.TextString, "").Trim()
                                    If tag = "" OrElse txt = "" Then Continue For

                                    If Not tagValues.ContainsKey(tag) Then tagValues(tag) = txt

                                    Select Case tag.ToUpperInvariant()
                                        Case "PLANTYPE"
                                            If Not PlanType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                PlanType.Add(attRef.TextString)
                                            End If

                                        Case "ELEV"
                                            If attRef.TextString.Contains(",") Then
                                                For Each value As String In attRef.TextString.Split(","c)
                                                    If Not Elevations.Contains(value) Then
                                                        Elevations.Add(value)
                                                    End If
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

                                        Case "OPTIONS"
                                            If Not Options.Contains(attRef.TextString) Then
                                                Options.Add(attRef.TextString)
                                            End If
                                    End Select
                                Next
                            End If

                            While valuesList.Count < order.Length
                                valuesList.Add(String.Empty)
                            End While

                            For i As Integer = 0 To order.Length - 1
                                Dim key = order(i)
                                If key IsNot Nothing AndAlso tagValues.ContainsKey(key) Then
                                    valuesList(i) = tagValues(key)
                                End If
                            Next

                            If valuesList.Count >= 7 Then
                                AllValues.Add(valuesList)
                            End If
                        End If
                    End If
                Next

                acTrans.Commit()
            End Using

            Elevations.Sort()
            Dim insertionPoint1 As Point3d
            Dim LeftLayoutList As New List(Of List(Of String))
            Dim RightLayoutList As New List(Of List(Of String))
            Dim FinalLeftList As New List(Of List(Of String))
            Dim FinalRightList As New List(Of List(Of String))
            Dim counter As Integer = 1

            Dim NewFolderLocation As String
            Dim SelectedFolder As String = GetCurrentDwgFolder()
            If String.IsNullOrEmpty(SelectedFolder) Then Return

            Dim TodaysDate As String = Date.Today.ToString("MM dd yy", CultureInfo.InvariantCulture)

            Application.SetSystemVariable("imageframe", 1)
            Application.SetSystemVariable("imageframe", 0)
            Application.SetSystemVariable("PDFSHX", 0)

            Dim SealLoop As New List(Of String)
            Dim SealToStamp As String = ""
            SealLoop.Add("SEAL-TML-FL-ES")

            If CustomPrinting = False Then
                For Each seal In SealLoop
                    SealToStamp = "24x36_TITLE BLK|" & seal
                    TurnOnOrOffLayer(SealToStamp, True)

                    For Each ElevValue In Elevations
                        For Each valueList In AllValues

                            If valueList(7).Contains(ElevValue) Then

                                If valueList(7).Contains(",") Then
                                    Dim result As New List(Of String)
                                    For Each part As String In valueList(7).Split(","c)
                                        result.Add(part.Trim())
                                    Next

                                    If Not result.Contains(ElevValue.Trim(), StringComparer.OrdinalIgnoreCase) Then
                                        Continue For
                                    End If

                                ElseIf valueList(7).Length > ElevValue.Length Then
                                    Continue For
                                End If

                                Dim insertionPointStr As String = valueList(1)
                                insertionPoint1 = CreatePoint3dFromString(insertionPointStr)

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
                        Dim framecounter As Integer = 0
                        Dim bracingcounter As Integer = 0
                        RightLayoutList = RightLayoutList.OrderBy(Function(entry) Integer.Parse(entry(4))).ToList()

                        For Each RightEntry In RightLayoutList
                            If RightEntry(4) = counter Then
                                RightEntry(8) = ElevValue
                                FinalRightList.Add(RightEntry)
                                counter += 1
                                ZoomObjectsInViewport(RightEntry, planname, Builder, framecounter, bracingcounter)
                                If RightEntry(2) = "FRAMING" Then
                                    framecounter += 1
                                End If
                                If RightEntry(2) = "BRACING" Then
                                    bracingcounter += 1
                                End If
                            End If
                        Next

                        NewFolderLocation = SelectedFolder & "\" & "EOR"

                        If FinalRightList.Count <> 0 Then
                            If Not Directory.Exists(NewFolderLocation) Then
                                Directory.CreateDirectory(NewFolderLocation)
                            End If

                            pdfname = UCase("RIGHT " & PropertyPlanType & " " & planname & " " & ElevValue)
                        End If

                        If FinalRightList.Count <> 0 Then
                            PlotTAutomatedTabs(FinalRightList, pdfname, NewFolderLocation)
                            FileManipulation.QueueLayoutsForCsv(FinalRightList, pdfname, Builder, planname, ProjectNumber)
                        End If

                        counter = 1
                        framecounter = 0
                        bracingcounter = 0
                        LeftLayoutList = LeftLayoutList.OrderBy(Function(entry) Integer.Parse(entry(4))).ToList()

                        For Each LeftEntry In LeftLayoutList
                            If LeftEntry(4) = counter Then
                                LeftEntry(8) = ElevValue
                                FinalLeftList.Add(LeftEntry)
                                counter += 1
                                ZoomObjectsInViewport(LeftEntry, planname, Builder, framecounter, bracingcounter)
                                If LeftEntry(2) = "FRAMING" Then
                                    framecounter += 1
                                End If
                                If LeftEntry(2) = "BRACING" Then
                                    bracingcounter += 1
                                End If
                            End If
                        Next

                        If FinalLeftList.Count <> 0 Then
                            If Not Directory.Exists(NewFolderLocation) Then
                                Directory.CreateDirectory(NewFolderLocation)
                            End If

                            pdfname = UCase("LEFT " & PropertyPlanType & " " & planname & " " & ElevValue)
                        End If

                        If FinalLeftList.Count <> 0 Then
                            PlotTAutomatedTabs(FinalLeftList, pdfname, NewFolderLocation)
                            FileManipulation.QueueLayoutsForCsv(FinalLeftList, pdfname, Builder, planname, ProjectNumber)
                        End If

                        FinalLeftList.Clear()
                        FinalRightList.Clear()
                        LeftLayoutList.Clear()
                        RightLayoutList.Clear()
                    Next

                    TurnOnOrOffLayer(SealToStamp, False)
                Next
            End If

            TurnOnOrOffLayer(SealToStamp, False)

            Dim lm As LayoutManager = LayoutManager.Current
            lm.CurrentLayout = "Model"
            PlanType.Clear()
            Elevations.Clear()
            Options.Clear()

            Dim cadFileName As String = Path.GetFileNameWithoutExtension(acCurDb.Filename)
            Dim emittedCsv As String = FileManipulation.FlushQueuedCsv(Builder, planname, cadFileName, PropertyPlanType)
            If Not String.IsNullOrEmpty(emittedCsv) Then
                acEd.WriteMessage(vbLf & "CSV written: " & emittedCsv)
            End If

            If Not String.IsNullOrWhiteSpace(acDb.Filename) Then
                DeleteStrayDsdFiles(SelectedFolder & "\" & TodaysDate & "\")
            End If

        End Sub

        Shared Function UpdateIrcIeccOnGeneralNotes(newIrc As String, newIecc As String) As Integer
            Dim doc = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return 0
            Dim db = doc.Database
            Dim updated As Integer = 0

            Using doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    For Each id As ObjectId In ms
                        Dim br As BlockReference = TryCast(tr.GetObject(id, OpenMode.ForRead, True), BlockReference)
                        If br Is Nothing OrElse br.IsErased Then Continue For

                        Dim def As BlockTableRecord = TryCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                        If def Is Nothing Then Continue For
                        If Not def.Name.Equals("AV-General Notes", StringComparison.OrdinalIgnoreCase) Then Continue For

                        Dim touched As Boolean = False
                        For Each attId As ObjectId In br.AttributeCollection
                            Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                            If attRef Is Nothing Then Continue For

                            Dim tag = If(attRef.Tag, String.Empty).Trim()
                            If tag.Equals("IRC", StringComparison.OrdinalIgnoreCase) Then
                                attRef.TextString = newIrc
                                touched = True
                            ElseIf tag.Equals("IECC", StringComparison.OrdinalIgnoreCase) Then
                                attRef.TextString = newIecc
                                touched = True
                            End If
                        Next

                        If touched Then updated += 1
                    Next

                    tr.Commit()
                End Using
            End Using

            Return updated
        End Function

        Shared Function GetCurrentDwgFolder() As String
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            If Not String.IsNullOrWhiteSpace(db.Filename) Then
                Return Path.GetDirectoryName(db.Filename)
            End If
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

        Shared Function CreatePoint3dFromString(pointStr As String) As Point3d
            pointStr = pointStr.Trim("(", ")")

            Dim coords() As String = pointStr.Split(New Char() {","c, " "c}, StringSplitOptions.RemoveEmptyEntries)

            If coords.Length <> 3 Then
                Throw New ArgumentException("The point String must contain exactly three coordinates.")
            End If

            Dim x As Double = Double.Parse(coords(0).Trim())
            Dim y As Double = Double.Parse(coords(1).Trim())
            Dim z As Double = Double.Parse(coords(2).Trim())

            Return New Point3d(x, y, z)
        End Function

        Shared Sub ZoomObjectsInViewport(Layout As List(Of String), plannumber As String, Builder As String, framingcounter As Integer, bracingcounter As Integer)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim PlanString As String = plannumber

            Dim PlanUpper As String = PlanString
            Dim Plan As String = ""
            Dim firstDash As Integer = PlanString.IndexOf("-"c)
            If firstDash >= 0 Then
                Dim secondDash As Integer = PlanString.IndexOf("-"c, firstDash + 1)
                If secondDash >= 0 Then
                    PlanUpper = PlanString.Substring(0, secondDash).Trim()
                    Plan = PlanString.Substring(secondDash + 1).Trim()
                End If
            End If

            Dim layoutId As ObjectId

            Dim SheetAbbrev As String = ""
            Dim TB1 As String = ""
            Dim TB2 As String = ""
            Dim TB3 As String = ""

            If UCase(Layout(2)) = "FOUNDATION" Then
                SheetAbbrev = "S1." & Layout(4)
                TB1 = "SLAB ON GRADE"
                TB2 = "FOUNDATION PLAN"
            ElseIf UCase(Layout(2)) = "FRAMING" Then
                SheetAbbrev = "S2." & framingcounter.ToString()
                If UCase(Layout(9)) = "MAIN FLOOR" Then
                    TB1 = "MAIN FLOOR"
                    TB2 = "WALL FRAMING PLAN"
                ElseIf UCase(Layout(9)) = "UPPER FLOOR" Then
                    TB1 = "UPPER FLOOR"
                    TB2 = "WALL FRAMING PLAN"
                ElseIf UCase(Layout(9)) = "ROOF" Then
                    TB1 = "ROOF/CEILING"
                    TB2 = "FRAMING PLAN"
                End If
            ElseIf UCase(Layout(2)) = "BRACING" Then
                SheetAbbrev = "S3." & bracingcounter.ToString()
                If UCase(Layout(9)) = "MAIN FLOOR" Then
                    TB1 = "MAIN FLOOR"
                    TB2 = "WALL BRACING PLAN"
                    TB3 = UCase(Layout(6))
                ElseIf UCase(Layout(9)) = "UPPER FLOOR" Then
                    TB1 = "UPPER FLOOR"
                    TB2 = "WALL BRACING PLAN"
                    TB3 = UCase(Layout(6))
                End If
            End If

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim acLayoutMgr As LayoutManager = LayoutManager.Current
                Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

                layoutId = acLayoutMgr.GetLayoutId(Layout(4))
                Dim acLayout As Layout = DirectCast(acTrans.GetObject(layoutId, OpenMode.ForWrite), Layout)

                Try
                    Dim Layid As ObjectId = layouts.GetAt(Layout(4))
                    Dim lay As Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForRead), Layout)
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim blkBlkRec As BlockTableRecord = acTrans.GetObject(lay.BlockTableRecordId, OpenMode.ForRead)
                    Dim vpIds As ObjectIdCollection = New ObjectIdCollection()

                    For Each objID As ObjectId In blkBlkRec

                        If (objID.ObjectClass.DxfName.ToUpper = "VIEWPORT") Then
                            vpIds.Add(objID)

                        ElseIf (objID.ObjectClass.DxfName.ToUpper = "INSERT") Then

                            Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(objID, OpenMode.ForRead), BlockReference)
                            Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                            Dim RevvblockName As String = RevTblRec.Name

                            If RevvblockName.Contains("Page Information") Then

                                For Each attId As ObjectId In RevBlockRef.AttributeCollection

                                    Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                    Dim tagvalue As String = attref.Tag

                                    If tagvalue.Contains("PLAN-Upper") Then
                                        attref.TextString = PlanUpper
                                    ElseIf tagvalue.Contains("Plan") AndAlso Not tagvalue.Contains("Upper") Then
                                        attref.TextString = Plan
                                    ElseIf tagvalue.Contains("ELEVATION") Then
                                        attref.TextString = ""
                                    ElseIf tagvalue.Contains("SWING") Then
                                        attref.TextString = ""
                                    ElseIf tagvalue.Contains("SHEET") Then
                                        attref.TextString = SheetAbbrev
                                    ElseIf tagvalue.Contains("TB1") Then
                                        attref.TextString = TB1
                                    ElseIf tagvalue.Contains("ADRESS") Or tagvalue.Contains("ADDRESS") Then
                                        attref.TextString = ""
                                    ElseIf tagvalue.Contains("COMMUNITY") Then
                                        attref.TextString = ""
                                    ElseIf tagvalue.Contains("TB2") Then
                                        attref.TextString = TB2
                                    ElseIf tagvalue.Contains("TB3") Then
                                        attref.TextString = TB3
                                    End If

                                Next

                            End If

                        End If

                    Next

                    ModifyViewPortCenter(vpIds, Layout(4), Layout(1), Layout(5))

                Catch es As Exception
                    MsgBox(es.Message)
                End Try

                acTrans.Commit()

            End Using

        End Sub

        Shared Sub ModifyViewPortCenter(VPIDS As ObjectIdCollection, LAYOUT As String, BASEPOINT As String, Scale As String)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = acDoc.Database
            Dim basePt As Point3d = CreatePoint3dFromString(BASEPOINT)

            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim lm As LayoutManager = LayoutManager.Current
                Dim layoutId As ObjectId = lm.GetLayoutId(LAYOUT)
                Dim layout1 As Layout = tr.GetObject(layoutId, OpenMode.ForRead)

                Dim paperCenter As New Point3d(16.1875, 12, 0)
                Dim dx As Double = 0.0, dy As Double = 0.0

                Dim customScale As Double
                Select Case Scale
                    Case "1/8" : customScale = 96.0 : dx = 709.92 : dy = -504.0
                    Case "3/32" : customScale = 128.0 : dx = 946.558 : dy = -672.0
                    Case "24X36" : customScale = 48 : dx = 752.5625 : dy = -558.125
                    Case Else : customScale = 1.0
                End Select

                For Each vpId As ObjectId In VPIDS
                    Dim vp = TryCast(tr.GetObject(vpId, OpenMode.ForRead), Viewport)
                    If vp Is Nothing Then Continue For
                    If vp.Number = 1 Then Continue For

                    vp.UpgradeOpen()
                    Dim wasLocked = vp.Locked
                    vp.Locked = False

                    vp.ViewDirection = Teigha.Geometry.Vector3d.ZAxis
                    vp.TwistAngle = 0.0

                    If vp.Height > 23.26 Then Continue For
                    If vp.CustomScale <> 1 / customScale Then
                        vp.CustomScale = 1 / customScale
                    End If

                    vp.ViewHeight = vp.Height * customScale

                    Dim cx As Double = basePt.X + dx
                    Dim cy As Double = basePt.Y + dy
                    vp.ViewCenter = New Point2d(cx, cy)
                    vp.CenterPoint = paperCenter

                    vp.Locked = wasLocked
                    vp.DowngradeOpen()
                Next

                tr.Commit()
            End Using
        End Sub

        Shared Sub PlotTAutomatedTabs(lAYOUTLIST As List(Of List(Of String)), Pdfname As String, NEWFOLDERLOCATION As String, Optional DPISCTB As Boolean = False)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                    Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                    Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                    Dim pdfFile As String = String.Empty
                    Application.SetSystemVariable("BackGroundPlot", 0)
                    Application.SetSystemVariable("CMDDIA", 0)
                    Application.SetSystemVariable("FILEDIA", 0)

                    Try
                        Dim outputDir As String = NEWFOLDERLOCATION
                        If Not Directory.Exists(outputDir) Then Directory.CreateDirectory(outputDir)

                        Dim dwgprefix As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
                        Dim DWGnm As String = CStr(Application.GetSystemVariable("DWGNAME"))
                        Dim dwgFile As String = Path.Combine(dwgprefix, DWGnm)

                        pdfFile = Path.Combine(outputDir, Pdfname & ".pdf")
                        Dim dsdFile As String = Path.Combine(outputDir, Pdfname & ".dsd")

                        If File.Exists(pdfFile) Then
                            acTrans.Commit()
                            Return
                        End If

                        Dim dsd As New DsdData()
                        Dim dsdEntries As New DsdEntryCollection()

                        For Each lay As List(Of String) In lAYOUTLIST
                            Dim title As String = Path.GetFileNameWithoutExtension(DWGnm) & "-" &
                                      lay(2) & "-" & lay(3) & "-" & lay(7) & "-" & lay(4)

                            Dim de As New DsdEntry()
                            de.DwgName = dwgFile
                            de.Layout = lay(4)
                            de.Title = title
                            de.Nps = "EOR-Florida"
                            de.NpsSourceDwg = dwgFile
                            dsdEntries.Add(de)
                        Next

                        dsd.SetDsdEntryCollection(dsdEntries)
                        dsd.SheetType = SheetType.MultiPdf
                        dsd.NoOfCopies = 1
                        dsd.IsHomogeneous = True
                        dsd.ProjectPath = outputDir
                        dsd.DestinationName = pdfFile
                        dsd.Dwf3dOptions.PublishWithMaterials = True
                        dsd.Dwf3dOptions.GroupByXrefHierarchy = True

                        dsd.SetUnrecognizedData("PromptForDwfName", "FALSE")
                        dsd.SetUnrecognizedData("PromptForName", "FALSE")

                        If File.Exists(dsdFile) Then File.Delete(dsdFile)
                        dsd.WriteDsd(dsdFile)

                        Dim text As String = File.ReadAllText(dsdFile)

                        Dim ensure As New List(Of String) From {
                            "PromptForDwfName=False",
                            "PromptForName=False",
                            "PwdProtectPublishedDWF=False",
                            "IncludeHyperlinks=TRUE",
                            "IncludeLayer=FALSE",
                            "Type=6",
                            "OutDir=" & outputDir.Replace("\", "\\"),
                            "Dst=" & pdfFile.Replace("\", "\\")
                        }

                        text = text.Replace("PromptForDwfName=True", "PromptForDwfName=False")
                        text = text.Replace("PromptForName=True", "PromptForName=False")
                        text = text.Replace("Type=3", "Type=6")

                        If Not text.Contains(vbCrLf & "OutDir=") Then text &= vbCrLf & "OutDir=" & outputDir
                        If Not text.Contains(vbCrLf & "Dst=") Then text &= vbCrLf & "Dst=" & pdfFile

                        For Each line In ensure
                            Dim key = line.Split("="c)(0)
                            Dim idx = text.IndexOf(key & "=", StringComparison.OrdinalIgnoreCase)
                            If idx >= 0 Then
                                Dim rowEnd = text.IndexOfAny({ControlChars.Cr, ControlChars.Lf}, idx)
                                If rowEnd < 0 Then rowEnd = text.Length
                                text = text.Remove(idx, rowEnd - idx).Insert(idx, line)
                            Else
                                text &= vbCrLf & line
                            End If
                        Next

                        File.WriteAllText(dsdFile, text)
                        dsd.ReadDsd(dsdFile)

                        Dim pc As PlotConfig = FileManipulation.ResolvePlotConfig()

                        Dim doc = Application.DocumentManager.MdiActiveDocument
                        Dim db = doc.Database

                        Using tr = db.TransactionManager.StartTransaction()
                            tr.Commit()
                        End Using

                        doc.Editor.Regen()

                        Dim auditJobId As String = HeadlessPublishAudit.StartJob("EORHeadlessPrinting.PublishPdf", pdfFile, dsdEntries.Count)
                        Application.Publisher.PublishExecute(dsd, pc)
                        HeadlessPublishAudit.MarkInfo(auditJobId, "EORHeadlessPrinting.PublishPdf", pdfFile, "PublishExecute returned to caller.")

                        If HeadlessPublishAudit.WaitForOutput(pdfFile, 60000) Then
                            Dim fi As New FileInfo(pdfFile)
                            HeadlessPublishAudit.MarkSuccess(auditJobId, "EORHeadlessPrinting.PublishPdf", pdfFile, "Output ready. Size=" & fi.Length.ToString() & " bytes.")
                        Else
                            HeadlessPublishAudit.MarkFailed(auditJobId, "EORHeadlessPrinting.PublishPdf", pdfFile, "Publish returned but output file did not appear within timeout.")
                        End If

                        If File.Exists(dsdFile) Then File.Delete(dsdFile)

                        acTrans.Commit()

                    Catch ex As System.Exception
                        HeadlessPublishAudit.MarkFailed("", "EORHeadlessPrinting.PublishPdf", pdfFile, ex.Message)
                        Throw
                    Finally
                        Application.SetSystemVariable("BackGroundPlot", bgPrev)
                        Application.SetSystemVariable("CMDDIA", cmdPrev)
                        Application.SetSystemVariable("FILEDIA", fileDiaPrev)
                    End Try

                End Using
            End Using
        End Sub

        Shared Sub TurnOnOrOffLayer(layerName As String, TurnOn As Boolean)
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

        Private Shared Sub DeleteStrayDsdFiles(rootFolder As String)
            If String.IsNullOrWhiteSpace(rootFolder) OrElse Not Directory.Exists(rootFolder) Then Exit Sub

            Try
                For Each dsdFile In Directory.GetFiles(rootFolder, "*.dsd", SearchOption.TopDirectoryOnly)
                    Try
                        File.Delete(dsdFile)
                    Catch ex As Exception
                    End Try
                Next

                For Each subDir In Directory.GetDirectories(rootFolder)
                    DeleteStrayDsdFiles(subDir)
                Next
            Catch ex As Exception
            End Try
        End Sub

    End Class
End Namespace
