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
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports Arcxis_Cad_Tools_Brics.Arcxis_Cad_Tools_Brics
Imports Excel = Microsoft.Office.Interop.Excel
Imports Path = System.IO.Path
Imports Tuple = System.Tuple
Imports Bricscad.ApplicationServices
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document
Imports Exception = Teigha.Runtime.Exception
Imports Layout = Teigha.DatabaseServices.Layout
Imports Color = Teigha.Colors.Color
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.AtticVentHeadlessPrinting))>
Namespace Arcxis_Cad_Tools

    Public Class AtticVentHeadlessPrinting
        Private Shared _pendingRows As New List(Of List(Of String))()
        Public Shared Sub PrintAtticVent()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            Dim acEd As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database

            Dim PlanType As New List(Of String)
            Dim Swings As New List(Of String)
            Dim Elevations As New List(Of String)
            Dim Options As New List(Of String)
            Dim AtticType As New List(Of String)
            Dim AllValues As New List(Of List(Of String))
            Dim insertionPoint2 As Point3d
            Dim pdfname As String = ""
            Dim planname As String = ""
            Dim CurrentLayerID As ObjectId = acCurDb.Clayer
            Dim Builder As String = ""
            Dim ProjectNumber As String = ""
            Dim CustomPrinting As Boolean = False


            Builder = FileManipulation.GetCustomDwgPropForDoc(acDb, "BUILDER")
            planname = FileManipulation.GetCustomDwgPropForDoc(acDb, "PLAN")

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

            ' --- Early validation pass (separate transaction) ---
            Dim hasInvalidAttributes As Boolean = False
            Dim invalidTags As New List(Of String)()

            Using txValidate As Transaction = acDb.TransactionManager.StartTransaction()
                Dim blkTable As BlockTable = txValidate.GetObject(acDb.BlockTableId, OpenMode.ForRead)
                Dim blkTableRec As BlockTableRecord = txValidate.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                For Each objId As ObjectId In blkTableRec
                    Dim entity As Entity = TryCast(txValidate.GetObject(objId, OpenMode.ForRead), Entity)
                    If TypeOf entity IsNot BlockReference Then Continue For

                    Dim blkRef As BlockReference = CType(entity, BlockReference)
                    Dim blkDef As BlockTableRecord = TryCast(txValidate.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    If blkDef Is Nothing OrElse Not blkDef.Name.Equals("AtticVentPage", StringComparison.OrdinalIgnoreCase) Then Continue For
                    If blkRef.AttributeCollection.Count = 0 Then Continue For

                    For Each attId As ObjectId In blkRef.AttributeCollection
                        Dim attRef As AttributeReference = TryCast(txValidate.GetObject(attId, OpenMode.ForRead), AttributeReference)
                        If attRef Is Nothing Then Continue For
                        Dim tag As String = If(attRef.Tag, String.Empty).Trim().ToUpperInvariant()
                        If tag.Contains("LXS") OrElse tag.Contains("MOD") OrElse tag.Contains("TND") Then
                            hasInvalidAttributes = True
                            If Not invalidTags.Contains(tag) Then invalidTags.Add(tag)
                        End If
                    Next

                    If hasInvalidAttributes Then Exit For
                Next

                txValidate.Abort()
            End Using

            If hasInvalidAttributes Then
                Dim errorReport As New System.Text.StringBuilder()
                errorReport.AppendLine("=== ATTIC VENT PAGE BLOCK ERROR ===")
                errorReport.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                errorReport.AppendLine($"DWG File: {Path.GetFileName(acDb.Filename)}")
                errorReport.AppendLine()
                errorReport.AppendLine("ERROR: The ""AtticVentPage"" block has incorrect attributes.")
                errorReport.AppendLine()
                errorReport.AppendLine("Invalid attributes found:")
                For Each tag In invalidTags
                    errorReport.AppendLine($"  - {tag}")
                Next
                errorReport.AppendLine()
                errorReport.AppendLine("RESOLUTION REQUIRED:")
                errorReport.AppendLine("The ""AtticVentPage"" has the incorrect attributes. An automation to fix this was attempted. Please review pdf's for accuracy")

                FileManipulation.CreateTextFileInDwgLocation("AtticVentPage_BlockError.log", errorReport.ToString())
                MechanicalFunctions.UpdateAttributeDefinitions() ' this saves the DWG
            End If

            ' Start a transaction
            Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
                ' Open the Block table for read
                Dim blkTable As BlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead)

                ' Open the BlockTableRecord (ModelSpace) for read
                Dim blkTableRec As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                ' --- define desired positions ONCE before the attribute loop ---
                Dim order As String() = {Nothing, Nothing, "PLANTYPE", "SW", "PRNT", "SIZE", "OPTIONS", "ELEV", "", "VENTTYPE", "ATTICTYPE"}

                ' Iterate through the ModelSpace block table record
                For Each objId As ObjectId In blkTableRec

                    Dim entity As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

                    ' Check if the entity is a block reference
                    If TypeOf entity Is BlockReference Then

                        Dim blkRef As BlockReference = CType(entity, BlockReference)

                        ' Get the block table record for the block reference
                        Dim blkDef As BlockTableRecord = TryCast(acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                        ' Check if the block name is "PAGE"
                        If blkDef.Name = "AtticVentPage" Then

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
                                        Case "VENTTYPE"

                                            If Not PlanType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                PlanType.Add(attRef.TextString)
                                            End If

                                        Case "ATTICTYPE"

                                            If Not AtticType.Contains(attRef.TextString) And attRef.TextString <> "" Then
                                                AtticType.Add(attRef.TextString)
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

                                        Case "OPTIONS"

                                            If Not Options.Contains(attRef.TextString) Then

                                                Options.Add(attRef.TextString)

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

            Dim IECCList As New List(Of String) From {"2015", "2018", "2021", "2024"}
            Dim IRCList As New List(Of String) From {"2015", "2018", "2021", "2024"}

            Dim SelectedFolder As String
            SelectedFolder = GetCurrentDwgFolder()

            If String.IsNullOrEmpty(SelectedFolder) Then Return

            Dim TodaysDate As String = Date.Today.ToString("MM dd yy", CultureInfo.InvariantCulture)
            Dim SealToStamp As String = "Master Seal File|S-SEAL-TML-TX"

            Application.SetSystemVariable("imageframe", 1)
            Application.SetSystemVariable("imageframe", 0)

            TurnOnOrOffLayer(SealToStamp, False)

            Dim db As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor
            Dim existingLayoutCount As Integer = 0
            Try
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    ' Count existing paperspace layouts (excluding Model)
                    Dim layoutDict As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                    If layoutDict Is Nothing Then
                        tr.Commit()
                        Return
                    End If


                    For Each entry As DBDictionaryEntry In layoutDict
                        Dim layout As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                        If layout IsNot Nothing AndAlso Not layout.ModelType Then
                            existingLayoutCount += 1
                        End If
                    Next
                End Using
            Catch ex As Exception
                ed.WriteMessage(Environment.NewLine & "Error counting layouts: " & ex.Message)
            End Try

            Dim CounterSkip As Boolean = False

            For Each type In PlanType
                For Each AttiType In AtticType

                    Dim SheetAbbrev As String = ""
                    Dim PlanAbbrev As String = ""

                    If type.ToUpper() = "FR" Or type.ToUpper() = "FIRE RATED" Then
                        PlanAbbrev = " - FR"
                    ElseIf type.ToUpper() = "EV" Or type.ToUpper() = "EDGE VENTS" Or type.ToUpper() = "EDGE" Or type.ToUpper() = "EDGE VENTING" Then
                        PlanAbbrev = " - EDGE VENTS"
                    End If

                    For Each iecc In IECCList

                        ' Early validation before the loops
                        If String.IsNullOrEmpty(SelectedFolder) Then Return

                        If Not NetworkHelpers.IsNetworkPathAccessible(SelectedFolder) Then
                            'acEd.WriteMessage(vbLf & "ERROR: Base folder is not accessible: " & SelectedFolder)
                            'acEd.WriteMessage(vbLf & "Please verify network drive is connected and try again.")
                            Return
                        End If

                        ' Replace the directory creation blocks (around lines 384, 393, etc.)
                        NewFolderLocation = SelectedFolder & "\" & planname & "\" & iecc & " IECC"

                        If Not NetworkHelpers.CreateDirectoryWithRetry(NewFolderLocation) Then
                            'acEd.WriteMessage(vbLf & "ERROR: Failed to create directory: " & NewFolderLocation)
                            Continue For ' Skip this iteration instead of crashing
                        End If

                        For Each irc In IRCList


                            If Not NetworkHelpers.IsNetworkPathAccessible(SelectedFolder) Then
                                'acEd.WriteMessage(vbLf & "ERROR: Base folder is not accessible: " & SelectedFolder)
                                'acEd.WriteMessage(vbLf & "Please verify network drive is connected and try again.")
                                Return
                            End If

                            ' Replace the directory creation blocks (around lines 384, 393, etc.)
                            NewFolderLocation = SelectedFolder & "\" & planname & "\" & iecc & " IECC\" & irc & " IRC"

                            If Not NetworkHelpers.CreateDirectoryWithRetry(NewFolderLocation) Then
                                'acEd.WriteMessage(vbLf & "ERROR: Failed to create directory: " & NewFolderLocation)
                                Continue For ' Skip this iteration instead of crashing
                            End If

                            UpdateIrcIeccOnGeneralNotes(irc, iecc)

                            For Each ElevValue In Elevations
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
                                    'valueList(6) = opt
                                    'CounterSkip = True

                                    'If valueList(9) = type AndAlso valueList(10) = AttiType AndAlso valueList(7).Contains(ElevValue) AndAlso valueList(6) = opt Then
                                    If valueList(9) = type AndAlso valueList(10) = AttiType AndAlso valueList(7).Contains(ElevValue) Then
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

                                            If valueList(3).ToUpper() = "L" Or valueList(3).ToUpper() = "LEFT" And Swings.Contains("Left") Then

                                                LeftLayoutList.Add(valueList)

                                            ElseIf valueList(3).ToUpper() = "R" Or valueList(3).ToUpper() = "RIGHT" And Swings.Contains("Right") Then

                                                RightLayoutList.Add(valueList)

                                            End If
                                        Else

                                            If valueList(3).ToUpper() = "L" Or valueList(3).ToUpper() = "LEFT" Then
                                                LeftLayoutList.Add(valueList)
                                            Else
                                                RightLayoutList.Add(valueList)
                                            End If

                                        End If
                                    End If

                                Next

                                counter = 1

                                If RightLayoutList.Count > existingLayoutCount Then
                                    FileManipulation.EnsureLayoutCount(RightLayoutList.Count)
                                End If

                                ' Sort RightLayoutList by sequence number (index 4) before processing
                                Dim sortedRightList = RightLayoutList.OrderBy(Function(entry) CInt(entry(4))).ToList()

                                For Each RightEntry In sortedRightList
                                    RightEntry(8) = ElevValue
                                    FinalRightList.Add(RightEntry)
                                    ZoomObjectsInViewport(RightEntry, planname, Builder, False)
                                Next

                                If FinalRightList.Count <> 0 Then

                                    pdfname = ("RIGHT ATTIC VENT " & planname & " " & ElevValue & PlanAbbrev & " " & iecc & " IECC" & " " & irc & " IRC").ToUpper()

                                End If

                                ' Right side
                                If FinalRightList.Count <> 0 Then
                                    PlotTAutomatedTabs(FinalRightList, pdfname, NewFolderLocation)
                                    FileManipulation.QueueLayoutsForCsv(FinalRightList, pdfname, Builder, planname, ProjectNumber, IRC:=irc, IECC:=iecc)
                                End If


                                counter = 1

                                If LeftLayoutList.Count > existingLayoutCount Then
                                    FileManipulation.EnsureLayoutCount(LeftLayoutList.Count)
                                End If

                                ' Sort LeftLayoutList by sequence number (index 4) before processing
                                Dim sortedLeftList = LeftLayoutList.OrderBy(Function(entry) CInt(entry(4))).ToList()

                                For Each LeftEntry In sortedLeftList
                                    LeftEntry(8) = ElevValue
                                    FinalLeftList.Add(LeftEntry)
                                    ZoomObjectsInViewport(LeftEntry, planname, Builder, False)
                                Next

                                If FinalLeftList.Count <> 0 Then

                                    pdfname = ("LEFT ATTIC VENT " & planname & " " & ElevValue & PlanAbbrev & " " & iecc & " IECC" & " " & irc & " IRC").ToUpper()

                                End If

                                ' Left side
                                If FinalLeftList.Count <> 0 Then
                                    PlotTAutomatedTabs(FinalLeftList, pdfname, NewFolderLocation)
                                    FileManipulation.QueueLayoutsForCsv(FinalLeftList, pdfname, Builder, planname, ProjectNumber, IRC:=irc, IECC:=iecc)
                                End If

                                FinalLeftList.Clear()
                                FinalRightList.Clear()
                                LeftLayoutList.Clear()
                                RightLayoutList.Clear()
                            Next

                        Next
                    Next
                Next
            Next

            Dim lm As LayoutManager = LayoutManager.Current

            lm.CurrentLayout = "Model"
            PlanType.Clear()
            Elevations.Clear()
            Options.Clear()

            Dim emittedCsv As String = FileManipulation.FlushQueuedCsv(Builder, planname)
            'If Not String.IsNullOrEmpty(emittedCsv) Then
            '    'acEd.WriteMessage(vbLf & "CSV written: " & emittedCsv)
            'End If

        End Sub

        Shared Sub PlotTAutomatedTabs(lAYOUTLIST As List(Of List(Of String)), Pdfname As String, NEWFOLDERLOCATION As String, Optional DPISCTB As Boolean = False)
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
                        If Not text.Contains(Environment.NewLine & "OutDir=") Then text &= Environment.NewLine & "OutDir=" & outputDir
                        If Not text.Contains(Environment.NewLine & "Dst=") Then text &= Environment.NewLine & "Dst=" & pdfFile

                        ' Re-apply our ensure list to be certain
                        For Each line In ensure
                            Dim key = line.Split("="c)(0)
                            Dim idx = text.IndexOf(key & "=", StringComparison.OrdinalIgnoreCase)
                            If idx >= 0 Then
                                ' replace the whole row - look for CR or LF
                                Dim rowEnd = text.IndexOfAny({Convert.ToChar(13), Convert.ToChar(10)}, idx)
                                If rowEnd < 0 Then rowEnd = text.Length
                                text = text.Remove(idx, rowEnd - idx).Insert(idx, line)
                            Else
                                text &= Environment.NewLine & line
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
                        Application.Publisher.PublishExecute(dsd, pc)

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

        ''' <summary>
        ''' Updates IRC and IECC attribute values on every "AV-General Notes" block in ModelSpace.
        ''' Returns the number of blocks updated.
        ''' </summary>
        Shared Function UpdateIrcIeccOnGeneralNotes(newIrc As String, newIecc As String) As Integer
            Dim doc = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return 0
            Dim db = doc.Database
            Dim ed = doc.Editor
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

            'ed.WriteMessage(vbLf & $"AV-General Notes updated: {updated} instance(s).")
            Return updated
        End Function

        Shared Function GetCurrentDwgFolder() As String
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            If Not String.IsNullOrWhiteSpace(db.Filename) Then
                Return Path.GetDirectoryName(db.Filename)
            End If
            ' Unsaved drawing fallback
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

        Shared Function CreatePoint3dFromString(pointStr As String) As Point3d
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

        Shared Sub ZoomObjectsInViewport(Layout As List(Of String), plannumber As String, Builder As String, FRSheets As Boolean)

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim PlanString As String
            Dim Framing As Boolean = False
            Dim OptionString As String = ""

            PlanString = "PLAN " & plannumber

            Dim Swing As String

            Dim layoutId As ObjectId

            If Layout(3) = "L" Or Layout(3).ToUpper() = "LEFT" Then
                Swing = "LEFT"
            Else
                Swing = "RIGHT"
            End If

            Dim SheetAbbrev As String = ""
            Dim PlanAbbrev As String = ""

            If Layout(9).ToUpper() = "SOFFIT" Or Layout(9).ToUpper() = "SNAP VENTING" Or Layout(9).ToUpper().Contains("SNAP") Then
                SheetAbbrev = "AV-"
                PlanAbbrev = ""
            ElseIf Layout(9).ToUpper() = "FR" Or Layout(9).ToUpper() = "FIRE RATED" Then
                SheetAbbrev = "FR-"
                PlanAbbrev = " - VENTS"
            ElseIf Layout(9).ToUpper() = "EDGE VENTS" Or Layout(9).ToUpper() = "EDGE VENTING" Or Layout(9).ToUpper() = "EDGE" Or Layout(9).ToUpper().Contains("EDGE VENT") Then
                SheetAbbrev = "FR-"
                PlanAbbrev = " - EDGE VENTS"
            End If

            If Layout(6).ToUpper() <> "BASE" Then
                OptionString = " - " & Layout(6).ToUpper()
            Else
                OptionString = ""
            End If

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                ' Get the "jaytxt" text style ID
                Dim txtStyleId As ObjectId = ObjectId.Null
                Dim txtStyleTable As TextStyleTable = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)

                If txtStyleTable.Has("BORDERTXT") Then
                    txtStyleId = txtStyleTable("BORDERTXT")
                End If

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
                                    Dim tagvalue As String = attref.Tag
                                    Dim textvalue As String = attref.TextString


                                    ' Flag to track if we made changes
                                    Dim modified As Boolean = False
                                    Dim isMiddleCenter As Boolean = False

                                    ' Check if this is one of the MiddleCenter attributes
                                    If tagvalue.Contains("PLANDATE") OrElse
                               tagvalue.Contains("1/8"" = 1'-0""") OrElse
                               tagvalue.Contains("FR-1") Then
                                        isMiddleCenter = True
                                    End If

                                    If tagvalue.Contains("PLAN ") Then

                                        attref.TextString = PlanString & OptionString & PlanAbbrev

                                        If attref.TextString.Length > 46 Then
                                            attref.Height = 3 / 32
                                            attref.WidthFactor = 0.85
                                        ElseIf attref.TextString.Length <= 46 Then
                                            attref.Height = 1 / 8
                                            attref.WidthFactor = 1.0
                                        End If

                                    ElseIf tagvalue.Contains("OPTIONAL") Then
                                        attref.TextString = "ATTIC VENTILATION DESIGN"
                                    ElseIf tagvalue.Contains("ELEVATION") Then
                                        attref.TextString = "ELEVATION " & Layout(8) & " - " & Swing & " SWING"
                                    ElseIf tagvalue.Contains("CUSTOMER'S NAME") Then
                                        attref.TextString = Builder.ToUpper()
                                    ElseIf tagvalue.Contains("PLANDATE") Then
                                        attref.TextString = Date.Today.ToString("d")
                                        modified = True
                                    ElseIf tagvalue.Contains("FR-1") Then
                                        attref.TextString = SheetAbbrev & Layout(4)
                                        modified = True
                                    ElseIf tagvalue.Contains("1/8"" = 1'-0""") Then
                                        attref.TextString = "NTS"
                                        modified = True
                                    End If

                                    ' Apply text style if available
                                    If Not txtStyleId.IsNull Then
                                        attref.TextStyleId = txtStyleId
                                    End If

                                    ' CRITICAL: Use AdjustAlignment() for MiddleCenter attributes
                                    If modified AndAlso isMiddleCenter Then
                                        ' AdjustAlignment recalculates the position based on current justification
                                        Try
                                            attref.AdjustAlignment(acCurDb)
                                        Catch
                                            ' If AdjustAlignment fails, ignore
                                        End Try
                                    End If

                                Next

                            End If

                        End If

                    Next

                    ModifyViewPortCenter(vpIds, Layout(4), Layout(1), Layout(5))

                Catch es As Exception
                    MessageBox.Show(es.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                ' Save the changes made
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
                                            Dim font As New XFont(fontName, fontSize, XFontStyleEx.Bold)
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


    End Class
End Namespace
