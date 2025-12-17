Imports System
Imports System.Linq
Imports System.Windows.Forms
Imports Arcxis_Cad_Tools.Arcxis_Cad_Tools
Imports System.Drawing
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports Color = Autodesk.AutoCAD.Colors.Color

Public Class Framing_Vault_Transfer

    Public ReadOnly Property PackageFrWbValue As String
        Get
            If WSFW.Checked Then Return "WSFW"
            If RFR.Checked Then Return "RFR" ' <-- change to your RB2 control name if needed
            Return ""
        End Get
    End Property
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ElevNumber As Integer = ElevCount.Text
        Dim FramingPages As Integer = FramingCount.Text
        Dim BracingPages115 As Integer = BracingCount115.Text
        Dim BracingPages130 As Integer = BracingCount130.Text
        Dim BracingPages142 As Integer = BracingCount142.Text
        Dim Builder As String = TextBox1.Text
        Dim PlanName As String = TextBox2.Text
        Dim TDIPages As Integer = TdiCount.Text
        Dim SheetLabeling As String = SheetLabels.SelectedItem
        'Dim SetupSwings As String = PlanSwing.SelectedItem
        Dim PaperSpaceScale As String = ViewportScale.SelectedItem
        Dim frm1 As New Form_Arcxis_TB4
        Dim Elevations As New List(Of String)
        Dim SequenceCounter As Integer 'this is for labeleing left to right to keep the page block in order later on
        Dim Loops As New List(Of (First As String, Second As Integer))
        Dim SetupSwings As New List(Of String)
        Dim NumberOfNeededLayouts As Integer = 0
        Dim SealLoop As New List(Of String)
        Dim FramingPagesList

        ' Save the single custom property once, based on radio selection
        SetCustomDwgPropReliable("PackageFRWB", PackageFrWbValue)

        ' Hide the parent form before showing the child form
        Me.Hide()


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''meta data collection form
        '''


        'Dim frm As New Form_FramingDefinitions(FramingPages)

        'Dim frmresult = frm.ShowDialog()

        'If frmresult = DialogResult.Cancel Then
        '    ' Show the parent form again if canceled
        '    Me.Show()
        '    Me.BringToFront()
        '    Exit Sub

        'Else
        '    FramingPagesList = frm.GroupedEntries
        'End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Loops.Add(("Framing", FramingPages))
        Loops.Add(("142 Bracing", BracingPages142))
        Loops.Add(("130 Bracing", BracingPages130))
        Loops.Add(("115 Bracing", BracingPages115))
        Loops.Add(("Windstorm", TDIPages))

        For Each selecteditem In SealsList.CheckedItems

            SealLoop.Add(selecteditem)

        Next

        'checking to see if we need more than 10 layouts with automation
        For Each PlanType In Loops
            If PlanType.Second > NumberOfNeededLayouts Then
                NumberOfNeededLayouts = PlanType.Second
            End If
        Next

        If PlanSwing.SelectedItem = "Right" Then
            SetupSwings.Add("Right")
        ElseIf PlanSwing.SelectedItem = "Left" Then
            SetupSwings.Add("Left")
        Else
            SetupSwings.Add("Right")
            SetupSwings.Add("Left")
        End If

        Me.Close()
        Me.Dispose()

        For x = 1 To ElevNumber

            frm1.Label1.Text = "Elevation #" & x

            frm1.TextBox1.Text = ""

            frm1.ShowDialog()

            If ATB_CustomLayoutLetter = "" Then

                Exit Sub

            Else

                Elevations.Add(UCase(ATB_CustomLayoutLetter))

            End If

        Next

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Dim otm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("orthomode")
        Dim oldos As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("osmode")
        '' Set system variable to new value
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("orthomode", 1)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("osmode", 32)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 1)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 0)

        Dim opts As New PromptPointOptions(vbLf & "Select top left corner of first layout: ")
        Dim res1 As PromptPointResult = aced.GetPoint(opts)
        If res1.Status <> PromptStatus.OK Then

            Application.SetSystemVariable("OSMODE", oldos)
            Application.SetSystemVariable("ORTHOMODE", otm)
            Return

        End If
        Dim StartingPoint As Point3d = res1.Value

        ' Ask for second point
        opts.Message = vbLf & "Select top left corner second layout:"
        Dim res2 As PromptPointResult = aced.GetPoint(opts)

        If res2.Status <> PromptStatus.OK Then

            Application.SetSystemVariable("OSMODE", oldos)
            Application.SetSystemVariable("ORTHOMODE", otm)
            Return

        End If
        Dim pt2 As Point3d = res2.Value

        Dim pt3 As Point3d
        If ElevNumber > 1 Then

            ' Ask for third point
            opts.Message = vbLf & "Select top left corner of next elevation:"
            Dim res3 As PromptPointResult = aced.GetPoint(opts)
            If res3.Status <> PromptStatus.OK Then

                Application.SetSystemVariable("OSMODE", oldos)
                Application.SetSystemVariable("ORTHOMODE", otm)
                Return

            End If

            pt3 = res3.Value

        End If

        Dim pt4 As Point3d
        If SetupSwings.Count = 2 Then

            ' Ask for fourth point
            opts.Message = vbLf & "Select top left corner of next swing:"
            Dim res4 As PromptPointResult = aced.GetPoint(opts)
            If res4.Status <> PromptStatus.OK Then

                Application.SetSystemVariable("OSMODE", oldos)
                Application.SetSystemVariable("ORTHOMODE", otm)
                Return

            End If

            pt4 = res4.Value

        End If


        Dim PanOver As Double = pt2.X - StartingPoint.X
        Dim PanDown As Double = StartingPoint.Y - pt3.Y
        Dim InsertionPoint As Point3d
        Dim SwingCounter As Integer = 1
        Dim ItemCount As Integer

        EnsurePlanInfoAtOrigin(StartingPoint, Builder, PlanName, SealLoop)

        For Each swing In SetupSwings

            For x = 0 To Elevations.Count - 1

                SequenceCounter = 1

                For Each item In Loops ' this is for the three options so far FRAME, BRACE, TDI

                    ItemCount = 1

                    If item.Second > 0 Then

                        For z = 1 To item.Second 'THIS IS TO LOOP THROUGH EACH DICIPLINE

                            InsertionPoint = New Point3d(StartingPoint.X + (PanOver * (SequenceCounter - 1)), StartingPoint.Y - (PanDown * x), 0)

                            If item.First = "142 Bracing" Then

                                PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), "Bracing", swing, PaperSpaceScale, "142 MPH")

                            ElseIf item.First = "130 Bracing" Then

                                PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), "Bracing", swing, PaperSpaceScale, "130 MPH")

                            ElseIf item.First = "115 Bracing" Then

                                PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), "Bracing", swing, PaperSpaceScale, "115 MPH")

                            Else

                                If item.First = "Framing" Then

                                    'Dim entry = FramingPagesList(z - 1)
                                    'Dim framingFloor As String = entry.Item1   ' GroupBox.Text
                                    'Dim framingMaterial As String = entry.Item2 ' ListBox.SelectedItem

                                    'PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), "Framing", swing, PaperSpaceScale, "", framingFloor, framingMaterial)

                                    PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), "Framing", swing, PaperSpaceScale)

                                Else

                                    PageBlockInsert(InsertionPoint, ItemCount, Elevations.Item(x), item.First, swing, PaperSpaceScale)

                                End If
                            End If

                            WorkSpaceBlockInsert(InsertionPoint)
                            SequenceCounter += 1
                            ItemCount += 1

                        Next

                    End If

                Next

            Next

            SwingCounter += 1
            StartingPoint = pt4
        Next
        Dim CTBForPages As String = ComboBox1.SelectedItem
        If PaperSpaceSetup.SelectedItem = "Yes" Then

            FileManipulation.CreateLayoutsWithTitleblock(NumberOfNeededLayouts, True, CTBForPages)

        End If

        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("orthomode", otm)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("osmode", oldos)

        SetCustomDwgPropReliable("PLAN", PlanName)
        SetCustomDwgPropReliable("BUILDER", Builder)
        SetCustomDwgPropReliable("PLAN TYPE", "Framing")
        SetCustomDwgPropReliable("SHEET LABELING", SheetLabeling)

        Dim parts As New List(Of String)()

        For Each item As String In SealLoop
            If Not String.IsNullOrWhiteSpace(item) Then
                parts.Add(item.Trim())
            End If
        Next

        Dim result As String = String.Join(", ", parts)

        SetCustomDwgPropReliable("STAMPS", result)

        Using acDoc.LockDocument()
            Dim folder As String = CStr(Application.GetSystemVariable("DWGPREFIX"))
            Dim name As String = CStr(Application.GetSystemVariable("DWGNAME"))
            Dim path = System.IO.Path.Combine(folder, name)

            acCurDb.SaveAs(path, True, DwgVersion.Current, acCurDb.SecurityParameters)

        End Using

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

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
    Public Sub CreatePageBlock()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)

                Dim blockName As String = "PAGE"

                If bt.Has(blockName) Then
                    acEd.WriteMessage(vbLf & "Block '" & blockName & "' already exists.")
                    Return
                End If

                ' Create new block definition
                Dim btr As New BlockTableRecord()
                btr.Name = blockName
                btr.Origin = Point3d.Origin

                ' Add new block definition to BlockTable
                bt.Add(btr)
                acTrans.AddNewlyCreatedDBObject(btr, True)

                ' Create vertical line (0,14.2833) to (0,-14.2833)
                Dim vertLine As New Line(New Point3d(0, 14.2833, 0), New Point3d(0, -14.2833, 0))
                vertLine.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                vertLine.Layer = "0"
                btr.AppendEntity(vertLine)
                acTrans.AddNewlyCreatedDBObject(vertLine, True)

                ' Create horizontal line (-14.2833,0) to (14.2833,0)
                Dim horizLine As New Line(New Point3d(-14.2833, 0, 0), New Point3d(14.2833, 0, 0))
                horizLine.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                horizLine.Layer = "0"
                btr.AppendEntity(horizLine)
                acTrans.AddNewlyCreatedDBObject(horizLine, True)

                ' Create attribute definition
                Dim attDef As New AttributeDefinition()
                attDef.Position = New Point3d(81.6266, 7.6399, 0)
                attDef.Height = 48.0092
                attDef.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attDef.Justify = AttachmentPoint.BaseLeft
                attDef.AdjustAlignment(acCurDb)
                attDef.Tag = "LXS"
                attDef.Prompt = "LONG AND SHORT BEAMS"
                attDef.TextString = "#x#"
                attDef.Rotation = 0
                attDef.WidthFactor = 0.8
                attDef.Constant = False
                attDef.Verifiable = False
                attDef.Invisible = False
                attDef.LockPositionInBlock = True
                attDef.Layer = "0"
                attDef.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                attDef.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attDef.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                Dim attDef1 As New AttributeDefinition()
                attDef1.Position = New Point3d(363.5456, 7.6399, 0)
                attDef1.Height = 48.0092
                attDef1.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attDef1.Justify = AttachmentPoint.BaseLeft
                attDef1.Tag = "TND"
                attDef1.Prompt = "SBT OR DBT"
                attDef1.TextString = "TND"
                attDef1.Rotation = 0
                attDef1.WidthFactor = 0.8
                attDef1.Verifiable = True
                attDef1.Invisible = False
                attDef1.LockPositionInBlock = True
                attDef1.Layer = "0"
                attDef1.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                attDef1.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attDef1.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                Dim attDef2 As New AttributeDefinition()
                attDef2.Position = New Point3d(741.7474, 7.6399, 0)
                attDef2.Height = 48.0092
                attDef2.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attDef2.Justify = AttachmentPoint.BaseLeft
                attDef2.Tag = "MOD"
                attDef2.Prompt = "MODIFIERS"
                attDef2.TextString = "MOD"
                attDef2.Rotation = 0
                attDef2.WidthFactor = 0.8
                attDef2.Verifiable = True
                attDef2.Invisible = False
                attDef2.LockPositionInBlock = False
                attDef2.Layer = "0"
                attDef2.Color = Color.FromColorIndex(ColorMethod.ByAci, 2)
                attDef2.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attDef2.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                Dim attdef3 As New AttributeDefinition()
                attdef3.Position = New Point3d(-6.3382, -471.8082, 0)
                attdef3.Height = 48.0092
                attdef3.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef3.Justify = AttachmentPoint.BaseLeft
                attdef3.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef3.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase
                attdef3.Tag = "SW"
                attdef3.Prompt = "LEFT OR RIGHT"
                attdef3.TextString = "SW"
                attdef3.Rotation = 1.5708
                attdef3.WidthFactor = 0.8
                attdef3.Verifiable = True
                attdef3.Invisible = False
                attdef3.LockPositionInBlock = False
                attdef3.Layer = "0"
                attdef3.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)

                Dim attDef4 As New AttributeDefinition()
                attDef4.Position = New Point3d(-74.6154, -57.1109, 0)
                attDef4.Height = 48.0092
                attDef4.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attDef4.Justify = AttachmentPoint.BaseLeft
                attDef4.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attDef4.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase
                attDef4.Tag = "PRNT"
                attDef4.Prompt = "PRINTING ORDER"
                attDef4.TextString = "1"
                attDef4.Rotation = 0
                attDef4.WidthFactor = 0.8
                attDef4.Verifiable = True
                attDef4.Invisible = False
                attDef4.LockPositionInBlock = True
                attDef4.Layer = "0"
                attDef4.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)

                Dim attdef5 As New AttributeDefinition()
                attdef5.Position = New Point3d(-6.3382, -786.075, 0)
                attdef5.Height = 48.0092
                attdef5.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef5.Justify = AttachmentPoint.BaseLeft
                attdef5.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef5.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase
                attdef5.Tag = "SIZE"
                attdef5.Prompt = "SIZE OF PAGE"
                attdef5.TextString = "11x17"
                attdef5.Rotation = 1.5708
                attdef5.WidthFactor = 0.8
                attdef5.Verifiable = True
                attdef5.Invisible = False
                attdef5.LockPositionInBlock = True
                attdef5.Layer = "0"
                attdef5.Color = Color.FromColorIndex(ColorMethod.ByAci, 2)

                Dim attdef6 As New AttributeDefinition()
                attdef6.Position = New Point3d(1275.0359, 7.6399, 0)
                attdef6.Height = 23.9247
                attdef6.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef6.Justify = AttachmentPoint.BaseLeft
                attdef6.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef6.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase
                attdef6.Tag = "OPTIONS"
                attdef6.Prompt = "OPTIONS"
                attdef6.TextString = "OPTIONS"
                attdef6.Rotation = 0
                attdef6.WidthFactor = 0.8
                attdef6.Verifiable = True
                attdef6.Invisible = False
                attdef6.LockPositionInBlock = True
                attdef6.Layer = "0"
                attdef6.Color = Color.FromColorIndex(ColorMethod.ByAci, 2)

                Dim attdef7 As New AttributeDefinition()
                attdef7.Position = New Point3d(-6.3382, -225.8515, 0)
                attdef7.Height = 48.0092
                attdef7.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef7.Justify = AttachmentPoint.BaseLeft
                attdef7.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef7.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase
                attdef7.Tag = "ELEV"
                attdef7.Prompt = "ELEVATION"
                attdef7.TextString = "ELEV"
                attdef7.Rotation = 1.5708
                attdef7.WidthFactor = 0.8
                attdef7.Verifiable = True
                attdef7.Invisible = False
                attdef7.LockPositionInBlock = False
                attdef7.Layer = "0"
                attdef7.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)



                btr.AppendEntity(attDef)
                acTrans.AddNewlyCreatedDBObject(attDef, True)
                btr.AppendEntity(attDef1)
                acTrans.AddNewlyCreatedDBObject(attDef1, True)
                btr.AppendEntity(attDef2)
                acTrans.AddNewlyCreatedDBObject(attDef2, True)
                btr.AppendEntity(attdef3)
                acTrans.AddNewlyCreatedDBObject(attdef3, True)
                btr.AppendEntity(attDef4)
                acTrans.AddNewlyCreatedDBObject(attDef4, True)
                btr.AppendEntity(attdef5)
                acTrans.AddNewlyCreatedDBObject(attdef5, True)
                btr.AppendEntity(attdef6)
                acTrans.AddNewlyCreatedDBObject(attdef6, True)
                btr.AppendEntity(attdef7)
                acTrans.AddNewlyCreatedDBObject(attdef7, True)

                acTrans.Commit()

                acEd.WriteMessage(vbLf & "Block '" & blockName & "' created successfully.")

            End Using

        End Using

    End Sub



    Private Sub PageBlockInsert(insPt As Point3d, SequenceCounter As Integer, Elevation As String, LayoutType As String, Swing As String, PaperSpaceScale As String, Optional ByVal Options As String = "", Optional ByVal FramingFloor As String = "", Optional ByVal FramingMaterial As String = "")

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim blkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)


                If Not blkTbl.Has("PAGE") Then

                    CreatePageBlock()

                End If

                Dim blkDefId As ObjectId = blkTbl("PAGE") ' Your block name here
                Dim modelSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Create the block reference and add it to model space
                Dim blkRef As New BlockReference(insPt, blkDefId)

                Dim lt As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                ' Check if layer "temp" exists
                If Not lt.Has("S-ANNO-AUTOMATION") Then
                    ' Upgrade layer table for writing
                    lt.UpgradeOpen()

                    ' Create new layer
                    Dim newLayer As New LayerTableRecord()
                    newLayer.Name = "S-ANNO-AUTOMATION"
                    newLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 0) ' Set color to 41
                    newLayer.IsPlottable = False ' Set layer to non-plot

                    ' Add to layer table and transaction
                    lt.Add(newLayer)
                    acTrans.AddNewlyCreatedDBObject(newLayer, True)

                End If

                blkRef.Layer = "S-ANNO-AUTOMATION"

                modelSpace.AppendEntity(blkRef)

                Dim blkDef As BlockTableRecord = TryCast(blkTbl("PAGE").GetObject(OpenMode.ForRead), BlockTableRecord)

                For Each id As ObjectId In blkDef

                    Dim obj As DBObject = id.GetObject(OpenMode.ForRead)
                    Dim attDef As AttributeDefinition = TryCast(obj, AttributeDefinition)

                    If (attDef IsNot Nothing) AndAlso (Not attDef.Constant) Then

                        'This is a non-constant AttributeDefinition 
                        'Create a new AttributeReference
                        Using attRef As New AttributeReference()

                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)

                            Dim tagvalue As String = attRef.Tag

                            If tagvalue = "LXS" Then

                                If LayoutType = "Framing" Then

                                    attRef.TextString = FramingFloor
                                    attRef.Height = 24

                                Else

                                    attRef.TextString = ""

                                End If


                            ElseIf tagvalue = "TND" Then


                                If LayoutType = "Framing" Then

                                    attRef.TextString = FramingMaterial
                                    attRef.Height = 24.0

                                Else

                                    attRef.TextString = ""

                                End If

                            ElseIf tagvalue = "MOD" Then

                                attRef.TextString = LayoutType

                            ElseIf tagvalue = "OPTIONS" Then

                                If Options <> "" Then

                                    attRef.TextString = Options

                                Else

                                    attRef.TextString = ""

                                End If

                            ElseIf tagvalue = "PRNT" Then

                                attRef.TextString = SequenceCounter

                            ElseIf tagvalue = "ELEV" Then

                                attRef.TextString = Elevation

                            ElseIf tagvalue = "SW" Then

                                attRef.TextString = Swing

                            ElseIf tagvalue = "SIZE" Then

                                attRef.TextString = PaperSpaceScale

                            End If

                            'Add the AttributeReference to the BlockReference
                            blkRef.AttributeCollection.AppendAttribute(attRef)

                            acTrans.AddNewlyCreatedDBObject(attRef, True)

                        End Using

                    End If

                Next

                acTrans.AddNewlyCreatedDBObject(blkRef, True)

                acTrans.Commit()

            End Using

        End Using
    End Sub

    Public Sub CreateWorkSpaceBlock()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)

                Dim blockName As String = "ArcxisWorkingSpace"

                If bt.Has(blockName) Then
                    acEd.WriteMessage(vbLf & "Block '" & blockName & "' already exists.")
                    Return
                End If

                ' Create new block definition
                Dim btr1 As New BlockTableRecord()
                btr1.Name = blockName
                btr1.Origin = Point3d.Origin

                bt.Add(btr1)
                acTrans.AddNewlyCreatedDBObject(btr1, True)
                ' Add a rectangle (as polyline)
                Dim rect As New Polyline()
                rect.Layer = "0"
                rect.AddVertexAt(0, New Point2d(0, 0), 0, 0, 0)
                rect.AddVertexAt(1, New Point2d(1419.837, 0), 0, 0, 0)
                rect.AddVertexAt(2, New Point2d(1419.837, -1008.0), 0, 0, 0)
                rect.AddVertexAt(3, New Point2d(0, -1008.0), 0, 0, 0)
                rect.AddVertexAt(4, New Point2d(0, 0), 0, 0, 0)
                rect.Closed = True
                btr1.AppendEntity(rect)
                acTrans.AddNewlyCreatedDBObject(rect, True)

                acTrans.Commit()

            End Using

        End Using

    End Sub


    Private Sub WorkSpaceBlockInsert(insPt As Point3d)

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim blkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)


                If Not blkTbl.Has("ArcxisWorkingSpace") Then

                    CreateWorkSpaceBlock()

                End If

                Dim blkDefId As ObjectId = blkTbl("ArcxisWorkingSpace") ' Your block name here
                Dim modelSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Create the block reference and add it to model space
                Dim blkRef As New BlockReference(insPt, blkDefId)

                Dim lt As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                ' Check if layer "temp" exists
                If Not lt.Has("temp") Then
                    ' Upgrade layer table for writing
                    lt.UpgradeOpen()

                    ' Create new layer
                    Dim newLayer As New LayerTableRecord()
                    newLayer.Name = "temp"
                    newLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 41) ' Set color to 41
                    newLayer.IsPlottable = False ' Set layer to non-plot

                    ' Add to layer table and transaction
                    lt.Add(newLayer)
                    acTrans.AddNewlyCreatedDBObject(newLayer, True)

                End If

                blkRef.Layer = "temp"

                modelSpace.AppendEntity(blkRef)

                acTrans.AddNewlyCreatedDBObject(blkRef, True)

                acTrans.Commit()

            End Using

        End Using
    End Sub
    Public Shared Sub CreatePlanInfoBlock()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)

                Dim blockName As String = "AutoPlanPrintInfo"

                If bt.Has(blockName) Then
                    acEd.WriteMessage(vbLf & "Block '" & blockName & "' already exists.")
                    Return
                End If

                ' Create new block definition
                Dim btr1 As New BlockTableRecord()
                btr1.Name = blockName
                btr1.Origin = Point3d.Origin

                bt.Add(btr1)
                acTrans.AddNewlyCreatedDBObject(btr1, True)

                ' Create vertical line (0,6) to (-92,6)
                Dim vertLine As New Line(New Point3d(0, 6, 0), New Point3d(-92, 6, 0))
                vertLine.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                vertLine.Layer = "0"
                btr1.AppendEntity(vertLine)
                acTrans.AddNewlyCreatedDBObject(vertLine, True)

                ' Create vertical line (0,6) to (-92,6)
                Dim vertLine1 As New Line(New Point3d(0, 12, 0), New Point3d(-92, 12, 0))
                vertLine1.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                vertLine1.Layer = "0"
                btr1.AppendEntity(vertLine1)
                acTrans.AddNewlyCreatedDBObject(vertLine1, True)

                ' Create vertical line (0,6) to (-92,6)
                Dim vertLine2 As New Line(New Point3d(-72, 0, 0), New Point3d(-72, 18, 0))
                vertLine2.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                vertLine1.Layer = "0"
                btr1.AppendEntity(vertLine2)
                acTrans.AddNewlyCreatedDBObject(vertLine2, True)


                ' Add a rectangle (as polyline)
                Dim rect As New Polyline()
                rect.Layer = "0"
                rect.AddVertexAt(0, New Point2d(0, 0), 0, 0, 0)
                rect.AddVertexAt(1, New Point2d(-92, 0), 0, 0, 0)
                rect.AddVertexAt(2, New Point2d(-92, 18.0), 0, 0, 0)
                rect.AddVertexAt(3, New Point2d(0, 18), 0, 0, 0)
                rect.AddVertexAt(4, New Point2d(0, 0), 0, 0, 0)
                rect.Closed = True
                btr1.AppendEntity(rect)
                acTrans.AddNewlyCreatedDBObject(rect, True)

                ' Create attribute definition
                Dim attDef As New AttributeDefinition()
                attDef.Position = New Point3d(-70.9474, 13, 0)
                attDef.Height = 4
                attDef.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attDef.Justify = AttachmentPoint.BaseLeft
                attDef.AdjustAlignment(acCurDb)
                attDef.Tag = "BUILDER"
                attDef.Prompt = "BUILDER"
                attDef.TextString = "###"
                attDef.Rotation = 0
                attDef.WidthFactor = 0.8
                attDef.Constant = False
                attDef.Verifiable = False
                attDef.Invisible = False
                attDef.LockPositionInBlock = True
                attDef.Layer = "0"
                attDef.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                attDef.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attDef.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                ' Create attribute definition
                Dim attdef1 As New AttributeDefinition()
                attdef1.Position = New Point3d(-70.9474, 7, 0)
                attdef1.Height = 4
                attdef1.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef1.Justify = AttachmentPoint.BaseLeft
                attdef1.AdjustAlignment(acCurDb)
                attdef1.Tag = "PLAN"
                attdef1.Prompt = "PLAN"
                attdef1.TextString = "###"
                attdef1.Rotation = 0
                attdef1.WidthFactor = 0.8
                attdef1.Constant = False
                attdef1.Verifiable = False
                attdef1.Invisible = False
                attdef1.LockPositionInBlock = True
                attdef1.Layer = "0"
                attdef1.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                attdef1.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef1.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                Dim attdef2 As New AttributeDefinition()
                attdef2.Position = New Point3d(-70.9474, 1, 0)
                attdef2.Height = 4
                attdef2.TextStyleId = acCurDb.Textstyle ' Uses current text style ("romans" if already set)
                attdef2.Justify = AttachmentPoint.BaseLeft
                attdef2.AdjustAlignment(acCurDb)
                attdef2.Tag = "STAMPS"
                attdef2.Prompt = "STAMPS"
                attdef2.TextString = "###"
                attdef2.Rotation = 0
                attdef2.WidthFactor = 0.8
                attdef2.Constant = False
                attdef2.Verifiable = False
                attdef2.Invisible = False
                attdef2.LockPositionInBlock = True
                attdef2.Layer = "0"
                attdef2.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                attdef2.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                attdef2.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBase

                btr1.AppendEntity(attDef)
                acTrans.AddNewlyCreatedDBObject(attDef, True)
                btr1.AppendEntity(attdef1)
                acTrans.AddNewlyCreatedDBObject(attdef1, True)
                btr1.AppendEntity(attdef2)
                acTrans.AddNewlyCreatedDBObject(attdef2, True)

                ' Add static single-line text entities
                Dim dbText1 As New DBText()
                dbText1.Position = New Point3d(-82.0809, 15, 0)
                dbText1.Height = 4 ' text height in drawing units
                dbText1.TextString = "BUILDER"
                dbText1.Layer = "0"
                dbText1.WidthFactor = 0.8
                dbText1.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                dbText1.Justify = AttachmentPoint.MiddleCenter
                dbText1.AlignmentPoint = dbText1.Position
                btr1.AppendEntity(dbText1)
                acTrans.AddNewlyCreatedDBObject(dbText1, True)

                Dim dbText2 As New DBText()
                dbText2.Position = New Point3d(-82.0809, 9, 0)
                dbText2.Height = 4
                dbText2.TextString = "PLAN"
                dbText2.Layer = "0"
                dbText2.WidthFactor = 0.8
                dbText2.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                dbText2.Justify = AttachmentPoint.MiddleCenter
                dbText2.AlignmentPoint = dbText2.Position
                btr1.AppendEntity(dbText2)
                acTrans.AddNewlyCreatedDBObject(dbText2, True)

                Dim dbText3 As New DBText()
                dbText3.Position = New Point3d(-82.0809, 3, 0)
                dbText3.Height = 4
                dbText3.TextString = "STAMPS"
                dbText3.Layer = "0"
                dbText3.WidthFactor = 0.8
                dbText3.Color = Color.FromColorIndex(ColorMethod.ByAci, 0)
                dbText3.Justify = AttachmentPoint.MiddleCenter
                dbText3.AlignmentPoint = dbText3.Position
                btr1.AppendEntity(dbText3)
                acTrans.AddNewlyCreatedDBObject(dbText3, True)

                acTrans.Commit()

            End Using

        End Using

    End Sub
    Public Shared Sub PlanInfoBlockInsert(insPt As Point3d, Builder As String, Plan As String, SealLoop As List(Of String))

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Autodesk.AutoCAD.EditorInput.Editor = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim blkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)


                If Not blkTbl.Has("AutoPlanPrintInfo") Then

                    CreatePlanInfoBlock()

                End If

                Dim blkDefId As ObjectId = blkTbl("AutoPlanPrintInfo") ' Your block name here
                Dim modelSpace As BlockTableRecord = acTrans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Create the block reference and add it to model space
                Dim blkRef As New BlockReference(insPt, blkDefId)

                Dim lt As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                blkRef.Layer = "DefPoints"

                modelSpace.AppendEntity(blkRef)

                Dim blkDef As BlockTableRecord = TryCast(blkTbl("AutoPlanPrintInfo").GetObject(OpenMode.ForRead), BlockTableRecord)

                For Each id As ObjectId In blkDef

                    Dim obj As DBObject = id.GetObject(OpenMode.ForRead)
                    Dim attDef As AttributeDefinition = TryCast(obj, AttributeDefinition)

                    If (attDef IsNot Nothing) AndAlso (Not attDef.Constant) Then

                        'This is a non-constant AttributeDefinition 
                        'Create a new AttributeReference
                        Using attRef As New AttributeReference()

                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)

                            Dim tagvalue As String = attRef.Tag

                            If tagvalue = "BUILDER" Then

                                attRef.TextString = Builder

                            ElseIf tagvalue = "PLAN" Then

                                attRef.TextString = Plan

                            ElseIf tagvalue = "STAMPS" Then

                                Dim parts As New List(Of String)()

                                For Each item As String In SealLoop
                                    If Not String.IsNullOrWhiteSpace(item) Then
                                        parts.Add(item.Trim())
                                    End If
                                Next

                                Dim result As String = String.Join(", ", parts)

                                attRef.TextString = result

                            End If

                            'Add the AttributeReference to the BlockReference
                            blkRef.AttributeCollection.AppendAttribute(attRef)

                            acTrans.AddNewlyCreatedDBObject(attRef, True)

                        End Using

                    End If

                Next

                acTrans.AddNewlyCreatedDBObject(blkRef, True)

                acTrans.Commit()

            End Using

        End Using
    End Sub

    Public Shared Sub EnsurePlanInfoAtOrigin(insPt As Point3d, Builder As String, Plan As String, SealLoop As List(Of String))
        Const BlockName As String = "AutoPlanPrintInfo"
        Dim doc = Application.DocumentManager.MdiActiveDocument
        Dim db = doc.Database
        Dim ed = doc.Editor

        Using doc.LockDocument()
            Using tr = db.TransactionManager.StartTransaction()

                Dim bt = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If Not bt.Has(BlockName) Then
                    CreatePlanInfoBlock()
                End If

                Dim msId = bt(BlockTableRecord.ModelSpace)
                Dim ms = DirectCast(tr.GetObject(msId, OpenMode.ForWrite), BlockTableRecord)
                Dim defId = bt(BlockName)
                Dim defBtr = DirectCast(tr.GetObject(defId, OpenMode.ForRead), BlockTableRecord)

                ' Find all PlanInfoBlock refs in Model Space
                Dim atOrigin As New List(Of ObjectId)()
                Dim elsewhere As New List(Of ObjectId)()

                For Each entId As ObjectId In ms
                    Dim br = TryCast(tr.GetObject(entId, OpenMode.ForRead), BlockReference)
                    If br Is Nothing Then Continue For
                    If Not IsBlockRefOfName(br, BlockName, tr) Then Continue For

                    If IsNearOrigin(br.Position) Then
                        atOrigin.Add(entId)
                    Else
                        elsewhere.Add(entId)
                    End If
                Next

                ' Case 1: already at origin -> do nothing
                If atOrigin.Count > 0 Then
                    tr.Commit()
                    Return
                End If

                ' Case 2: none found anywhere -> do nothing (you'll insert with your normal command)
                If elsewhere.Count = 0 Then
                    PlanInfoBlockInsert(insPt, Builder, Plan, SealLoop)
                    tr.Commit()
                    Return
                End If

                ' Case 3: found elsewhere -> move to origin (preserving best attribute values)
                Dim bestValues As Dictionary(Of String, String) = Nothing
                Dim bestScore As Integer = -1

                For Each id In elsewhere
                    Dim br = DirectCast(tr.GetObject(id, OpenMode.ForRead), BlockReference)
                    Dim vals = CollectAttrValues(br, tr)
                    Dim score = 0
                    For Each kv In vals
                        If Not String.IsNullOrWhiteSpace(kv.Value) Then score += 1
                    Next
                    If score > bestScore Then
                        bestScore = score
                        bestValues = vals
                    End If
                Next

                ' Delete all existing refs in Model Space
                For Each id In elsewhere
                    Dim br = DirectCast(tr.GetObject(id, OpenMode.ForWrite), BlockReference)
                    br.Erase(True)
                Next

                ' Insert new at origin with preserved values
                Dim ins As New BlockReference(Point3d.Origin, defId)
                ms.AppendEntity(ins)
                tr.AddNewlyCreatedDBObject(ins, True)

                CreateBlockAttributesFromDef(ins, defBtr, bestValues, tr, db)
                ins.RecordGraphicsModified(True)

                tr.Commit()
            End Using
        End Using
    End Sub

    Private Shared Function IsBlockRefOfName(br As BlockReference, target As String, tr As Transaction) As Boolean
        Dim name As String = br.Name
        If br.IsDynamicBlock Then
            Dim dyn = DirectCast(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
            name = dyn.Name
        End If
        Return String.Equals(name, target, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function IsNearOrigin(pt As Point3d, Optional tol As Double = 0.000001) As Boolean
        Return Math.Abs(pt.X) <= tol AndAlso Math.Abs(pt.Y) <= tol AndAlso Math.Abs(pt.Z) <= tol
    End Function

    Private Shared Function CollectAttrValues(br As BlockReference, tr As Transaction) As Dictionary(Of String, String)
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim atts = br.AttributeCollection
        If atts Is Nothing Then Return map
        For Each attId As ObjectId In atts
            Dim ar = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
            If ar Is Nothing Then Continue For
            If Not String.IsNullOrEmpty(ar.Tag) Then
                map(ar.Tag) = ar.TextString
            End If
        Next
        Return map
    End Function

    Private Shared Sub CreateBlockAttributesFromDef(br As BlockReference,
                                               defBtr As BlockTableRecord,
                                               priorValues As Dictionary(Of String, String),
                                               tr As Transaction,
                                               db As Database)
        Dim acoll = br.AttributeCollection
        For Each id As ObjectId In defBtr
            Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
            If ent Is Nothing Then Continue For
            Dim ad = TryCast(ent, AttributeDefinition)
            If ad Is Nothing OrElse ad.Constant Then Continue For

            Dim ar As New AttributeReference()
            ar.SetAttributeFromBlock(ad, br.BlockTransform)

            Dim val As String = Nothing
            If priorValues IsNot Nothing AndAlso priorValues.TryGetValue(ad.Tag, val) AndAlso Not String.IsNullOrWhiteSpace(val) Then
                ar.TextString = val
            Else
                ar.TextString = ad.TextString
            End If

            ar.AdjustAlignment(db)
            acoll.AppendAttribute(ar)
            tr.AddNewlyCreatedDBObject(ar, True)
        Next
    End Sub

    Private Sub SheetLabels_SelectedValueChanged(sender As Object, e As EventArgs) Handles SheetLabels.SelectedValueChanged
        SheetLabels.BackColor = SystemColors.Window
        Button1.Visible = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub PackageToggle_Click(sender As Object, e As EventArgs) Handles RFR.Click, WSFW.Click
        Dim rb = DirectCast(sender, RadioButton)

        ' If already checked, allow user to uncheck and re-enable the other
        If rb.Checked Then
            rb.Checked = False
            RFR.Enabled = True
            WSFW.Enabled = True
        Else
            ' Activate this one and lock the other
            rb.Checked = True
            If rb Is RFR Then
                WSFW.Enabled = False
            Else
                RFR.Enabled = False
            End If
        End If
    End Sub

    Private Sub Framing_Vault_Transfer_Load(sender As Object, e As EventArgs) Handles Me.Load
        RFR.AutoCheck = False
        WSFW.AutoCheck = False
    End Sub
End Class