' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.IO
Imports System.Drawing.Printing
Imports Autodesk.AutoCAD.Colors
Imports System.Windows.Documents
Imports Autodesk.AutoCAD.Interop
Imports System.Linq
Imports System.Text


' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.Arcxis_Titleblock_Menu))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!

    Public Class Arcxis_Titleblock_Menu

        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx

        <CommandMethod("ArcxisTitleBlock", CommandFlags.Modal)>
        Public Sub Arcxis_Titleblock_Menu() ' This method can have any name

            Dim frm As New Form_Arcxis_TB1
            ' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor

            Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                For Each entry As DBDictionaryEntry In layDict
                    Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                    layAndTab.Add(lay.TabOrder, lay.LayoutName)

                Next

                For Each layStr In layAndTab.Values


                    If layStr <> "Model" Then

                        frm.ListBox1.Items.Add(layStr)

                    End If

                Next

                If frm.ListBox1.Items.Count > 1 Then
                    frm.TextBox1.Enabled = False
                    frm.ComboBox1.Enabled = False
                Else
                    frm.TextBox1.Enabled = True
                    frm.ComboBox1.Enabled = True
                End If

                Dim lytab As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Dim alllayers As New ArrayList
                Dim cleanlayerstring As String
                For Each layer In lytab
                    Dim lytr As LayerTableRecord = acTrans.GetObject(layer, OpenMode.ForRead)
                    If lytr.Name Like "Master Seal File|S-SEAL-*" Then
                        cleanlayerstring = lytr.Name
                        alllayers.Add(cleanlayerstring.Substring(24, cleanlayerstring.Length - 24))
                        frm.ComboBox8.Items.Add(cleanlayerstring.Substring(24, cleanlayerstring.Length - 24))
                    End If
                Next

                Dim printDoc As New PrintDocument

                For Each strPrinter As [String] In PrinterSettings.InstalledPrinters

                    frm.ComboBox6.Items.Add(strPrinter)

                    If (printDoc.PrinterSettings.IsDefaultPrinter()) Then
                        frm.ComboBox6.Text = printDoc.PrinterSettings.PrinterName

                        For Each printerformat In printDoc.PrinterSettings.PaperSizes()

                            If printerformat.PaperName.Contains("Letter") OrElse printerformat.PaperName.Contains("8.5x11") Then

                                frm.TextBox39.Text = printerformat.PaperName

                            ElseIf printerformat.PaperName.Contains("11") OrElse printerformat.PaperName.Contains("11x17") OrElse printerformat.PaperName.Contains("11 x 17") OrElse printerformat.PaperName.Contains("11"" x 17""") OrElse printerformat.PaperName.Contains("11""x17""") OrElse printerformat.PaperName.Contains("Tabloid") Then

                                frm.TextBox40.Text = printerformat.PaperName

                            ElseIf printerformat.PaperName.Contains("24x36") OrElse printerformat.PaperName.Contains("24 x 36") OrElse printerformat.PaperName.Contains("24"" x 36""") OrElse printerformat.PaperName.Contains("24""x36""") Then

                                frm.TextBox41.Text = printerformat.PaperName

                            End If

                        Next

                    End If

                Next

                If Trim(frm.TextBox39.Text).Length = 0 Then

                    frm.TextBox39.Text = "Paper Size not Available"
                End If

                If Trim(frm.TextBox40.Text).Length = 0 Then

                    frm.TextBox40.Text = "Paper Size not Available"

                End If

                If Trim(frm.TextBox41.Text).Length = 0 Then

                    frm.TextBox41.Text = "Paper Size not Available"

                End If

                ' Abort the changes to the database
                acTrans.Commit()

            End Using

            If frm.ListBox1.Items.Count > 1 Then
                frm.Button8.Enabled = True
                frm.Button8.BackColor = System.Drawing.SystemColors.Control
            Else
                frm.Button8.Enabled = False
                frm.Button8.BackColor = System.Drawing.SystemColors.Control
            End If

            frm.PlotPDF11x17.Enabled = False
            frm.PDF24x36.Enabled = False
            frm.ComboBox8.BackColor = System.Drawing.Color.Red

            CheckBlockTable()
            PurgeBlockMethod()

            Dim CurTrustPath As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TRUSTEDPATHS")
            Dim CurPaths As String = LCase(CurTrustPath)
            Dim RemCurTrustPathList As List(Of String) = New List(Of String)
            Dim isittrue As Integer = 0
            RemCurTrustPathList = CurPaths.Split(";").ToList
            Dim CleanedRemCurTrustPathList As List(Of String) = New List(Of String)

            For Each entry In RemCurTrustPathList
                If entry <> "" Then
                    If Not CleanedRemCurTrustPathList.Contains(entry) Then
                        CleanedRemCurTrustPathList.Add(entry)
                    End If
                End If
            Next

            If CleanedRemCurTrustPathList.Count = 0 Then
                isittrue = 1
            End If

            For Each curpath As String In CleanedRemCurTrustPathList

                If Not LCase(curpath).Contains("/Shared/Arcxis/Engineering/") Then

                    isittrue = isittrue + 1

                End If

            Next

            If isittrue > 0 Then

                SupplementalPaths(CleanedRemCurTrustPathList)

            End If

            ATB_FirstLayoutName = frm.ListBox1.Items(0)

            frm.ShowDialog()

        End Sub

        Private Sub CheckBlockTable()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor

            ed.WriteMessage(vbLf & "CheckBlockTable = ")

            ' get the working Database
            Dim db = HostApplicationServices.WorkingDatabase

            Using acLckDoc As DocumentLock = doc.LockDocument()

                ' start a transaction
                Using tr = db.TransactionManager.StartTransaction()
                    ' open the block table
                    Dim bt = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                    If bt.Has("DPIS TitleBlock") = False Then


                        If bt.Has("A$C429541B6") Then

                            Module_Arcxis_TB.DpisOldName = "A$C429541B6"
                            Module_Arcxis_TB.DpisNewName = "DPIS TitleBlock"
                            RenameTitleBlock3()

                        End If

                        If bt.Has("A$C48014A83") Then

                            Module_Arcxis_TB.DpisOldName = "A$C48014A83"
                            Module_Arcxis_TB.DpisNewName = "DPIS TitleBlock"
                            RenameTitleBlock3()

                        End If

                        If bt.Has("DPIS_Border") Then

                            Module_Arcxis_TB.DpisOldName = "DPIS_Border"
                            Module_Arcxis_TB.DpisNewName = "DPIS TitleBlock"

                            RenameTitleBlock3()

                        End If

                        If bt.Has("A$C083C5A08") Then

                            Module_Arcxis_TB.DpisOldName = "A$C083C5A08"
                            Module_Arcxis_TB.DpisNewName = "DPIS TitleBlock 24x36"

                            RenameTitleBlock3()

                        End If

                    End If

                    If bt.Has("DPIS TitleBlock") Then

                        Module_Arcxis_TB.DpisOldName = "DPIS TitleBlock"
                        Module_Arcxis_TB.DpisNewName = "Arcxis Title Block"


                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 1)

                        ' Define the name and image to use
                        Dim strImgName As String = "ArcxisLogo"
                        Dim strFileName As String = "\\egnytedrive\energyinspectors\Shared\Arcxis\Engineering\Drafting Standards\CAD Blocks\FullColor_HorizontalArcxis_689x350.png"
                        'Dim strFileName As String = "C:\temp\FullColor_HorizontalArcxis_689x350.png"

                        Dim acRasterDef As RasterImageDef
                        Dim bRasterDefCreated As Boolean = False
                        Dim acImgDefId As ObjectId

                        ' Get the image dictionary
                        Dim acImgDctID As ObjectId = RasterImageDef.GetImageDictionary(db)

                        ' Check to see if the dictionary does not exist, it not then create it
                        If acImgDctID.IsNull Then
                            acImgDctID = RasterImageDef.CreateImageDictionary(db)
                        End If

                        ' Open the image dictionary
                        Dim acImgDict As DBDictionary = tr.GetObject(acImgDctID, OpenMode.ForRead)

                        ' Check to see if the image definition already exists
                        If acImgDict.Contains(strImgName) Then
                            'acImgDefId = acImgDict.GetAt(strImgName)
                            'acRasterDef = tr.GetObject(acImgDefId, OpenMode.ForWrite)

                            Exit Sub

                        Else
                            ' Create a raster image definition
                            Dim acRasterDefNew As New RasterImageDef

                            ' Set the source for the image file
                            acRasterDefNew.SourceFileName = strFileName

                            ' Load the image into memory
                            acRasterDefNew.Load()

                            ' Add the image definition to the dictionary
                            acImgDict.UpgradeOpen()
                            acImgDefId = acImgDict.SetAt(strImgName, acRasterDefNew)

                            tr.AddNewlyCreatedDBObject(acRasterDefNew, True)

                            acRasterDef = acRasterDefNew

                            bRasterDefCreated = True
                        End If

                        ' Redefine the block if it exists
                        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(bt.Item("DPIS TitleBlock"), OpenMode.ForWrite)
                        Dim LineList As New List(Of Entity)
                        Dim PLineList As New List(Of Entity)

                        ' Step through each object in the block table record
                        For Each objID As ObjectId In acBlkTblRec

                            Dim dbObj As DBObject = tr.GetObject(objID, OpenMode.ForRead)


                            If objID.ObjectClass.DxfName = "TEXT" Then

                                Dim txt As DBText = CType(objID.GetObject(OpenMode.ForWrite), DBText)
                                Dim CoNam As String = txt.TextString
                                Dim CoNamPos As Point3d = txt.Position
                                Dim CoNamHei As Double = Math.Round(txt.Height, 4)

                                Dim CoNamAng As Double = txt.Rotation
                                Dim CoNamAngD As Integer = CoNamAng * 180.0 / Math.PI

                                If CoNam = "DPIS ENGINEERING,LLC." AndAlso CoNamAngD = "0" Then

                                    txt.Rotation = 1.570796

                                    '' Create a matrix and move the circle using a vector from (0,0,0) to (2,0,0)
                                    Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                                    Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(0.2102, 0.3598, 0))
                                    txt.TransformBy(Matrix3d.Displacement(acVec3d))

                                ElseIf CoNam = "TX REG. NO. 16584" AndAlso CoNamAngD = "0" Then

                                    txt.TextString = "TX REG. NO. F-16584"
                                    txt.Rotation = 1.570796
                                    'txt.Position = New Point3d(16.2821, 7.0089, 0)

                                    Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                                    Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(0.3015, 0.4512, 0))
                                    txt.TransformBy(Matrix3d.Displacement(acVec3d))

                                ElseIf CoNam = "19450 STATE HIGHWAY 249" AndAlso CoNamAngD = "0" Then

                                    txt.Rotation = 1.570796
                                    Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                                    Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(0.394, 0.5264, 0))
                                    txt.TransformBy(Matrix3d.Displacement(acVec3d))

                                ElseIf CoNam = "HOUSTON, TEXAS 77070" AndAlso CoNamAngD = "0" Then

                                    txt.Rotation = 1.570796

                                    Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                                    Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(0.4539, 0.5863, 0))
                                    txt.TransformBy(Matrix3d.Displacement(acVec3d))

                                ElseIf CoNam = "PH: 281.351.0048  FAX: 281.351.0148" AndAlso CoNamAngD = "0" Then

                                    txt.Rotation = 1.570796

                                    Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                                    Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(0.5128, 0.6452, 0))
                                    txt.TransformBy(Matrix3d.Displacement(acVec3d))

                                End If

                            ElseIf objID.ObjectClass.DxfName = "HATCH" Then

                                Dim Hlogo As Hatch = CType(objID.GetObject(OpenMode.ForWrite), Hatch)
                                Dim PatNam As String = Hlogo.PatternName
                                Dim PatAng As String = Hlogo.PatternAngle
                                Dim PatAngD As Integer = PatAng * 180.0 / Math.PI
                                Dim PatOri As Point2d = Hlogo.Origin

                                Hlogo.Erase()

                            ElseIf objID.ObjectClass.DxfName = "LINE" Then

                                Dim acEnt As Object = tr.GetObject(objID, OpenMode.ForRead, False, True)

                                Dim Linelogo As Line = CType(objID.GetObject(OpenMode.ForWrite), Line)
                                Dim LineStart As Point3d = Linelogo.StartPoint
                                Dim LineXipt As Double = LineStart.X
                                Dim LineYipt As Double = LineStart.Y
                                Dim linelen As Double = Linelogo.Length
                                Dim lineAng As Double = Linelogo.Angle
                                Dim lineAngD As Integer = lineAng * 180.0 / Math.PI
                                Dim Rlinelen As Double = Math.Round(linelen, 4)

                                If linelen < 0.03 Then

                                    Linelogo.Erase()

                                End If

                                If lineAngD = 161 OrElse lineAngD = 29 OrElse lineAngD = 30 Then

                                    Linelogo.Erase()

                                End If

                                If lineAngD = 90 AndAlso Rlinelen = 0.1614 Then

                                    Linelogo.Erase()

                                End If

                                If lineAngD = 90 AndAlso Rlinelen = 0.1451 Then

                                    Linelogo.Erase()

                                End If

                                If LineXipt > 15.5315 AndAlso LineXipt < 16.3708 AndAlso LineYipt > 6.7375 AndAlso LineYipt < 7.6873 Then


                                    Linelogo.Erase()

                                Else

                                End If

                            ElseIf objID.ObjectClass.DxfName = "LWPOLYLINE" Then

                                Dim acEnt As Object = tr.GetObject(objID, OpenMode.ForRead, False, True)


                                Dim PLinelogo As Polyline = CType(objID.GetObject(OpenMode.ForWrite), Polyline)
                                Dim PLineStart As Point3d = PLinelogo.StartPoint
                                Dim PLineXipt As Double = PLineStart.X
                                Dim PLineYipt As Double = PLineStart.Y

                                If PLineXipt > 15.4393 AndAlso PLineXipt < 16.3708 AndAlso PLineYipt > 6.3322 AndAlso PLineYipt < 7.6873 Then

                                    PLinelogo.Erase()


                                End If


                            End If

                        Next

                        ' Create the new image and assign it the image definition
                        Using acRaster As New RasterImage
                            acRaster.ImageDefId = acImgDefId

                            ' Use ImageWidth and ImageHeight to get the size of the image in pixels (1024 x 768).
                            ' Use ResolutionMMPerPixel to determine the number of millimeters in a pixel so you 
                            ' can convert the size of the drawing into other units or millimeters based on the 
                            ' drawing units used in the current drawing.

                            ' Define the width and height of the image
                            Dim width As Vector3d
                            Dim height As Vector3d

                            ' Check to see if the measurement is set to English (Imperial) or Metric units
                            If db.Measurement = MeasurementValue.English Then
                                width = New Vector3d((acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth) / 25.4, 0, 0)
                                height = New Vector3d(0, (acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / 25.4, 0)
                            Else
                                width = New Vector3d(acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0)
                                height = New Vector3d(0, acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0)
                            End If

                            ' Define the position for the image 
                            Dim insPt As New Point3d(16.0022, 6.2376, 0.0)

                            ' Define and assign a coordinate system for the image's orientation
                            Dim coordinateSystem As New CoordinateSystem3d(insPt, width * 0.5993, height * 0.5993)
                            acRaster.Orientation = coordinateSystem
                            ' Set the rotation angle for the image
                            acRaster.Rotation = 1.570796
                            acRaster.Layer = "0"
                            acRaster.ColorIndex = 7

                            ' Add the new object to the block table record and the transaction
                            acBlkTblRec.AppendEntity(acRaster)
                            tr.AddNewlyCreatedDBObject(acRaster, True)

                            ' Connect the raster definition and image together so the definition
                            ' does not appear as "unreferenced" in the External References palette.
                            RasterImage.EnableReactors(True)
                            acRaster.AssociateRasterDef(acRasterDef)

                            If bRasterDefCreated Then
                                acRasterDef.Dispose()
                            End If
                        End Using

                        Dim tsId = db.Textstyle
                        Dim ts = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
                        If ts.Has("ROMANS") Then

                            tsId = ts("ROMANS")

                        End If

                        Using acText As New DBText

                            acText.SetDatabaseDefaults()
                            acText.TextStyleId = tsId
                            acText.Height = 0.0277
                            acText.TextString = "DBA"
                            acText.Layer = "0"
                            acText.ColorIndex = 7
                            acText.Rotation = 1.570796
                            acText.WidthFactor = 0.8
                            acText.Oblique = 0
                            acText.Justify = AttachmentPoint.BottomLeft
                            acText.IsMirroredInX = False
                            acText.IsMirroredInY = False
                            acText.HorizontalMode = Autodesk.AutoCAD.DatabaseServices.TextHorizontalMode.TextLeft
                            acText.VerticalMode = Autodesk.AutoCAD.DatabaseServices.TextVerticalMode.TextBottom
                            acText.AlignmentPoint = New Autodesk.AutoCAD.Geometry.Point3d(15.5708, 6.25, 0)

                            acBlkTblRec.AppendEntity(acText)

                        End Using

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 0)

                        ' Update existing block references
                        For Each objID As ObjectId In acBlkTblRec.GetBlockReferenceIds(False, True)
                            Dim acBlkRef As BlockReference = tr.GetObject(objID, OpenMode.ForWrite)
                            acBlkRef.RecordGraphicsModified(True)

                        Next

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 0)

                        RenameTitleBlock4()

                    Else

                        Exit Sub

                    End If

                    tr.Commit()

                End Using

            End Using

        End Sub

        Private Sub RenameTitleBlock3()

            ' Get the current database and start a transaction
            Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database
            Dim acdoc As Document = Application.DocumentManager.MdiActiveDocument

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                ' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                'acBlkTbl.UpgradeOpen()

                ' Redefine the block if it exists
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl.Item(DpisOldName), OpenMode.ForWrite)
                acBlkTblRec.UpgradeOpen()

                ' Step through each object in the block table record


                acBlkTblRec.Name = DpisNewName

                Dim btrr = DirectCast(acTrans.GetObject(acBlkTbl(DpisNewName), OpenMode.ForWrite), BlockTableRecord)

                ' Update existing block references
                For Each objID As ObjectId In btrr.GetBlockReferenceIds(False, True)
                    Dim acBlkRef As BlockReference = acTrans.GetObject(objID, OpenMode.ForWrite)
                    acBlkRef.RecordGraphicsModified(True)

                Next

                acBlkTblRec.DowngradeOpen()


                ' Save the new object to the database
                acTrans.Commit()

                ' Dispose of the transaction
            End Using

        End Sub

        Private Sub RenameTitleBlock4()

            ' Get the current database and start a transaction
            Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database
            Dim acdoc As Document = Application.DocumentManager.MdiActiveDocument

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                ' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                'acBlkTbl.UpgradeOpen()

                ' Redefine the block if it exists
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl.Item(DpisOldName), OpenMode.ForWrite)
                acBlkTblRec.UpgradeOpen()

                ' Step through each object in the block table record


                acBlkTblRec.Name = DpisNewName

                Dim btrr = DirectCast(acTrans.GetObject(acBlkTbl(DpisNewName), OpenMode.ForWrite), BlockTableRecord)

                ' Update existing block references
                For Each objID As ObjectId In btrr.GetBlockReferenceIds(False, True)
                    Dim acBlkRef As BlockReference = acTrans.GetObject(objID, OpenMode.ForWrite)
                    acBlkRef.RecordGraphicsModified(True)

                Next

                acBlkTblRec.DowngradeOpen()


                ' Save the new object to the database
                acTrans.Commit()

                ' Dispose of the transaction
            End Using

        End Sub

        Public Sub PurgeBlockMethod()

            Dim acdoc As Document = Application.DocumentManager.MdiActiveDocument
            Using db As Database = HostApplicationServices.WorkingDatabase
                Using acLckDoc As DocumentLock = acdoc.LockDocument()
                    Using tr As Transaction = db.TransactionManager.StartTransaction()
                        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                        Try
                            Dim tv() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, "DESIGN X")}
                            Dim sf As SelectionFilter = New SelectionFilter(tv)
                            Dim psr As PromptSelectionResult = ed.SelectAll(sf)
                            If psr.Status = PromptStatus.OK Then

                                Dim ss As SelectionSet = psr.Value
                                Dim idarr As ObjectId() = ss.GetObjectIds()
                                Dim id As ObjectId
                                For Each id In idarr
                                    Dim e As Entity = CType(tr.GetObject(id, OpenMode.ForWrite, True), Entity)
                                    e.Erase()
                                Next id
                            End If
                            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                            If Not bt.Has("DESIGN X") Then
                                'MessageBox.Show("block MyBlock does not exist")
                                Return
                            End If
                            Dim btr As BlockTableRecord = tr.GetObject(bt("DESIGN X"), OpenMode.ForWrite, True, True)
                            Dim idcoll As ObjectIdCollection = New ObjectIdCollection()
                            idcoll.Add(btr.ObjectId)
                            db.Purge(idcoll)
                            btr.Erase(True)
                            tr.Commit()
                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            'MessageBox.Show(ex.StackTrace)
                        End Try
                    End Using
                End Using
            End Using
        End Sub

    End Class

End Namespace