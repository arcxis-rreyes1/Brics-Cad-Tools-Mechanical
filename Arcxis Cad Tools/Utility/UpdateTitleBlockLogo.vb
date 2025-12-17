Imports System
Imports System.Linq
Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices


' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.UpdateTitleBlockLogo))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!

    Public Class UpdateTitleBlockLogo

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
        <CommandMethod("UDL")>
        Public Sub TBUpdate()
            ' Get the current database and start a transaction
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor
            Dim LineList As New List(Of Entity)
            Dim PLineList As New List(Of Entity)

            Application.SetSystemVariable("imageframe", 0)

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                ' Define the name and image to use
                Dim strImgName As String = "ArcxisLogo"
                Dim strFileName As String = "C:\temp\FullColor_HorizontalArcxis_689x350.png"

                Dim acRasterDef As RasterImageDef
                Dim bRasterDefCreated As Boolean = False
                Dim acImgDefId As ObjectId

                ' Get the image dictionary
                Dim acImgDctID As ObjectId = RasterImageDef.GetImageDictionary(acCurDb)

                ' Check to see if the dictionary does not exist, it not then create it
                If acImgDctID.IsNull Then
                    acImgDctID = RasterImageDef.CreateImageDictionary(acCurDb)
                End If

                ' Open the image dictionary
                Dim acImgDict As DBDictionary = acTrans.GetObject(acImgDctID, OpenMode.ForRead)

                ' Check to see if the image definition already exists
                If acImgDict.Contains(strImgName) Then
                    'acImgDefId = acImgDict.GetAt(strImgName)
                    'acRasterDef = acTrans.GetObject(acImgDefId, OpenMode.ForWrite)
                    'aced.WriteMessage(vbLf & "Image Exist")

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
                    acTrans.AddNewlyCreatedDBObject(acRasterDefNew, True)
                    acRasterDef = acRasterDefNew
                    bRasterDefCreated = True
                End If

                ' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                If acBlkTbl.Has("TEMP BLOCK") Then

                    ' Redefine the block if it exists
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl.Item("TEMP BLOCK"), OpenMode.ForWrite)

                    ' Step through each object in the block table record
                    For Each objID As ObjectId In acBlkTblRec

                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)

                        ' Revise the circle in the block
                        ' If TypeOf dbObj Is Text Then
                        'Dim acCirc As Circle = dbObj

                        'ans.GetObject(objID, OpenMode.ForWrite)
                        'arc.Radius = acCirc.Radius * 2

                        If objID.ObjectClass.DxfName = "TEXT" Then

                            Dim txt As DBText = CType(objID.GetObject(OpenMode.ForWrite), DBText)
                            Dim CoNam As String = txt.TextString
                            Dim CoNamPos As Point3d = txt.Position

                            Dim CoNamAng As Double = txt.Rotation
                            Dim CoNamAngD As Integer = CoNamAng * 180.0 / Math.PI

                            If CoNam = "POST-TENSION SOLUTIONS OF TEXAS, INC." AndAlso CoNamAngD = "90" Then

                                txt.TextString = "DPIS ENGINEERING,LLC."

                            ElseIf CoNam = "11000 CORPORATE CENTRE DR, SUITE 100  HOUSTON, TX 77041" AndAlso CoNamAngD = "90" Then

                                txt.TextString = "19450 STATE HIGHWAY 249, SUITE 300, HOUSTON, TX 77070"

                            ElseIf CoNam = "P: 713-996-9422   F: 713-996-9499" AndAlso CoNamAngD = "90" Then

                                txt.TextString = "OFFICE: (855) 500-3747"

                            ElseIf CoNam = "WWW.SOLUTIONS-PT.COM" AndAlso CoNamAngD = "90" Then

                                txt.TextString = "HTTPS://WWW.ARCXIS.COM"

                            End If


                        ElseIf objID.ObjectClass.DxfName = "MTEXT" Then

                            Dim MTlogo As MText = CType(objID.GetObject(OpenMode.ForWrite), MText)
                            Dim MTcont As String = MTlogo.Contents
                            aced.WriteMessage(vbLf & "MTcont = " & MTcont.ToString())
                            Dim MTAng As String = MTlogo.Rotation
                            Dim MTAngD As Integer = MTAng * 180.0 / Math.PI
                            aced.WriteMessage(vbLf & "MTAngD = " & MTAngD.ToString())
                            Dim MTSize As Double = Math.Round(MTlogo.Height, 4)
                            aced.WriteMessage(vbLf & "MTSize = " & MTSize.ToString())

                            If MTcont.Contains("PTS") Then

                                MTlogo.Erase()

                            End If

                        ElseIf objID.ObjectClass.DxfName = "HATCH" Then

                            Dim Hlogo As Hatch = CType(objID.GetObject(OpenMode.ForWrite), Hatch)
                            Dim PatNam As String = Hlogo.PatternName
                            aced.WriteMessage(vbLf & "PatNam = " & PatNam.ToString())
                            Dim PatAng As String = Hlogo.PatternAngle
                            Dim PatAngD As Integer = PatAng * 180.0 / Math.PI
                            aced.WriteMessage(vbLf & "PatAngD = " & PatAngD.ToString())
                            Dim PatOri As Point2d = Hlogo.Origin
                            aced.WriteMessage(vbLf & "PatOri = " & PatOri.ToString())
                            Hlogo.Erase()


                        ElseIf objID.ObjectClass.DxfName = "LINE" Then

                            Dim acEnt As Object = acTrans.GetObject(objID, OpenMode.ForRead, False, True)

                            Dim Linelogo As Line = CType(objID.GetObject(OpenMode.ForWrite), Line)
                            Dim LineStart As Point3d = Linelogo.StartPoint
                            Dim LineXipt As Double = LineStart.X
                            Dim LineYipt As Double = LineStart.Y
                            Dim linelen As Double = Linelogo.Length
                            Dim lineAng As Double = Linelogo.Angle
                            Dim lineAngD As Integer = lineAng * 180.0 / Math.PI
                            Dim Rlinelen As Double = Math.Round(linelen, 4)


                            If LineXipt > 15.53 AndAlso LineXipt < 16.27 AndAlso LineYipt > 8.229 AndAlso LineYipt < 9.033 Then

                                Linelogo.Erase()

                            Else

                                aced.WriteMessage(vbLf & "LineStart = " & LineStart.ToString())
                                aced.WriteMessage(vbLf & "LineXipt = " & LineXipt.ToString())
                                aced.WriteMessage(vbLf & "Lineyipt = " & LineYipt.ToString())
                                aced.WriteMessage(vbLf & "linelen = " & linelen.ToString())

                            End If

                        ElseIf objID.ObjectClass.DxfName = "LWPOLYLINE" Then

                            Dim acEnt As Object = acTrans.GetObject(objID, OpenMode.ForRead, False, True)


                            Dim PLinelogo As Polyline = CType(objID.GetObject(OpenMode.ForWrite), Polyline)
                            Dim PLineStart As Point3d = PLinelogo.StartPoint
                            Dim PLineXipt As Double = PLineStart.X
                            Dim PLineYipt As Double = PLineStart.Y
                            Dim PLineArea As Double = Math.Round(PLinelogo.Area, 2)

                            If PLineXipt > 15.53 AndAlso PLineXipt < 16.27 AndAlso PLineYipt > 8.229 AndAlso PLineYipt < 9.033 Then

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
                        If acCurDb.Measurement = MeasurementValue.English Then
                            width = New Vector3d((acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth) / 25.4, 0, 0)
                            height = New Vector3d(0, (acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / 25.4, 0)
                        Else
                            width = New Vector3d(acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0)
                            height = New Vector3d(0, acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0)
                        End If

                        ' Define the position for the image 
                        'Dim insPt As New Point3d(-45.0742, 787.1624, 0.0)

                        Dim insPt As New Point3d(16.5305, 8.1996, 0.0)

                        ' Define and assign a coordinate system for the image's orientation
                        Dim coordinateSystem As New CoordinateSystem3d(insPt, width * 1.11, height * 1.09)
                        acRaster.Orientation = coordinateSystem

                        ' Set the rotation angle for the image
                        acRaster.Rotation = 1.570796

                        ' Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acRaster)
                        acTrans.AddNewlyCreatedDBObject(acRaster, True)

                        ' Connect the raster definition and image together so the definition
                        ' does not appear as "unreferenced" in the External References palette.
                        RasterImage.EnableReactors(True)
                        acRaster.AssociateRasterDef(acRasterDef)

                        If bRasterDefCreated Then
                            acRasterDef.Dispose()
                        End If
                    End Using


                    Dim tsId = acCurDb.Textstyle
                    Dim ts = CType(acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
                    If ts.Has("ROMANS") Then

                        tsId = ts("ROMANS")

                    End If

                    Using acText As New DBText

                        acText.SetDatabaseDefaults()
                        acText.TextStyleId = tsId
                        acText.Height = 0.0504
                        acText.TextString = "DBA"
                        acText.Layer = "0"
                        acText.Rotation = 1.570796
                        acText.WidthFactor = 0.8
                        acText.Oblique = 0
                        acText.Justify = AttachmentPoint.BottomLeft
                        acText.IsMirroredInX = False
                        acText.IsMirroredInY = False
                        acText.HorizontalMode = TextHorizontalMode.TextLeft
                        acText.VerticalMode = TextVerticalMode.TextBottom
                        acText.AlignmentPoint = New Teigha.Geometry.Point3d(15.6874, 8.2728, 0)

                        acBlkTblRec.AppendEntity(acText)

                    End Using

                    ' Update existing block references
                    For Each objID As ObjectId In acBlkTblRec.GetBlockReferenceIds(False, True)
                        Dim acBlkRef As BlockReference = acTrans.GetObject(objID, OpenMode.ForWrite)
                        acBlkRef.RecordGraphicsModified(True)
                    Next

                    'MsgBox("Title Block has been revised.")
                End If

                ' Save the new object to the database
                acTrans.Commit()

                ' Dispose of the transaction
            End Using

        End Sub

    End Class

End Namespace