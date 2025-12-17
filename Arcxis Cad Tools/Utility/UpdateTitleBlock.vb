' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Linq
Imports System.IO
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Colors


' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.UpdateTitleBlock))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!

    Public Class UpdateTitleBlock

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
        <CommandMethod("UTB")>
        Public Sub UpdateTitleBlock()

            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim TitleBlockIptList As New List(Of Point3d)
            Dim aced As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            AddTBFiles()

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, "TB-STAMP,TB-INFO")}
                    Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
                    Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

                    If prSelRes.Status = PromptStatus.OK Then

                        Dim SS As SelectionSet = prSelRes.Value

                        If SS IsNot Nothing Then

                            For Each brId As ObjectId In prSelRes.Value.GetObjectIds()

                                ' Open the block reference
                                Dim BlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)
                                Dim TblRec As BlockTableRecord = TryCast(acTrans.GetObject(BlockRef.DynamicBlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                Dim BlockName As String = TblRec.Name
                                Dim Blkacent As Entity = CType(acTrans.GetObject(brId, OpenMode.ForWrite, True), Entity)

                                If BlockName = "TB-STAMP" OrElse BlockName = "TB-INFO" Then

                                    If BlockName = "TB-STAMP" Then

                                        Dim BlockIpt As Point3d = BlockRef.Position

                                        TitleBlockIptList.Add(BlockIpt)

                                    End If

                                    Blkacent.Erase()

                                End If

                            Next

                            For Each BlockInsert As Point3d In TitleBlockIptList

                                ErasePline(acDoc, acCurDb, aced, acTrans, BlockInsert)

                                Dim BlockName As String = "DPIS_TitleBlock"
                                InsertTitleBlock(acDoc, acCurDb, aced, acTrans, BlockName, BlockInsert)

                                BlockInsert = New Point3d(BlockInsert.X + 1488, BlockInsert.Y + 754.9603, BlockInsert.Z)
                                BlockName = "DPIS_RevisionBlock"
                                InsertTitleBlock(acDoc, acCurDb, aced, acTrans, BlockName, BlockInsert)

                            Next

                        End If

                    End If

                    acDoc.Editor.Regen()

                    ' Save the new object to the database
                    acTrans.Commit()

                    If File.Exists("C:\Temp\DPIS_TitleBlock.dwg") Then File.Delete("C:\Temp\DPIS_TitleBlock.dwg")
                    If File.Exists("C:\Temp\DPIS_RevisionBlock.dwg") Then File.Delete("C:\Temp\DPIS_RevisionBlock.dwg")

                End Using

            End Using

        End Sub

        Private Sub AddTBFiles()

            '' Get the current document and database
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim ACed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim BlkNameList As List(Of String) = New List(Of String)
            Dim BlkDir As String = "C:\Temp\"

            BlkNameList.Add("DPIS_TitleBlock")
            BlkNameList.Add("DPIS_RevisionBlock")

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    For Each BlkName As String In BlkNameList

                        'Dim resourceName As String = "ArcxisLisp"

                        Dim filePathName As String = "C:\Temp\" & BlkName & ".dwg"

                        Try

                            ' Determine whether the directory exists.
                            If Directory.Exists(BlkDir) Then

                                'Return

                            Else

                                ' Try to create the directory.
                                Dim di As DirectoryInfo = Directory.CreateDirectory(BlkDir)

                            End If


                            Dim bytes = CType(My.Resources.ResourceManager.GetObject(BlkName), Byte())
                            If File.Exists(filePathName) Then File.Delete(filePathName)

                            Using stream = New FileStream(filePathName, FileMode.Create, FileAccess.Write)
                                stream.Write(bytes, 0, bytes.Length)
                                stream.Close()
                            End Using

                        Catch

                        End Try

                    Next

                    'Save the new object to the database
                    acTrans.Commit()

                End Using

            End Using

        End Sub

        Private Sub ErasePline(acDoc As Document, acCurDb As Database, aced As Editor, acTrans As Transaction, BlockInsert As Point3d)

            Dim TypValAr() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "S-ANNO-NOPLT")}
            Dim SelFtr As SelectionFilter = New SelectionFilter(TypValAr)
            Dim pSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(SelFtr)

            If (pSelRes.Status = PromptStatus.OK) Then

                Dim acSS As SelectionSet = pSelRes.Value

                '' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSS

                    Dim TBPLPointsList As New List(Of Point3d)
                    TBPLPointsList.Clear()
                    Dim SortedTBPLPointsList As New List(Of Point3d)
                    SortedTBPLPointsList.Clear()
                    Dim SortedTBPLPoints As New List(Of Point3d)
                    SortedTBPLPoints.Clear()
                    Dim SortedTBPLXPointsList As New List(Of Point3d)
                    SortedTBPLXPointsList.Clear()
                    Dim SortedTBPLXPoints As New List(Of Double)
                    SortedTBPLXPoints.Clear()
                    Dim SortedTBPLYPointsList As New List(Of Point3d)
                    SortedTBPLYPointsList.Clear()
                    Dim SortedTBPLYPoints As New List(Of Double)
                    SortedTBPLYPoints.Clear()
                    'Dim TBPLVerticeList As New List(Of Point3d)
                    'Dim SortedTBPLVerticeList As New List(Of Point3d)

                    Dim Typ As String = acSSObj.ObjectId.ObjectClass.Name()

                    If Typ = "AcDbPolyline" Then

                        Dim acObj As Object = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead, False, True)
                        Dim ID As String = acObj.objectid.ToString
                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite, False, True)

                        Dim lwp As Polyline = TryCast(acEnt, Polyline)
                        Dim vn As Integer = 0

                        If lwp IsNot Nothing Then
                            vn = lwp.NumberOfVertices

                            For i As Integer = 0 To vn - 1

                                Dim pt As Point3d = lwp.GetPoint3dAt(i)
                                TBPLPointsList.Add(pt)

                            Next

                        End If

                        For Each wpt As Point3d In TBPLPointsList

                            If Not SortedTBPLPointsList.Contains(wpt) Then

                                SortedTBPLPointsList.Add(wpt)

                            End If

                        Next

                        SortedTBPLPoints = (From pnt In SortedTBPLPointsList Order By pnt.X, pnt.Y, pnt.Z Select pnt).ToList
                        SortedTBPLXPointsList = (From Xpnt In SortedTBPLPointsList Order By Xpnt.X Select Xpnt).ToList
                        SortedTBPLYPointsList = (From Ypnt In SortedTBPLPointsList Order By Ypnt.Y Select Ypnt).ToList

                        For Each xpt As Point3d In SortedTBPLXPointsList

                            Dim XPOINT As Double = Math.Round(xpt.X, 10)

                            If Not SortedTBPLXPoints.Contains(XPOINT) Then

                                SortedTBPLXPoints.Add(XPOINT)

                            End If

                        Next

                        For Each ypt As Point3d In SortedTBPLYPointsList

                            Dim YPOINT As Double = Math.Round(ypt.Y, 10)

                            If Not SortedTBPLYPoints.Contains(YPOINT) Then

                                SortedTBPLYPoints.Add(YPOINT)

                            End If

                        Next

                        Dim XPOINTCount As Integer = SortedTBPLXPoints.Count
                        Dim YPOINTCount As Integer = SortedTBPLYPoints.Count

                        SortedTBPLXPoints.Sort()
                        SortedTBPLYPoints.Sort()

                        For Each pt As Double In SortedTBPLXPoints
                            'aced.WriteMessage(vbLf & pt.ToString())
                        Next

                        For Each pt As Double In SortedTBPLYPoints
                            'aced.WriteMessage(vbLf & pt.ToString())
                        Next

                        Dim TBPLStartX As Double = SortedTBPLXPoints(0)
                        Dim TBPLEndX As Double = SortedTBPLXPoints(SortedTBPLXPoints.Count - 1)
                        Dim TBPLBottY As Double = SortedTBPLYPoints(0)
                        Dim TBPLTopY As Double = SortedTBPLYPoints(SortedTBPLYPoints.Count - 1)
                        Dim TBPLBottStart As Point3d = New Point3d(TBPLStartX, TBPLBottY, 0)
                        Dim dist As Double = Math.Sqrt((Math.Abs(TBPLBottStart.X - BlockInsert.X) ^ 2) + (Math.Abs(TBPLBottStart.Y - BlockInsert.Y) ^ 2))

                        If dist < 57 Then

                            lwp.Erase()

                        End If

                    End If

                Next

            Else

                'aced.WriteMessage(vbLf & "No Selection Set")

            End If

        End Sub

        Private Sub InsertTitleBlock(acDoc As Document, acCurDb As Database, aced As Editor, acTrans As Transaction, BlockName As String, BlockInsert As Point3d)

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

            Dim sLayerName As String = "DPIS-TITLEBLOCK"

            If acLyrTbl.Has(sLayerName) = False Then
                Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                '' Assign the layer the ACI color 1 and a name
                acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 9)
                acLyrTblRec.Name = sLayerName

                '' Upgrade the Layer table for write
                acLyrTbl.UpgradeOpen()

                '' Append the new layer to the Layer table and the transaction
                acLyrTbl.Add(acLyrTblRec)
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)

                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sLayerName)

            ElseIf acLyrTbl.Has(sLayerName) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sLayerName)

            End If

            Dim blkRecId As ObjectId '= ObjectId.Null
            Dim FileName As String = "C:\Temp\" & BlockName & ".dwg"
            Dim TempDB As New Database(False, True)

            If IO.File.Exists(FileName) Then

                TempDB.ReadDwgFile(FileName, IO.FileShare.Read, True, Nothing)
                blkRecId = acCurDb.Insert(BlockName, TempDB, True)

                If blkRecId <> ObjectId.Null Then

                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(blkRecId, OpenMode.ForRead)

                    Using acBlkRef As New BlockReference(New Point3d(BlockInsert.X, BlockInsert.Y, BlockInsert.Z), acBlkTblRec.Id)

                        acBlkRef.Layer = "DPIS-TITLEBLOCK"
                        acBlkRef.ScaleFactors = New Scale3d(1, 1, 1)

                        Dim acCurSpaceBlkTblRec As BlockTableRecord
                        acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)

                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                        acTrans.AddNewlyCreatedDBObject(acBlkRef, True)

                        'ext = acBlkRef.GeometricExtents

                        If acBlkTblRec.HasAttributeDefinitions Then

                            For Each objID As ObjectId In acBlkTblRec

                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then
                                    Dim acAtt As AttributeDefinition = dbObj

                                    If Not acAtt.Constant Then
                                        Using acAttRef As New AttributeReference

                                            acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform)
                                            acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform)

                                            Dim tagvalue As String = acAttRef.Tag

                                            'If tagvalue = "DRAWN" Then

                                            'acAttRef.TextString = DRN

                                            'ElseIf tagvalue = "SCALE" Then

                                            'acAttRef.TextString = Module_DODI_Toolbar_Menu.Scale

                                            'ElseIf tagvalue = "DATE" Then

                                            'acAttRef.TextString = dwgdate

                                            'End If

                                            acBlkRef.AttributeCollection.AppendAttribute(acAttRef)
                                            acTrans.AddNewlyCreatedDBObject(acAttRef, True)

                                        End Using

                                    End If

                                End If

                            Next

                        End If

                    End Using

                End If

            Else

                'Throw exception for missing file 
                Throw New System.Exception("File " & FileName & " is not found")
                Return

            End If

        End Sub

    End Class

End Namespace