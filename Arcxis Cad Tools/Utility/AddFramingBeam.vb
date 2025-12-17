' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Linq
Imports Autodesk.AutoCAD.Colors

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.AddFramingBeam))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class AddFramingBeam

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
        <CommandMethod("AFB")>
        Public Sub AddBeam()
            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor

            '' Lock the new document
            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    '' Get the current value from a system variable
                    Dim osm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("osmode")
                    Dim ech As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("cmdecho")
                    Dim otm As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("orthomode")
                    Dim cle As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("clayer")
                    Dim dsc As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("dimscale")
                    Dim atr As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("attreq")

                    '' Set system variable to new value
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("cmdecho", 0)
                    'Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("osmode", 0)
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("orthomode", 1)
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("attreq", 0)

                    Dim pts As New Point3dCollection()

                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

                    '' Prompt for the start point
                    pPtOpts.Message = vbLf & "Specify Beam Start Point: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptStart As Point3d = pPtRes.Value
                    pts.Add(ptStart)

                    '' Exit if the user presses ESC or cancels the command
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                    '' Prompt for the end point
                    pPtOpts.Message = vbLf & "Specify Beam End Point: "
                    pPtOpts.UseBasePoint = True
                    pPtOpts.BasePoint = ptStart
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim ptEnd As Point3d = pPtRes.Value
                    pts.Add(ptEnd)

                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                    Dim Beamdist As Double = Math.Sqrt((Math.Abs(ptEnd.X - ptStart.X) ^ 2) + (Math.Abs(ptEnd.Y - ptStart.Y) ^ 2))
                    'Windows.MessageBox.Show(Beamdist & " Beamdist")
                    'Module1.Distance = Beamdist

                    Dim Beamspan As Double = Beamdist / 12
                    Dim BeamAng As Double = Math.Atan2(ptEnd.Y - ptStart.Y, ptEnd.X - ptStart.X)
                    'Windows.MessageBox.Show(BeamAng & "BeamAng")
                    'Module1.Rotation = BeamAng

                    Dim mdpt As Point3d = New Point3d(ptStart.X + Beamdist * Math.Cos(BeamAng), ptStart.Y + Beamdist * Math.Sin(BeamAng), ptStart.Z)
                    'Module1.BeamMdpt = mdpt
                    'Windows.MessageBox.Show(mdpt.ToString & "mdpt")

                    Dim acBlkTbl As BlockTable = DirectCast(acTrans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim acBlkTblRec As BlockTableRecord = DirectCast(acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                    Dim acLine As Line
                    '' Create a line that starts at 5,5 and ends at 12,3
                    acLine = New Line(ptStart, ptEnd)
                    acLine.Layer = "S-FRM-BEAM"
                    'acLine.Linetype = "Hidden"
                    'Dim FLineAng As Double = acLine.Angle
                    'Dim FLineAngD As Integer = FLineAng * 180.0 / Math.PI

                    '' Create a polyline with two segments (3 points)
                    Using acPoly As Polyline = New Polyline()
                        acPoly.AddVertexAt(0, New Point2d(ptStart.X, ptStart.Y), 0, 0, 0)
                        acPoly.AddVertexAt(1, New Point2d(4, 2), 0, 0, 0)
                        acPoly.AddVertexAt(2, New Point2d(6, 4), 0, 0, 0)


                        acBlkTblRec.AppendEntity(acLine)
                        acTrans.AddNewlyCreatedDBObject(acLine, True)

                        '' Save the changes and dispose of the transaction
                        acTrans.Commit()


                    End Using

                End Using

            End Using



        End Sub

    End Class

End Namespace