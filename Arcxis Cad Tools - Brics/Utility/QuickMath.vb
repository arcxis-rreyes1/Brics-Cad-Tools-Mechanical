' (C) Copyright 2011 by  
'
Imports System
Imports System.Linq
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.QuickMath))>
Namespace Arcxis_Cad_Tools

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class QuickMath

        <CommandMethod("HB")>
        Sub ReactionCalculator()

            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim aced As Editor = acDoc.Editor

            Try
                '' Prompt user to enter TotalLoad/Reaction
                Dim pdr As PromptDoubleResult = aced.GetDouble(vbCrLf & "Enter TotalLoad/Reaction: ")

                If pdr.Status <> PromptStatus.OK Then
                    Return
                End If

                Dim totalLoadReaction As Double = pdr.Value

                '' Store initial input for reference
                Dim initialInput As Double = totalLoadReaction

                '' Calculate LiveLoad = Result / 1.625
                Dim liveLoad As Double = totalLoadReaction / 1.625

                '' Calculate DeadLoad = TotalLoad/Reaction - LiveLoad
                Dim deadLoad As Double = totalLoadReaction - liveLoad

                '' Prompt user to select spacing
                Dim pko As PromptKeywordOptions = New PromptKeywordOptions(vbCrLf & "Select spacing [24/19.2/16]: ")
                pko.Keywords.Add("24")
                pko.Keywords.Add("19.2")
                pko.Keywords.Add("16")
                pko.AllowNone = False
                Dim pkr As PromptResult = aced.GetKeywords(pko)

                If pkr.Status <> PromptStatus.OK Then
                    Return
                End If

                Dim spacingDivisor As Double = 1.0
                If pkr.StringResult = "24" Then
                    spacingDivisor = 2.0
                ElseIf pkr.StringResult = "19.2" Then
                    spacingDivisor = 1.6
                ElseIf pkr.StringResult = "16" Then
                    spacingDivisor = 1.333
                End If

                '' Apply spacing divisor to all loads
                totalLoadReaction = totalLoadReaction / spacingDivisor
                liveLoad = liveLoad / spacingDivisor
                deadLoad = deadLoad / spacingDivisor

                '' Prompt user to select insertion point
                Dim ppo As PromptPointOptions = New PromptPointOptions(vbCrLf & "Select insertion point: ")
                Dim ppr As PromptPointResult = aced.GetPoint(ppo)

                If ppr.Status <> PromptStatus.OK Then
                    Return
                End If

                Dim insertionPoint As Point3d = ppr.Value

                '' Display the calculated values
                aced.WriteMessage(vbCrLf & "========== Reaction Calculator Results ==========" & vbCrLf)
                aced.WriteMessage("Initial Input: " & initialInput.ToString("F2") & vbCrLf)
                aced.WriteMessage("Spacing: " & pkr.StringResult & """" & vbCrLf)
                aced.WriteMessage("TotalLoad/Reaction: " & totalLoadReaction.ToString("F2") & vbCrLf)
                aced.WriteMessage("LiveLoad: " & liveLoad.ToString("F2") & vbCrLf)
                aced.WriteMessage("DeadLoad: " & deadLoad.ToString("F2") & vbCrLf)
                aced.WriteMessage("================================================" & vbCrLf)

                '' Create or get DPIS-Loads layer
                Dim layerName As String = "DPIS-Loads"
                Dim layerId As ObjectId = GetOrCreateLayer(acCurDb, layerName)

                '' Insert text at the selected point
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    '' Create multiline text with results
                    Dim acText As MText = New MText()
                    acText.SetDatabaseDefaults()
                    acText.Location = insertionPoint
                    acText.Contents = "Initial Input: " & initialInput.ToString("F2") & "\P" &
                                     "Spacing: " & pkr.StringResult & """" & "\P" &
                                     "TotalLoad/Reaction: " & totalLoadReaction.ToString("F2") & "\P" &
                                     "LiveLoad: " & liveLoad.ToString("F2") & "\P" &
                                     "DeadLoad: " & deadLoad.ToString("F2")
                    acText.TextHeight = 4
                    acText.LayerId = layerId

                    acBlkTblRec.AppendEntity(acText)
                    acTrans.AddNewlyCreatedDBObject(acText, True)
                    acTrans.Commit()
                End Using

            Catch ex As System.Exception
                aced.WriteMessage(vbCrLf & "Error: " & ex.Message & vbCrLf)
            End Try

        End Sub

        Private Function GetOrCreateLayer(acCurDb As Database, layerName As String) As ObjectId
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim layerTable As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                '' Check if layer already exists
                If layerTable.Has(layerName) Then
                    Return layerTable(layerName)
                End If

                '' Layer doesn't exist, so create it from layer 0
                layerTable.UpgradeOpen()
                Dim layer0 As LayerTableRecord = acTrans.GetObject(layerTable(0), OpenMode.ForRead)

                Dim newLayer As New LayerTableRecord()
                newLayer.Name = layerName
                newLayer.Color = layer0.Color
                newLayer.LineWeight = layer0.LineWeight
                newLayer.LinetypeObjectId = layer0.LinetypeObjectId
                newLayer.IsPlottable = False

                Dim layerId As ObjectId = layerTable.Add(newLayer)
                acTrans.AddNewlyCreatedDBObject(newLayer, True)
                acTrans.Commit()

                Return layerId
            End Using
        End Function
    End Class

End Namespace