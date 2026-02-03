Imports System
Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Teigha.PlotSettingsValidator
Imports Bricscad.PlottingServices
Imports Exception = Teigha.Runtime.Exception
Imports System.Reflection.Metadata.Ecma335

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.BlockReset))>
Namespace Arcxis_Cad_Tools
    Public Class BlockReset

        Private Class BlockViewState
            Public Property Position As Point3d
            Public Property Rotation As Double
            Public Property ScaleX As Double
            Public Property ScaleY As Double
            Public Property ScaleZ As Double
            Public Property Layer As String
            Public Property Linetype As String
            Public Property LineWeight As LineWeight
            Public Property Color As Color
            Public Property Transparency As Transparency
            Public Property LinetypeScale As Double
            Public Property DynamicBlockProperties As Dictionary(Of String, Object)
            Public Property IsDynamic As Boolean
            Public Property BlockName As String
        End Class

        Public Sub ResetBlockWithViewStatePreservation(blockId As ObjectId)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Get the block reference
                    Dim blockRef As BlockReference = CType(trans.GetObject(blockId, OpenMode.ForRead), BlockReference)

                    ' Capture the current viewstate
                    Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                    ' Reset the block reference
                    Dim blockDef As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    blockRef.UpgradeOpen()
                    blockRef.ResetBlock()

                    ' Restore the viewstate
                    RestoreBlockViewState(blockRef, viewState)

                    trans.Commit()
                    ' ed.WriteMessage(vbLf & "Dynamic block reset with viewstate preserved." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        Public Sub ExplodeBlockWithViewStatePreservation(blockId As ObjectId)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Get the block reference
                    Dim blockRef As BlockReference = CType(trans.GetObject(blockId, OpenMode.ForRead), BlockReference)

                    ' Capture the current viewstate
                    Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                    ' Get the parent block table record (modelspace or paperspace)
                    Dim parentBtr As BlockTableRecord = CType(trans.GetObject(blockRef.OwnerId, OpenMode.ForWrite), BlockTableRecord)

                    ' Create a collection to hold the exploded entities
                    Dim explodedEntities As New DBObjectCollection()

                    ' Upgrade to write mode and explode
                    blockRef.UpgradeOpen()
                    blockRef.Explode(explodedEntities)

                    ' Apply the viewstate transformations to exploded entities
                    ApplyViewStateToExplodedEntities(explodedEntities, viewState, parentBtr, trans)

                    ' Erase the original block reference
                    blockRef.Erase()

                    trans.Commit()
                    'ed.WriteMessage(vbLf & $"Block exploded into {explodedEntities.Count} entities with viewstate preserved." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        Private Sub ApplyViewStateToExplodedEntities(explodedEntities As DBObjectCollection, viewState As BlockViewState, parentBtr As BlockTableRecord, trans As Transaction)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor

            ed.WriteMessage(vbLf & "=== APPLYING VIEWSTATE TO EXPLODED ENTITIES ===" & vbLf)
            ed.WriteMessage($"  Applying transformation to {explodedEntities.Count} entities" & vbLf)

            For Each entity As Entity In explodedEntities
                Try
                    ' Apply layer
                    entity.Layer = viewState.Layer
                    ed.WriteMessage($"  - {entity.GetType().Name} layer set to: {viewState.Layer}" & vbLf)

                    ' Apply color
                    entity.Color = viewState.Color
                    ed.WriteMessage($"  - {entity.GetType().Name} color applied" & vbLf)

                    ' Apply linetype
                    entity.Linetype = viewState.Linetype
                    ed.WriteMessage($"  - {entity.GetType().Name} linetype set to: {viewState.Linetype}" & vbLf)

                    ' Apply line weight
                    entity.LineWeight = viewState.LineWeight
                    ed.WriteMessage($"  - {entity.GetType().Name} line weight applied" & vbLf)

                    ' Apply transparency
                    entity.Transparency = viewState.Transparency
                    ed.WriteMessage($"  - {entity.GetType().Name} transparency applied" & vbLf)

                    ' Apply linetype scale
                    entity.LinetypeScale = viewState.LinetypeScale
                    ed.WriteMessage($"  - {entity.GetType().Name} linetype scale applied" & vbLf)

                    ' Add entity to modelspace/paperspace
                    parentBtr.AppendEntity(entity)
                    trans.AddNewlyCreatedDBObject(entity, True)

                Catch ex As Exception
                    ed.WriteMessage($"  - Warning: Could not apply properties to {entity.GetType().Name}: {ex.Message}" & vbLf)
                End Try
            Next

            ed.WriteMessage("========================" & vbLf)
        End Sub

        Private Function CaptureBlockViewState(blockRef As BlockReference) As BlockViewState
            Dim dynamicProps As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

            ' Capture dynamic block properties if it's a dynamic block
            If blockRef.IsDynamicBlock Then
                Try
                    ' Use the actual Bricscad API for dynamic block properties
                    Dim propCollection = blockRef.DynamicBlockReferencePropertyCollection

                    If propCollection IsNot Nothing AndAlso propCollection.Count > 0 Then
                        For Each prop As DynamicBlockReferenceProperty In propCollection
                            Try
                                Dim propName = prop.PropertyName
                                Dim propValue = prop.Value
                                dynamicProps(propName) = propValue
                            Catch
                                ' Skip problematic properties
                            End Try
                        Next
                    End If
                Catch ex As Exception
                    ' If error capturing properties, continue anyway
                End Try
            End If

            Dim viewState As New BlockViewState With {
                .Position = blockRef.Position,
                .Rotation = blockRef.Rotation,
                .ScaleX = blockRef.ScaleFactors.X,
                .ScaleY = blockRef.ScaleFactors.Y,
                .ScaleZ = blockRef.ScaleFactors.Z,
                .Layer = blockRef.Layer,
                .Linetype = blockRef.Linetype,
                .LineWeight = blockRef.LineWeight,
                .Color = blockRef.Color,
                .Transparency = blockRef.Transparency,
                .LinetypeScale = blockRef.LinetypeScale,
                .DynamicBlockProperties = dynamicProps,
                .IsDynamic = blockRef.IsDynamicBlock,
                .BlockName = blockRef.Name
            }

            ' Debug output
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            'ed.WriteMessage(vbLf & "=== CAPTURED VIEWSTATE ===" & vbLf)
            'ed.WriteMessage($"  Block Name: {viewState.BlockName}" & vbLf)
            'ed.WriteMessage($"  Is Dynamic: {viewState.IsDynamic}" & vbLf)
            'ed.WriteMessage($"  Position: X={viewState.Position.X:F4}, Y={viewState.Position.Y:F4}, Z={viewState.Position.Z:F4}" & vbLf)
            'ed.WriteMessage($"  Rotation: {viewState.Rotation:F6} radians ({viewState.Rotation * 180 / Math.PI:F2} degrees)" & vbLf)
            'ed.WriteMessage($"  Scale: X={viewState.ScaleX:F6}, Y={viewState.ScaleY:F6}, Z={viewState.ScaleZ:F6}" & vbLf)
            'ed.WriteMessage($"  Layer: {viewState.Layer}" & vbLf)
            'ed.WriteMessage($"  Linetype: {viewState.Linetype}" & vbLf)
            'ed.WriteMessage($"  LineWeight: {viewState.LineWeight}" & vbLf)
            'ed.WriteMessage($"  Color: {viewState.Color.ColorValue}" & vbLf)
            'ed.WriteMessage($"  Transparency: {viewState.Transparency.ToString()}" & vbLf)
            'ed.WriteMessage($"  LinetypeScale: {viewState.LinetypeScale:F6}" & vbLf)

            If viewState.IsDynamic AndAlso viewState.DynamicBlockProperties.Count > 0 Then
                'ed.WriteMessage($"  Dynamic Properties ({viewState.DynamicBlockProperties.Count}):" & vbLf)
                For Each kvp In viewState.DynamicBlockProperties
                    'ed.WriteMessage($"    - {kvp.Key}: {kvp.Value}" & vbLf)
                Next
            ElseIf viewState.IsDynamic Then
                'ed.WriteMessage("  Dynamic Properties: (none captured - may be handled by ResetBlock)" & vbLf)
            End If

            'ed.WriteMessage("========================" & vbLf)

            Return viewState
        End Function

        Private Sub RestoreBlockViewState(blockRef As BlockReference, viewState As BlockViewState)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor

            'ed.WriteMessage(vbLf & "=== RESTORING VIEWSTATE ===" & vbLf)
            'ed.WriteMessage($"  Block Name: {viewState.BlockName}" & vbLf)
            'ed.WriteMessage($"  Is Dynamic: {viewState.IsDynamic}" & vbLf)

            blockRef.Position = viewState.Position
            'ed.WriteMessage($"  Position restored to: X={viewState.Position.X:F4}, Y={viewState.Position.Y:F4}, Z={viewState.Position.Z:F4}" & vbLf)

            blockRef.Rotation = viewState.Rotation
            'ed.WriteMessage($"  Rotation restored to: {viewState.Rotation:F6} radians ({viewState.Rotation * 180 / Math.PI:F2} degrees)" & vbLf)

            blockRef.ScaleFactors = New Scale3d(viewState.ScaleX, viewState.ScaleY, viewState.ScaleZ)
            'ed.WriteMessage($"  Scale restored to: X={viewState.ScaleX:F6}, Y={viewState.ScaleY:F6}, Z={viewState.ScaleZ:F6}" & vbLf)

            blockRef.Layer = viewState.Layer
            'ed.WriteMessage($"  Layer restored to: {viewState.Layer}" & vbLf)

            blockRef.Linetype = viewState.Linetype
            'ed.WriteMessage($"  Linetype restored to: {viewState.Linetype}" & vbLf)

            blockRef.LineWeight = viewState.LineWeight
            'ed.WriteMessage($"  LineWeight restored to: {viewState.LineWeight}" & vbLf)

            blockRef.Color = viewState.Color
            'ed.WriteMessage($"  Color restored to: {viewState.Color.ColorValue}" & vbLf)

            blockRef.Transparency = viewState.Transparency
            'ed.WriteMessage($"  Transparency restored to: {viewState.Transparency.ToString()}" & vbLf)

            blockRef.LinetypeScale = viewState.LinetypeScale
            'ed.WriteMessage($"  LinetypeScale restored to: {viewState.LinetypeScale:F6}" & vbLf)

            ' Restore dynamic block properties if any were captured
            If viewState.IsDynamic AndAlso viewState.DynamicBlockProperties.Count > 0 Then
                Try
                    Dim propCollection = blockRef.DynamicBlockReferencePropertyCollection

                    If propCollection IsNot Nothing AndAlso propCollection.Count > 0 Then
                        'ed.WriteMessage($"  Restoring {viewState.DynamicBlockProperties.Count} dynamic properties:" & vbLf)
                        For Each prop As DynamicBlockReferenceProperty In propCollection
                            Try
                                Dim propName = prop.PropertyName
                                If viewState.DynamicBlockProperties.ContainsKey(propName) Then
                                    Dim originalValue = viewState.DynamicBlockProperties(propName)
                                    prop.Value = originalValue
                                    'ed.WriteMessage($"    - {propName} restored to: {originalValue}" & vbLf)
                                End If
                            Catch ex As Exception
                                ed.WriteMessage($"    - Warning: Could not restore {prop.PropertyName}: {ex.Message}" & vbLf)
                            End Try
                        Next
                    End If
                Catch ex As Exception
                    ed.WriteMessage($"  Warning: Error restoring dynamic block properties: {ex.Message}" & vbLf)
                End Try
            End If

            'ed.WriteMessage("========================" & vbLf)
        End Sub


        <CommandMethod("RSM")>
        Public Sub ResetAllBlocksInModelspace()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Get the modelspace block table record
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    Dim blockCount As Integer = 0

                    ' Iterate through all entities in modelspace
                    For Each objId As ObjectId In btr
                        Dim obj As DBObject = trans.GetObject(objId, OpenMode.ForRead)

                        ' Check if the entity is a block reference
                        If TypeOf obj Is BlockReference Then
                            Dim blockRef As BlockReference = CType(obj, BlockReference)

                            ' Capture the current viewstate
                            Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                            ' Upgrade to write mode
                            blockRef.UpgradeOpen()

                            ' Reset the block reference
                            blockRef.ResetBlock()

                            ' Restore the viewstate
                            RestoreBlockViewState(blockRef, viewState)

                            blockCount += 1
                        End If
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"Reset {blockCount} blocks in modelspace with viewstate preserved." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        <CommandMethod("RSMS")>
        Public Sub ResetSelectedBlocksWithViewState()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Prompt user to select blocks
            Dim psr As PromptSelectionResult = ed.GetSelection()
            If psr.Status <> PromptStatus.OK Then
                ed.WriteMessage(vbLf & "No blocks selected." & vbLf)
                Return
            End If

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim blockCount As Integer = 0

                    ' Process each selected object
                    For Each selObj As SelectedObject In psr.Value
                        Try
                            Dim obj As DBObject = trans.GetObject(selObj.ObjectId, OpenMode.ForRead)

                            ' Check if the entity is a block reference
                            If TypeOf obj Is BlockReference Then
                                Dim blockRef As BlockReference = CType(obj, BlockReference)

                                ' Capture the current viewstate
                                Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                                ' Upgrade to write mode
                                blockRef.UpgradeOpen()

                                ' Reset the block reference
                                blockRef.ResetBlock()

                                ' Restore the viewstate
                                RestoreBlockViewState(blockRef, viewState)

                                blockCount += 1
                            End If
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & $"Warning: Could not reset object: {ex.Message}" & vbLf)
                        End Try
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"Reset {blockCount} selected blocks with viewstate preserved." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        <CommandMethod("MBE")>
        Public Sub MakeAllBlocksInModelspaceExplodable()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Get the modelspace block table record
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    Dim blockCount As Integer = 0
                    Dim processedBlockDefs As New HashSet(Of String)() ' Track unique block definitions

                    ' Iterate through all entities in modelspace
                    For Each objId As ObjectId In btr
                        Dim obj As DBObject = trans.GetObject(objId, OpenMode.ForRead)

                        ' Check if the entity is a block reference
                        If TypeOf obj Is BlockReference Then
                            Dim blockRef As BlockReference = CType(obj, BlockReference)

                            ' Get the block definition
                            Dim blockDef As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                            ' Only process each unique block definition once
                            If Not processedBlockDefs.Contains(blockDef.Name) Then
                                ' Set the block as explodable
                                blockDef.Explodable = True
                                processedBlockDefs.Add(blockDef.Name)
                                blockCount += 1

                                ed.WriteMessage($"  - Block '{blockDef.Name}' is now explodable" & vbLf)
                            End If
                        End If
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"Made {blockCount} unique block definitions explodable." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        <CommandMethod("MBES")>
        Public Sub MakeSelectedBlocksExplodable()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Prompt user to select blocks
            Dim psr As PromptSelectionResult = ed.GetSelection()
            If psr.Status <> PromptStatus.OK Then
                ed.WriteMessage(vbLf & "No blocks selected." & vbLf)
                Return
            End If

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim blockCount As Integer = 0
                    Dim processedBlockDefs As New HashSet(Of String)() ' Track unique block definitions

                    ' Process each selected object
                    For Each selObj As SelectedObject In psr.Value
                        Try
                            Dim obj As DBObject = trans.GetObject(selObj.ObjectId, OpenMode.ForRead)

                            ' Check if the entity is a block reference
                            If TypeOf obj Is BlockReference Then
                                Dim blockRef As BlockReference = CType(obj, BlockReference)

                                ' Get the block definition
                                Dim blockDef As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                                ' Only process each unique block definition once
                                If Not processedBlockDefs.Contains(blockDef.Name) Then
                                    ' Set the block as explodable
                                    blockDef.Explodable = True
                                    processedBlockDefs.Add(blockDef.Name)
                                    blockCount += 1

                                    ed.WriteMessage($"  - Block '{blockDef.Name}' is now explodable" & vbLf)
                                End If
                            End If
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & $"Warning: Could not process object: {ex.Message}" & vbLf)
                        End Try
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"Made {blockCount} unique selected block definitions explodable." & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        <CommandMethod("RSS")>
        Public Sub ResetAndMakeSelectedBlocksExplodable()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Prompt user to select blocks
            Dim psr As PromptSelectionResult = ed.GetSelection()
            If psr.Status <> PromptStatus.OK Then
                ed.WriteMessage(vbLf & "No blocks selected." & vbLf)
                Return
            End If

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim blockCount As Integer = 0
                    Dim explodableCount As Integer = 0
                    Dim colorCount As Integer = 0
                    Dim hatchDeleteCount As Integer = 0
                    Dim processedBlockDefs As New HashSet(Of String)() ' Track unique block definitions
                    Dim blockIds As New List(Of ObjectId)()

                    ' PART 1-3: Process each selected block reference
                    For Each selObj As SelectedObject In psr.Value
                        Try
                            Dim obj As DBObject = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite)

                            ' Check if the entity is a block reference
                            If TypeOf obj Is BlockReference Then
                                Dim blockRef As BlockReference = CType(obj, BlockReference)
                                blockIds.Add(selObj.ObjectId)

                                ' PART 1: Reset the block with viewstate preservation
                                ' Capture the current viewstate
                                Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                                ' Upgrade to write mode
                                blockRef.UpgradeOpen()

                                ' Reset the block reference
                                blockRef.ResetBlock()

                                ' Restore the viewstate
                                RestoreBlockViewState(blockRef, viewState)

                                blockCount += 1

                                ' PART 2: Make the block definition explodable
                                ' Get the block definition
                                Dim blockDef As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                                ' Only process each unique block definition once
                                If Not processedBlockDefs.Contains(blockDef.Name) Then
                                    ' Set the block as explodable
                                    blockDef.Explodable = True
                                    processedBlockDefs.Add(blockDef.Name)
                                    explodableCount += 1
                                End If

                                ' PART 3: Set color to 253 for the block reference
                                blockRef.ColorIndex = 253
                                colorCount += 1
                            End If
                        Catch ex As Exception
                            ed.WriteMessage(vbLf & $"Warning: Could not process object: {ex.Message}" & vbLf)
                        End Try
                    Next

                    ' PART 4: Explode blocks and process all entities
                    For Each blockId As ObjectId In blockIds
                        Try
                            Dim blockRef As BlockReference = CType(trans.GetObject(blockId, OpenMode.ForRead), BlockReference)
                            Dim parentBtr As BlockTableRecord = CType(trans.GetObject(blockRef.OwnerId, OpenMode.ForWrite), BlockTableRecord)
                            Dim explodedEntities As New DBObjectCollection()

                            blockRef.UpgradeOpen()
                            blockRef.Explode(explodedEntities)

                            ' Process exploded entities
                            For Each entity As Entity In explodedEntities
                                Try
                                    ' Set color to 253 for all exploded entities
                                    entity.ColorIndex = 253
                                    colorCount += 1

                                    ' Delete solid hatches
                                    If TypeOf entity Is Hatch Then
                                        Dim hatch As Hatch = CType(entity, Hatch)
                                        If hatch.PatternName = "SOLID" Then
                                            entity.Erase()
                                            hatchDeleteCount += 1
                                            Continue For
                                        End If
                                    End If

                                    ' Add entity to parent space
                                    parentBtr.AppendEntity(entity)
                                    trans.AddNewlyCreatedDBObject(entity, True)

                                Catch ex As Exception
                                    ed.WriteMessage($"  - Warning: Could not process entity: {ex.Message}" & vbLf)
                                End Try
                            Next

                            ' Erase the original block reference
                            blockRef.Erase()

                        Catch ex As Exception
                            ed.WriteMessage($"  - Warning: Could not explode block: {ex.Message}" & vbLf)
                        End Try
                    Next

                    ' PART 5: Process any non-block selected entities (color 253 and delete solid hatches)
                    For Each selObj As SelectedObject In psr.Value
                        Try
                            Dim obj As DBObject = trans.GetObject(selObj.ObjectId, OpenMode.ForRead)

                            ' Process non-block entities
                            If Not (TypeOf obj Is BlockReference) Then
                                Dim entity As Entity = CType(obj, Entity)
                                entity.UpgradeOpen()

                                ' Set color to 253
                                entity.ColorIndex = 253
                                colorCount += 1

                                ' Delete solid hatches
                                If TypeOf entity Is Hatch Then
                                    Dim hatch As Hatch = CType(entity, Hatch)
                                    If hatch.PatternName = "SOLID" Then
                                        entity.Erase()
                                        hatchDeleteCount += 1
                                        Continue For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ' Silently skip problematic entities
                        End Try
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"=== OPERATION SUMMARY ===" & vbLf)
                    ed.WriteMessage($"Reset {blockCount} selected blocks" & vbLf)
                    ed.WriteMessage($"Made {explodableCount} block definitions explodable" & vbLf)
                    ed.WriteMessage($"Set color 253 on {colorCount} entities" & vbLf)
                    ed.WriteMessage($"Deleted {hatchDeleteCount} solid hatches" & vbLf)
                    ed.WriteMessage($"===================" & vbLf)
                Catch ex As Exception
                    trans.Abort()
                    ed.WriteMessage(vbLf & "Error: Operation cancelled. " & ex.Message & vbLf)
                End Try
            End Using
        End Sub

        <CommandMethod("RSMBM")>
        Public Sub ResetAndMakeAllBlocksExplodable()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Get the modelspace block table record
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    Dim blockCount As Integer = 0
                    Dim explodableCount As Integer = 0
                    Dim colorCount As Integer = 0
                    Dim hatchDeleteCount As Integer = 0
                    Dim processedBlockDefs As New HashSet(Of String)() ' Track unique block definitions

                    ' Collect all entity IDs first to avoid collection modification during iteration
                    Dim entityIds As New List(Of ObjectId)()
                    Dim blockIds As New List(Of ObjectId)()

                    For Each objId As ObjectId In btr
                        Dim obj As DBObject = trans.GetObject(objId, OpenMode.ForRead)
                        entityIds.Add(objId)
                        If TypeOf obj Is BlockReference Then
                            blockIds.Add(objId)
                        End If
                    Next

                    ' PART 1-3: Process each block reference
                    For Each blockId As ObjectId In blockIds
                        Try
                            Dim blockRef As BlockReference = CType(trans.GetObject(blockId, OpenMode.ForRead), BlockReference)

                            ' PART 1: Reset the block with viewstate preservation
                            ' Capture the current viewstate
                            Dim viewState As BlockViewState = CaptureBlockViewState(blockRef)

                            ' Upgrade to write mode
                            blockRef.UpgradeOpen()

                            ' Reset the block reference
                            blockRef.ResetBlock()

                            ' Restore the viewstate
                            RestoreBlockViewState(blockRef, viewState)

                            blockCount += 1

                            ' PART 2: Make the block definition explodable
                            ' Get the block definition
                            Dim blockDef As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                            ' Only process each unique block definition once
                            If Not processedBlockDefs.Contains(blockDef.Name) Then
                                ' Set the block as explodable
                                blockDef.Explodable = True
                                processedBlockDefs.Add(blockDef.Name)
                                explodableCount += 1
                            End If

                            ' PART 3: Set color to 253 for the block reference
                            blockRef.ColorIndex = 253
                            colorCount += 1
                        Catch ex As Exception
                            ed.WriteMessage($"Warning: Could not process block: {ex.Message}" & vbLf)
                        End Try
                    Next

                    ' PART 4: Explode blocks and process exploded entities
                    For Each blockId As ObjectId In blockIds
                        Try
                            Dim blockRef As BlockReference = CType(trans.GetObject(blockId, OpenMode.ForRead), BlockReference)

                            ' Skip if already erased
                            If blockRef.IsErased Then
                                Continue For
                            End If

                            Dim parentBtr As BlockTableRecord = CType(trans.GetObject(blockRef.OwnerId, OpenMode.ForWrite), BlockTableRecord)
                            Dim explodedEntities As New DBObjectCollection()

                            blockRef.UpgradeOpen()
                            blockRef.Explode(explodedEntities)

                            ' Process exploded entities
                            For Each entity As Entity In explodedEntities
                                Try
                                    ' Set color to 253 for all exploded entities
                                    entity.ColorIndex = 253
                                    colorCount += 1

                                    ' Delete solid hatches
                                    If TypeOf entity Is Hatch Then
                                        Dim hatch As Hatch = CType(entity, Hatch)
                                        If hatch.PatternName = "SOLID" Then
                                            entity.Erase()
                                            hatchDeleteCount += 1
                                            ed.WriteMessage($"  - Solid hatch deleted" & vbLf)
                                            Continue For
                                        End If
                                    End If

                                    ' Add entity to parent space
                                    parentBtr.AppendEntity(entity)
                                    trans.AddNewlyCreatedDBObject(entity, True)

                                Catch ex As Exception
                                    ed.WriteMessage($"Warning: Could not process entity: {ex.Message}" & vbLf)
                                End Try
                            Next

                            ' Erase the original block reference
                            blockRef.Erase()

                        Catch ex As Exception
                            ed.WriteMessage($"Warning: Could not explode block: {ex.Message}" & vbLf)
                        End Try
                    Next

                    ' PART 5: Process any non-block entities in modelspace (color 253 and delete solid hatches)
                    For Each entityId As ObjectId In entityIds
                        Try
                            Dim obj As DBObject = trans.GetObject(entityId, OpenMode.ForRead)

                            ' Process non-block entities
                            If Not (TypeOf obj Is BlockReference) Then
                                Dim entity As Entity = CType(obj, Entity)
                                entity.UpgradeOpen()

                                ' Set color to 253
                                entity.ColorIndex = 253
                                colorCount += 1

                                ' Delete solid hatches
                                If TypeOf entity Is Hatch Then
                                    Dim hatch As Hatch = CType(entity, Hatch)
                                    If hatch.PatternName = "SOLID" Then
                                        entity.Erase()
                                        hatchDeleteCount += 1
                                        Continue For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ed.WriteMessage($"Warning: Could not process entity: {ex.Message}" & vbLf)
                        End Try
                    Next

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"=== OPERATION SUMMARY ===" & vbLf)
                    ed.WriteMessage($"Reset {blockCount} blocks in modelspace" & vbLf)
                    ed.WriteMessage($"Made {explodableCount} block definitions explodable" & vbLf)
                    ed.WriteMessage($"Set color 253 on {colorCount} entities" & vbLf)
                    ed.WriteMessage($"Deleted {hatchDeleteCount} solid hatches" & vbLf)
                    ed.WriteMessage($"===================" & vbLf)
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message & vbLf)
                End Try
            End Using
        End Sub


        <CommandMethod("EAL")>
        Public Sub ExportAllPaperspaceLayouts()
            'Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acLckDoc As DocumentLock = acDoc.LockDocument()

                'Get the layout dictionary of the current database
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim bgPrev = Application.GetSystemVariable("BackGroundPlot")
                    Dim cmdPrev = Application.GetSystemVariable("CMDDIA")
                    Dim fileDiaPrev = Application.GetSystemVariable("FILEDIA")
                    Dim plotTransPrev = Application.GetSystemVariable("PLOTTRANSPARENCYOVERRIDE")
                    Application.SetSystemVariable("BackGroundPlot", 0)
                    Application.SetSystemVariable("CMDDIA", 0)
                    Application.SetSystemVariable("FILEDIA", 0)
                    'acDoc.SendStringToExecute("-updatefields all 0 ", True, False, False)
                    Try
                        ' 1) Ensure output folder exists
                        Dim outputDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\"
                        Dim DWGname As String = DirectCast(Application.GetSystemVariable("DWGNAME"), String)
                        DWGname = DWGname.Remove(DWGname.Length - 4)
                        Dim pdfFile As String = outputDir & DWGname & ".pdf"
                        Dim dsdFile As String = outputDir & DWGname & ".dsd"

                        Dim dwgprefix As String = Application.GetSystemVariable("dwgprefix")
                        Dim DWGnm As String = Application.GetSystemVariable("dwgName")
                        Dim dwgFile As String = dwgprefix & DWGnm

                        If File.Exists(dsdFile) Then
                            File.Delete(dsdFile)
                        End If

                        If File.Exists(pdfFile) Then
                            File.Delete(pdfFile)
                        End If

                        ' 4) Build DSD entries
                        Dim dsd As New DsdData()
                        Dim dsdEntries As New DsdEntryCollection()

                        Dim dictLayouts As DBDictionary = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead)

                        Dim alllayouts As New List(Of String)
                        For Each entry As DBDictionaryEntry In dictLayouts
                            Dim loId As ObjectId = entry.Value
                            Dim lo As Layout = TryCast(acTrans.GetObject(loId, OpenMode.ForRead), Layout)
                            If lo IsNot Nothing AndAlso Not lo.ModelType Then
                                alllayouts.Add(lo.LayoutName)
                            End If
                        Next
                        For Each lay As Object In alllayouts
                            Dim title As String = DWGnm.Remove(DWGnm.Length - 4) & "-" & lay

                            Dim de As New DsdEntry()
                            de.DwgName = dwgFile
                            de.Layout = lay           ' layout name
                            de.Title = title
                            'de.Nps = "Arcxis"            ' named page setup (ensure it exists)
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
                "IncludeLayer=TRUE",
                "Type=6",                             ' 6 = PDF in many DSDs
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
                            pc = PlotConfigManager.SetCurrentConfig("AutoCAD PDF (High Quality Print).pc3")
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

                    ' Save the changes made
                    acTrans.Commit()

                End Using

            End Using

        End Sub


    End Class
End Namespace