Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Teigha.Colors
Imports Teigha.DatabaseServices
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Exception = System.Exception

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.StandardsConversionCommands))>

Namespace Arcxis_Cad_Tools
    Public Class StandardsConversionCommands
        Private Const TableMax As Integer = 6

        <CommandMethod("STC", CommandFlags.Modal)>
        Public Sub StdConvertHyphen()
            RunStdConvert()
        End Sub

        <CommandMethod("STDCONVERT", CommandFlags.Modal)>
        Public Sub StdConvertAlias()
            RunStdConvert()
        End Sub

        Private Sub RunStdConvert()
            Dim doc = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return

            Dim db = doc.Database
            Dim ed = doc.Editor

            Dim rules As StandardsConversionRules
            Try
                rules = StandardsConversionRules.LoadDefault()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "[STD-CONVERT] " & ex.Message)
                Return
            End Try

            ed.WriteMessage(vbLf & "=== Standards Conversion Tool ===")
            ed.WriteMessage(vbLf & "[STD-CONVERT] Rules: " & rules.RulesPath)
            ed.WriteMessage(vbLf & "[STD-CONVERT] Layer definitions found: " & rules.LayerDefinitions.Count.ToString(CultureInfo.InvariantCulture))

            Dim runStdFull = PromptRunStdFull(ed)
            ed.WriteMessage(vbLf & "Entity mode: LINE, CIRCLE, ARC, TEXT, MTEXT, POLYLINE, MLINE, DIMENSION, LEADER, MLEADER, INSERT, and HATCH entities are converted.")
            ed.WriteMessage(vbLf & "Pipeline: Table 1 through Table 6, in order.")

            Dim selection = ed.GetSelection()
            If selection.Status <> PromptStatus.OK OrElse selection.Value Is Nothing OrElse selection.Value.Count = 0 Then
                ed.WriteMessage(vbLf & "No entities selected. Command cancelled.")
                Return
            End If

            Dim result As StdConversionSummary
            Dim layerResult As StdLayerUpdateSummary

            Using docLock = doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    layerResult = ApplyLayerDefinitions(rules.LayerDefinitions, tr, db)
                    result = ProcessSelection(rules, selection.Value.GetObjectIds(), tr, db)
                    tr.Commit()
                End Using
            End Using

            ed.WriteMessage(vbLf & "Conversion complete!")
            ed.WriteMessage(vbLf & "  Processed (unique entities): " & result.ProcessedEntities.ToString(CultureInfo.InvariantCulture))
            ed.WriteMessage(vbLf & "  Conversion applications: " & result.ConversionApplications.ToString(CultureInfo.InvariantCulture))
            ed.WriteMessage(vbLf & "  Deleted: " & result.DeletedEntities.ToString(CultureInfo.InvariantCulture))
            ed.WriteMessage(vbLf & "  Skipped/unmatched: " & result.SkippedEntities.ToString(CultureInfo.InvariantCulture))
            If layerResult IsNot Nothing AndAlso layerResult.Definitions > 0 Then
                ed.WriteMessage(vbLf & "  Layers preflighted: " & layerResult.Definitions.ToString(CultureInfo.InvariantCulture) & " definitions, " & layerResult.Created.ToString(CultureInfo.InvariantCulture) & " created, " & layerResult.Updated.ToString(CultureInfo.InvariantCulture) & " existing updated")
            End If
            If result.MissingLinetypes.Count > 0 Then
                ed.WriteMessage(vbLf & "  Missing linetypes skipped: " & String.Join(", ", result.MissingLinetypes))
            End If
            If layerResult IsNot Nothing AndAlso layerResult.MissingLinetypes.Count > 0 Then
                ed.WriteMessage(vbLf & "  Missing layer linetypes skipped: " & String.Join(", ", layerResult.MissingLinetypes))
            End If

            If runStdFull Then
                LaunchStdFull(ed)
            Else
                ed.WriteMessage(vbLf & "[STD-CONVERT] STD-FULL auto-run skipped.")
            End If
        End Sub

        Private Shared Function PromptRunStdFull(ed As Editor) As Boolean
            Dim options As New PromptKeywordOptions(vbLf & "Run STD-FULL automatically after conversion? [Yes/No] <Yes>: ")
            options.Keywords.Add("Yes")
            options.Keywords.Add("No")
            options.Keywords.Default = "Yes"
            options.AllowNone = True

            Dim result = ed.GetKeywords(options)
            If result.Status <> PromptStatus.OK OrElse String.IsNullOrWhiteSpace(result.StringResult) Then Return True
            Return result.StringResult.Equals("Yes", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function ProcessSelection(rules As StandardsConversionRules, ids As ObjectId(), tr As Transaction, db As Database) As StdConversionSummary
            Dim summary As New StdConversionSummary()
            Dim plan = BuildSelectionPlan(ids, tr)
            Dim propertyCache As New StdEntityPropertyCache(tr)
            Dim processedHandles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim processedTablesByHandle As New Dictionary(Of String, HashSet(Of Integer))(StringComparer.OrdinalIgnoreCase)
            Dim lastRuleByEntityTable As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For tableId = 1 To TableMax
                Dim cfg = rules.GetTableConfig(tableId)
                Dim tableRules = rules.GetRulesForTable(tableId)
                If tableRules.Count = 0 Then Continue For

                Dim maxRounds = If(cfg.StairStepping, Math.Max(tableRules.Count, 1), 1)
                Dim round = 0
                Dim tableChanged = True

                While tableChanged AndAlso round < maxRounds
                    round += 1
                    tableChanged = False

                    For Each phase In {StdProcessingPhase.Geometry, StdProcessingPhase.Annotation, StdProcessingPhase.Block}
                        If phase = StdProcessingPhase.Block AndAlso Not cfg.BlockProcessing Then Continue For

                        Dim phaseIds = plan.GetIds(phase)
                        If phaseIds.Count = 0 Then Continue For

                        Dim proximityIndex = BuildProximityIndexIfNeeded(tableRules, plan.AllIds, propertyCache, tr, db)

                        For Each id In phaseIds
                            If id.IsNull OrElse id.IsErased Then Continue For

                            Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead, False), Entity)
                            If ent Is Nothing OrElse ent.IsErased Then Continue For

                            Dim handle = ent.Handle.ToString()
                            If IsExcludedForTable(handle, cfg, processedTablesByHandle) Then Continue For
                            If Not cfg.StairStepping AndAlso WasProcessedInTable(handle, tableId, processedTablesByHandle) Then Continue For

                            Dim lastRuleKey = tableId.ToString(CultureInfo.InvariantCulture) & ":" & handle
                            Dim lastRule = If(lastRuleByEntityTable.ContainsKey(lastRuleKey), lastRuleByEntityTable(lastRuleKey), 0)
                            Dim props = propertyCache.GetProperties(id)
                            If props Is Nothing Then Continue For

                            Dim match = rules.FindBestRule(props, tableId, proximityIndex, lastRule)
                            If match Is Nothing Then Continue For

                            ent.UpgradeOpen()
                            Dim applied = ApplyRule(ent, match.Rule, rules.Settings, tr, db, summary)
                            If applied Then
                                propertyCache.Invalidate(id)
                                summary.ConversionApplications += 1
                                If processedHandles.Add(handle) Then summary.ProcessedEntities += 1
                                MarkProcessed(handle, tableId, processedTablesByHandle)
                                lastRuleByEntityTable(lastRuleKey) = match.Rule.Index
                                tableChanged = True
                            End If
                        Next
                    Next
                End While
            Next

            For Each id In plan.AllIds
                If id.IsNull OrElse id.IsErased Then Continue For

                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead, False), Entity)
                If ent Is Nothing Then Continue For

                If Not processedHandles.Contains(ent.Handle.ToString()) AndAlso IsSupported(ent) Then
                    summary.SkippedEntities += 1
                End If
            Next

            Return summary
        End Function

        Private Shared Function ApplyLayerDefinitions(definitions As List(Of StdLayerDefinition), tr As Transaction, db As Database) As StdLayerUpdateSummary
            Dim summary As New StdLayerUpdateSummary()
            If definitions Is Nothing OrElse definitions.Count = 0 Then Return summary

            Dim layerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

            For Each definition In definitions
                If definition Is Nothing OrElse String.IsNullOrWhiteSpace(definition.Name) Then Continue For

                summary.Definitions += 1

                Dim layer As LayerTableRecord
                If layerTable.Has(definition.Name) Then
                    layer = CType(tr.GetObject(layerTable(definition.Name), OpenMode.ForWrite), LayerTableRecord)
                    summary.Updated += 1
                Else
                    If Not layerTable.IsWriteEnabled Then layerTable.UpgradeOpen()
                    layer = New LayerTableRecord With {.Name = definition.Name}
                    layerTable.Add(layer)
                    tr.AddNewlyCreatedDBObject(layer, True)
                    summary.Created += 1
                End If

                ApplyLayerDefinition(layer, definition, tr, db, summary)
            Next

            Return summary
        End Function

        Private Shared Sub ApplyLayerDefinition(layer As LayerTableRecord, definition As StdLayerDefinition, tr As Transaction, db As Database, summary As StdLayerUpdateSummary)
            If definition.Description IsNot Nothing Then
                layer.Description = definition.Description
            End If

            If Not String.IsNullOrWhiteSpace(definition.ColorToken) Then
                ApplyLayerColor(layer, definition.ColorToken)
            End If

            If Not String.IsNullOrWhiteSpace(definition.LinetypeName) AndAlso
               Not definition.LinetypeName.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase) AndAlso
               Not definition.LinetypeName.Equals("BYBLOCK", StringComparison.OrdinalIgnoreCase) Then
                Dim linetypeId = ResolveLinetypeId(db, tr, definition.LinetypeName)
                If linetypeId.IsNull Then
                    summary.MissingLinetypes.Add(definition.LinetypeName)
                Else
                    layer.LinetypeObjectId = linetypeId
                End If
            End If

            Dim lineweight = ParseLineweight(definition.LineweightToken)
            If lineweight.HasValue Then layer.LineWeight = lineweight.Value

            If Not String.IsNullOrWhiteSpace(definition.PlottableToken) Then
                layer.IsPlottable = ParsePlotToken(definition.PlottableToken)
            End If

            If Not String.IsNullOrWhiteSpace(definition.OffToken) Then
                layer.IsOff = StandardsConversionRules.ToBool(definition.OffToken)
            End If

            If Not String.IsNullOrWhiteSpace(definition.FrozenToken) Then
                layer.IsFrozen = StandardsConversionRules.ToBool(definition.FrozenToken)
            End If

            If Not String.IsNullOrWhiteSpace(definition.LockedToken) Then
                layer.IsLocked = StandardsConversionRules.ToBool(definition.LockedToken)
            End If
        End Sub

        Private Shared Function BuildSelectionPlan(ids As ObjectId(), tr As Transaction) As StdSelectionPlan
            Dim plan As New StdSelectionPlan()

            For Each id In ids
                If id.IsNull OrElse id.IsErased Then Continue For

                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead, False), Entity)
                If ent Is Nothing OrElse ent.IsErased OrElse Not IsSupported(ent) Then Continue For

                plan.AllIds.Add(id)
                If BelongsToPhase(ent, StdProcessingPhase.Geometry) Then plan.GeometryIds.Add(id)
                If BelongsToPhase(ent, StdProcessingPhase.Annotation) Then plan.AnnotationIds.Add(id)
                If BelongsToPhase(ent, StdProcessingPhase.Block) Then plan.BlockIds.Add(id)
            Next

            Return plan
        End Function

        ''' <summary>
        ''' Build a proximity index from the current selection plus any supported entities in the current
        ''' drawing space whose bounds overlap an axis-aligned box: union of selected entity bounds,
        ''' expanded by half of the largest rule search window (feet). This approximates Lisp STD-CONVERT
        ''' ssget crossing-window behavior so linework does not have to be explicitly selected.
        ''' </summary>
        Private Shared Function BuildProximityIndexIfNeeded(tableRules As List(Of StdConversionRule), ids As List(Of ObjectId), propertyCache As StdEntityPropertyCache, tr As Transaction, db As Database) As StdProximityIndex
            Dim maxSearchWindow = 0.0
            For Each rule In tableRules
                If rule.FromPart.SearchWindow.HasValue AndAlso rule.FromPart.SearchWindow.Value > maxSearchWindow Then
                    maxSearchWindow = rule.FromPart.SearchWindow.Value
                End If
            Next

            If maxSearchWindow <= 0.0 Then Return Nothing

            Dim halfWin = maxSearchWindow / 2.0

            Dim selMinX, selMinY, selMaxX, selMaxY As Double
            If Not TryComputeSelectionPlanarBounds(ids, tr, propertyCache, selMinX, selMinY, selMaxX, selMaxY) Then
                Return BuildProximityIndexFromIds(ids, propertyCache, maxSearchWindow)
            End If

            Dim qMinX = selMinX - halfWin
            Dim qMinY = selMinY - halfWin
            Dim qMaxX = selMaxX + halfWin
            Dim qMaxY = selMaxY + halfWin

            Dim mergedIds As New HashSet(Of ObjectId)()
            For Each id In ids
                If Not id.IsNull AndAlso Not id.IsErased Then mergedIds.Add(id)
            Next

            Try
                Dim spaceBtr = TryCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead), BlockTableRecord)
                If spaceBtr IsNot Nothing Then
                    For Each oid In spaceBtr
                        If oid.IsNull OrElse oid.IsErased OrElse mergedIds.Contains(oid) Then Continue For

                        Dim ent = TryCast(tr.GetObject(oid, OpenMode.ForRead, False), Entity)
                        If ent Is Nothing OrElse ent.IsErased OrElse Not IsSupported(ent) Then Continue For

                        Dim eminX, eminY, emaxX, emaxY As Double
                        If Not TryGetPlanarBounds(ent, eminX, eminY, emaxX, emaxY) Then
                            Dim p = propertyCache.GetProperties(oid)
                            If p Is Nothing OrElse Not p.Origin.HasValue Then Continue For
                            Dim o = p.Origin.Value
                            eminX = o.X
                            emaxX = o.X
                            eminY = o.Y
                            emaxY = o.Y
                        End If

                        If BoxesOverlap2d(eminX, eminY, emaxX, emaxY, qMinX, qMinY, qMaxX, qMaxY) Then
                            mergedIds.Add(oid)
                        End If
                    Next
                End If
            Catch
            End Try

            Return BuildProximityIndexFromIds(mergedIds, propertyCache, maxSearchWindow)
        End Function

        Private Shared Function BuildProximityIndexFromIds(ids As IEnumerable(Of ObjectId), propertyCache As StdEntityPropertyCache, maxSearchWindow As Double) As StdProximityIndex
            Dim properties As New List(Of StdEntityProperties)()
            For Each id In ids
                Dim props = propertyCache.GetProperties(id)
                If props IsNot Nothing Then properties.Add(props)
            Next

            Return StdProximityIndex.Build(properties, maxSearchWindow)
        End Function

        Private Shared Function TryGetPlanarBounds(ent As Entity, ByRef minX As Double, ByRef minY As Double, ByRef maxX As Double, ByRef maxY As Double) As Boolean
            Try
                Dim ext = ent.GeometricExtents
                Dim x0 = ext.MinPoint.X
                Dim x1 = ext.MaxPoint.X
                Dim y0 = ext.MinPoint.Y
                Dim y1 = ext.MaxPoint.Y
                minX = Math.Min(x0, x1)
                maxX = Math.Max(x0, x1)
                minY = Math.Min(y0, y1)
                maxY = Math.Max(y0, y1)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function BoxesOverlap2d(aminX As Double, aminY As Double, amaxX As Double, amaxY As Double, bminX As Double, bminY As Double, bmaxX As Double, bmaxY As Double) As Boolean
            Return Not (amaxX < bminX OrElse aminX > bmaxX OrElse amaxY < bminY OrElse aminY > bmaxY)
        End Function

        Private Shared Function TryComputeSelectionPlanarBounds(ids As List(Of ObjectId), tr As Transaction, propertyCache As StdEntityPropertyCache, ByRef minX As Double, ByRef minY As Double, ByRef maxX As Double, ByRef maxY As Double) As Boolean
            Dim hasAny = False

            For Each id In ids
                If id.IsNull OrElse id.IsErased Then Continue For

                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead, False), Entity)
                If ent Is Nothing OrElse ent.IsErased Then Continue For

                Dim eminX, eminY, emaxX, emaxY As Double
                If TryGetPlanarBounds(ent, eminX, eminY, emaxX, emaxY) Then
                    If Not hasAny Then
                        minX = eminX
                        minY = eminY
                        maxX = emaxX
                        maxY = emaxY
                        hasAny = True
                    Else
                        minX = Math.Min(minX, eminX)
                        minY = Math.Min(minY, eminY)
                        maxX = Math.Max(maxX, emaxX)
                        maxY = Math.Max(maxY, emaxY)
                    End If
                Else
                    Dim p = propertyCache.GetProperties(id)
                    If p Is Nothing OrElse Not p.Origin.HasValue Then Continue For
                    Dim o = p.Origin.Value
                    If Not hasAny Then
                        minX = o.X
                        minY = o.Y
                        maxX = o.X
                        maxY = o.Y
                        hasAny = True
                    Else
                        minX = Math.Min(minX, o.X)
                        minY = Math.Min(minY, o.Y)
                        maxX = Math.Max(maxX, o.X)
                        maxY = Math.Max(maxY, o.Y)
                    End If
                End If
            Next

            Return hasAny
        End Function

        Private Shared Function ApplyRule(ent As Entity, rule As StdConversionRule, settings As StdConversionSettings, tr As Transaction, db As Database, summary As StdConversionSummary) As Boolean
            If rule.FromPart.DeleteEntity OrElse rule.ToPart.DeleteEntity Then
                ent.Erase()
                summary.DeletedEntities += 1
                Return True
            End If

            Dim changed = False

            If Not String.IsNullOrWhiteSpace(rule.ToPart.LayerName) Then
                EnsureLayer(db, tr, rule.ToPart.LayerName, settings.CreateMissingLayers)
                If Not ent.Layer.Equals(rule.ToPart.LayerName, StringComparison.OrdinalIgnoreCase) Then
                    ent.Layer = rule.ToPart.LayerName
                    changed = True
                End If
            End If

            If Not String.IsNullOrWhiteSpace(rule.ToPart.LineType) Then
                If LinetypeExists(db, tr, rule.ToPart.LineType) Then
                    If Not ent.Linetype.Equals(rule.ToPart.LineType, StringComparison.OrdinalIgnoreCase) Then
                        ent.Linetype = rule.ToPart.LineType
                        changed = True
                    End If
                Else
                    summary.MissingLinetypes.Add(rule.ToPart.LineType)
                End If
            End If

            If ApplyTextActions(ent, rule, settings, tr, db) Then changed = True

            If rule.ToPart.ForceAnnotationToDefault OrElse Not String.IsNullOrWhiteSpace(rule.ToPart.AnnotationStyle) Then
                If ApplyAnnotationStyle(ent, rule.ToPart.AnnotationStyle, settings, tr, db) Then changed = True
            End If

            Return changed
        End Function

        Private Shared Function ApplyTextActions(ent As Entity, rule As StdConversionRule, settings As StdConversionSettings, tr As Transaction, db As Database) As Boolean
            Dim changed = False
            Dim dbText = TryCast(ent, DBText)
            If dbText IsNot Nothing Then
                If rule.ToPart.TextHeight.HasValue AndAlso Math.Abs(dbText.Height - rule.ToPart.TextHeight.Value) > settings.TextHeightTolerance Then
                    dbText.Height = rule.ToPart.TextHeight.Value
                    changed = True
                End If

                If Not String.IsNullOrWhiteSpace(rule.ToPart.AnnotationStyle) AndAlso TrySetTextStyle(dbText, rule.ToPart.AnnotationStyle, settings, tr, db) Then
                    changed = True
                End If

                If Not String.IsNullOrWhiteSpace(rule.FromPart.TextContents) AndAlso rule.ToPart.TextContents IsNot Nothing Then
                    Dim replaced = StandardsConversionRules.ReplaceText(dbText.TextString, rule.FromPart.TextContents, rule.ToPart.TextContents)
                    If Not String.Equals(replaced, dbText.TextString, StringComparison.Ordinal) Then
                        dbText.TextString = replaced
                        changed = True
                    End If
                End If

                Return changed
            End If

            Dim mtext = TryCast(ent, MText)
            If mtext IsNot Nothing Then
                If rule.ToPart.TextHeight.HasValue AndAlso Math.Abs(mtext.TextHeight - rule.ToPart.TextHeight.Value) > settings.TextHeightTolerance Then
                    mtext.TextHeight = rule.ToPart.TextHeight.Value
                    changed = True
                End If

                If Not String.IsNullOrWhiteSpace(rule.ToPart.AnnotationStyle) AndAlso TrySetTextStyle(mtext, rule.ToPart.AnnotationStyle, settings, tr, db) Then
                    changed = True
                End If

                If Not String.IsNullOrWhiteSpace(rule.FromPart.TextContents) AndAlso rule.ToPart.TextContents IsNot Nothing Then
                    Dim replaced = StandardsConversionRules.ReplaceText(mtext.Contents, rule.FromPart.TextContents, rule.ToPart.TextContents)
                    If Not String.Equals(replaced, mtext.Contents, StringComparison.Ordinal) Then
                        mtext.Contents = replaced
                        changed = True
                    End If
                End If
            End If

            Return changed
        End Function

        Private Shared Function ApplyAnnotationStyle(ent As Entity, styleName As String, settings As StdConversionSettings, tr As Transaction, db As Database) As Boolean
            If String.IsNullOrWhiteSpace(styleName) Then Return False

            Dim dbText = TryCast(ent, DBText)
            If dbText IsNot Nothing Then Return TrySetTextStyle(dbText, styleName, settings, tr, db)

            Dim mtext = TryCast(ent, MText)
            If mtext IsNot Nothing Then Return TrySetTextStyle(mtext, styleName, settings, tr, db)

            Return TrySetNamedStyleProperty(ent, styleName)
        End Function

        Private Shared Function TrySetTextStyle(text As DBText, styleName As String, settings As StdConversionSettings, tr As Transaction, db As Database) As Boolean
            Dim styleId = EnsureTextStyle(db, tr, styleName, settings.CreateMissingStyles)
            If styleId.IsNull OrElse text.TextStyleId = styleId Then Return False
            text.TextStyleId = styleId
            Return True
        End Function

        Private Shared Function TrySetTextStyle(text As MText, styleName As String, settings As StdConversionSettings, tr As Transaction, db As Database) As Boolean
            Dim styleId = EnsureTextStyle(db, tr, styleName, settings.CreateMissingStyles)
            If styleId.IsNull OrElse text.TextStyleId = styleId Then Return False
            text.TextStyleId = styleId
            Return True
        End Function

        Private Shared Function TrySetNamedStyleProperty(ent As Entity, styleName As String) As Boolean
            For Each propName In {"DimensionStyleName", "MLeaderStyleName", "StyleName"}
                Try
                    Dim prop = ent.GetType().GetProperty(propName)
                    If prop IsNot Nothing AndAlso prop.CanWrite AndAlso prop.PropertyType Is GetType(String) Then
                        Dim current = TryCast(prop.GetValue(ent), String)
                        If Not String.Equals(current, styleName, StringComparison.OrdinalIgnoreCase) Then
                            prop.SetValue(ent, styleName)
                            Return True
                        End If
                    End If
                Catch
                End Try
            Next

            Return False
        End Function

        Private Shared Sub EnsureLayer(db As Database, tr As Transaction, layerName As String, createMissing As Boolean)
            Dim layerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            If Not layerTable.Has(layerName) Then
                If Not createMissing Then Return
                layerTable.UpgradeOpen()
                Dim rec As New LayerTableRecord With {.Name = layerName}
                layerTable.Add(rec)
                tr.AddNewlyCreatedDBObject(rec, True)
            End If

            Dim layer = CType(tr.GetObject(layerTable(layerName), OpenMode.ForWrite), LayerTableRecord)
            layer.IsOff = False
            layer.IsFrozen = False
            layer.IsLocked = False
        End Sub

        Private Shared Function EnsureTextStyle(db As Database, tr As Transaction, styleName As String, createMissing As Boolean) As ObjectId
            Dim table = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
            If table.Has(styleName) Then Return table(styleName)
            If Not createMissing Then Return ObjectId.Null

            table.UpgradeOpen()
            Dim rec As New TextStyleTableRecord With {.Name = styleName}
            Dim id = table.Add(rec)
            tr.AddNewlyCreatedDBObject(rec, True)
            Return id
        End Function

        Private Shared Function LinetypeExists(db As Database, tr As Transaction, lineTypeName As String) As Boolean
            If String.Equals(lineTypeName, "BYLAYER", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(lineTypeName, "BYBLOCK", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim table = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
            Return table.Has(lineTypeName)
        End Function

        Private Shared Sub ApplyLayerColor(layer As LayerTableRecord, token As String)
            If token.StartsWith("RGB:", StringComparison.OrdinalIgnoreCase) Then
                Dim parts = token.Substring(4).Split(","c)
                If parts.Length = 3 Then
                    Dim r = StandardsConversionRules.ToInt(parts(0), -1)
                    Dim g = StandardsConversionRules.ToInt(parts(1), -1)
                    Dim b = StandardsConversionRules.ToInt(parts(2), -1)
                    If r >= 0 AndAlso r <= 255 AndAlso g >= 0 AndAlso g <= 255 AndAlso b >= 0 AndAlso b <= 255 Then
                        layer.Color = Color.FromRgb(CByte(r), CByte(g), CByte(b))
                    End If
                End If
                Return
            End If

            If token.StartsWith("ACI:", StringComparison.OrdinalIgnoreCase) Then
                token = token.Substring(4)
            End If

            Dim aci = StandardsConversionRules.ToInt(token, Integer.MinValue)
            If aci >= 1 AndAlso aci <= 255 Then
                layer.Color = Color.FromColorIndex(ColorMethod.ByAci, CShort(aci))
            End If
        End Sub

        Private Shared Function ParseLineweight(token As String) As LineWeight?
            If String.IsNullOrWhiteSpace(token) Then Return Nothing

            Dim cleaned = token.Trim()
            If cleaned.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase) Then Return LineWeight.ByLayer
            If cleaned.Equals("BYBLOCK", StringComparison.OrdinalIgnoreCase) Then Return LineWeight.ByBlock
            If cleaned.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase) Then Return LineWeight.ByLineWeightDefault

            cleaned = cleaned.Replace("mm", "", StringComparison.OrdinalIgnoreCase).Trim()

            Dim numeric As Double
            If Double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, numeric) Then
                Dim hundredths = If(numeric > 0 AndAlso numeric < 3, CInt(Math.Round(numeric * 100.0)), CInt(Math.Round(numeric)))
                Return CType(hundredths, LineWeight)
            End If

            Return Nothing
        End Function

        Private Shared Function ParsePlotToken(token As String) As Boolean
            Select Case token.Trim().ToUpperInvariant()
                Case "PLOT", "PLOTTABLE", "YES", "Y", "TRUE", "T", "1"
                    Return True
                Case "NO-PLOT", "NOPLOT", "NO PLOT", "NO", "N", "FALSE", "F", "0"
                    Return False
                Case Else
                    Return StandardsConversionRules.ToBool(token)
            End Select
        End Function

        Private Shared Function ResolveLinetypeId(db As Database, tr As Transaction, lineTypeName As String) As ObjectId
            If String.IsNullOrWhiteSpace(lineTypeName) OrElse
               lineTypeName.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase) OrElse
               lineTypeName.Equals("BYBLOCK", StringComparison.OrdinalIgnoreCase) Then
                Return ObjectId.Null
            End If

            Dim table = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
            If table.Has(lineTypeName) Then Return table(lineTypeName)

            For Each fileName In {"acad.lin", "default.lin", "iso.lin"}
                Try
                    db.LoadLineTypeFile(lineTypeName, fileName)
                    If table.Has(lineTypeName) Then Return table(lineTypeName)
                Catch
                End Try
            Next

            Return ObjectId.Null
        End Function

        Private Shared Function IsSupported(ent As Entity) As Boolean
            Return TypeOf ent Is Line OrElse
                   TypeOf ent Is Circle OrElse
                   TypeOf ent Is Arc OrElse
                   TypeOf ent Is DBText OrElse
                   TypeOf ent Is MText OrElse
                   TypeOf ent Is BlockReference OrElse
                   TypeOf ent Is Polyline OrElse
                   TypeOf ent Is Polyline2d OrElse
                   TypeOf ent Is Hatch OrElse
                   ent.GetType().Name.IndexOf("Mline", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   ent.GetType().Name.IndexOf("Dimension", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   ent.GetType().Name.IndexOf("Leader", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function BelongsToPhase(ent As Entity, phase As StdProcessingPhase) As Boolean
            Select Case phase
                Case StdProcessingPhase.Geometry
                    Return TypeOf ent Is Line OrElse TypeOf ent Is Circle OrElse TypeOf ent Is Arc OrElse
                           TypeOf ent Is Polyline OrElse TypeOf ent Is Polyline2d OrElse TypeOf ent Is Hatch OrElse
                           ent.GetType().Name.IndexOf("Mline", StringComparison.OrdinalIgnoreCase) >= 0
                Case StdProcessingPhase.Annotation
                    Return TypeOf ent Is DBText OrElse TypeOf ent Is MText OrElse
                           ent.GetType().Name.IndexOf("Dimension", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           ent.GetType().Name.IndexOf("Leader", StringComparison.OrdinalIgnoreCase) >= 0
                Case StdProcessingPhase.Block
                    Return TypeOf ent Is BlockReference
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsExcludedForTable(handle As String, cfg As StdTableConfig, processedTablesByHandle As Dictionary(Of String, HashSet(Of Integer))) As Boolean
            If Not processedTablesByHandle.ContainsKey(handle) Then Return False
            For Each tableId In cfg.ExcludeFromTables
                If processedTablesByHandle(handle).Contains(tableId) Then Return True
            Next
            Return False
        End Function

        Private Shared Function WasProcessedInTable(handle As String, tableId As Integer, processedTablesByHandle As Dictionary(Of String, HashSet(Of Integer))) As Boolean
            Return processedTablesByHandle.ContainsKey(handle) AndAlso processedTablesByHandle(handle).Contains(tableId)
        End Function

        Private Shared Sub MarkProcessed(handle As String, tableId As Integer, processedTablesByHandle As Dictionary(Of String, HashSet(Of Integer)))
            If Not processedTablesByHandle.ContainsKey(handle) Then
                processedTablesByHandle(handle) = New HashSet(Of Integer)()
            End If
            processedTablesByHandle(handle).Add(tableId)
        End Sub

        Private Shared Sub LaunchStdFull(ed As Editor)
            Try
                ed.WriteMessage(vbLf & "[STD-CONVERT] Launching STD-FULL. If the command is not loaded, load the STD-FULL tool and rerun it manually.")
                Application.DocumentManager.MdiActiveDocument.SendStringToExecute("STD-FULL ", True, False, False)
            Catch ex As Exception
                ed.WriteMessage(vbLf & "[STD-CONVERT] STD-FULL launch failed: " & ex.Message)
            End Try
        End Sub
    End Class

    Friend Enum StdProcessingPhase
        Geometry
        Annotation
        Block
    End Enum

    Friend Class StdSelectionPlan
        Public ReadOnly Property AllIds As New List(Of ObjectId)()
        Public ReadOnly Property GeometryIds As New List(Of ObjectId)()
        Public ReadOnly Property AnnotationIds As New List(Of ObjectId)()
        Public ReadOnly Property BlockIds As New List(Of ObjectId)()

        Public Function GetIds(phase As StdProcessingPhase) As List(Of ObjectId)
            Select Case phase
                Case StdProcessingPhase.Geometry
                    Return GeometryIds
                Case StdProcessingPhase.Annotation
                    Return AnnotationIds
                Case StdProcessingPhase.Block
                    Return BlockIds
                Case Else
                    Return AllIds
            End Select
        End Function
    End Class

    Friend Class StdEntityPropertyCache
        Private ReadOnly _transaction As Transaction
        Private ReadOnly _items As New Dictionary(Of ObjectId, StdEntityProperties)()

        Public Sub New(transaction As Transaction)
            _transaction = transaction
        End Sub

        Public Function GetProperties(id As ObjectId) As StdEntityProperties
            If id.IsNull OrElse id.IsErased Then Return Nothing

            If _items.ContainsKey(id) Then Return _items(id)

            Try
                Dim ent = TryCast(_transaction.GetObject(id, OpenMode.ForRead, False), Entity)
                If ent Is Nothing OrElse ent.IsErased Then Return Nothing

                Dim props = StdEntityProperties.FromEntity(ent, _transaction)
                _items(id) = props
                Return props
            Catch
                Return Nothing
            End Try
        End Function

        Public Sub Invalidate(id As ObjectId)
            If _items.ContainsKey(id) Then _items.Remove(id)
        End Sub
    End Class

    Friend Class StdConversionSummary
        Public Property ProcessedEntities As Integer
        Public Property ConversionApplications As Integer
        Public Property DeletedEntities As Integer
        Public Property SkippedEntities As Integer
        Public ReadOnly Property MissingLinetypes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End Class

    Friend Class StdLayerUpdateSummary
        Public Property Definitions As Integer
        Public Property Created As Integer
        Public Property Updated As Integer
        Public ReadOnly Property MissingLinetypes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End Class
End Namespace
