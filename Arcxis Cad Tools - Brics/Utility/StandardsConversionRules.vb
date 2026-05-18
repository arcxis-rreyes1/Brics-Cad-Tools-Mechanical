Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Microsoft.VisualBasic.FileIO
Imports Teigha.Colors
Imports Teigha.DatabaseServices
Imports Teigha.Geometry

Namespace Arcxis_Cad_Tools
    Friend NotInheritable Class StandardsConversionRules
        Private Const EnvCsvPath As String = "ARCXIS_STD_CONVERT_CSV"
        Private Const EnvRulesPath As String = "ARCXIS_STD_CONVERT_RULES"
        Private Const EnvXlsxPath As String = "ARCXIS_STD_CONVERT_XLSX"

        ''' <summary>Canonical shared STD-CONVERT workbook (Frame-E CAD lisp bundle).</summary>
        Private Const SharedXlsxPath As String = "E:\Shared\Arcxis\Engineering\User Development Files\Brandon Hartman\ARCXIS CAD Customization\Frame-E CAD lisp\(LD)STD-CONVERT\conversion table.xlsx"

        Public ReadOnly Property RulesPath As String
        Public ReadOnly Property CsvPath As String
        Public ReadOnly Property Settings As StdConversionSettings
        Public ReadOnly Property TableConfigs As Dictionary(Of Integer, StdTableConfig)
        Public ReadOnly Property Rules As List(Of StdConversionRule)
        Public ReadOnly Property LayerDefinitions As List(Of StdLayerDefinition)
        Private ReadOnly _rulesByTable As Dictionary(Of Integer, List(Of StdConversionRule))

        Private Sub New(rulesPath As String, settings As StdConversionSettings, tableConfigs As Dictionary(Of Integer, StdTableConfig), rules As List(Of StdConversionRule), layerDefinitions As List(Of StdLayerDefinition))
            Me.RulesPath = rulesPath
            Me.CsvPath = rulesPath
            Me.Settings = settings
            Me.TableConfigs = tableConfigs
            Me.Rules = rules
            Me.LayerDefinitions = layerDefinitions
            _rulesByTable = rules.GroupBy(Function(rule) rule.TableId).
                ToDictionary(Function(group) group.Key, Function(group) group.OrderBy(Function(rule) rule.Index).ToList())
        End Sub

        Public Shared Function LoadDefault() As StandardsConversionRules
            Dim rulesPath = ResolveRulesPath()
            If String.IsNullOrWhiteSpace(rulesPath) Then
                Throw New FileNotFoundException("Could not resolve STD-CONVERT rules. Set ARCXIS_STD_CONVERT_RULES, ARCXIS_STD_CONVERT_XLSX, ARCXIS_STD_CONVERT_CSV, or deploy the packaged workbook/CSV.")
            End If

            Return Load(rulesPath)
        End Function

        Private Shared Function ResolveRulesPath() As String
            Dim rulesPath = Environment.GetEnvironmentVariable(EnvRulesPath)
            If IsReadableFile(rulesPath) Then Return rulesPath

            Dim xlsxPath = Environment.GetEnvironmentVariable(EnvXlsxPath)
            If IsReadableFile(xlsxPath) Then Return xlsxPath

            Dim envPath = Environment.GetEnvironmentVariable(EnvCsvPath)
            If IsReadableFile(envPath) Then Return envPath

            If IsReadableFile(SharedXlsxPath) Then Return SharedXlsxPath

            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
            Dim packagedXlsxPath = Path.Combine(baseDir, "Resources", "STD-CONVERT", "conversion table.xlsx")
            If IsReadableFile(packagedXlsxPath) Then Return packagedXlsxPath

            Dim packagedCsvPath = Path.Combine(baseDir, "Resources", "STD-CONVERT", "conversion table.csv")
            If IsReadableFile(packagedCsvPath) Then Return packagedCsvPath

            Return Nothing
        End Function

        Private Shared Function IsReadableFile(candidate As String) As Boolean
            If String.IsNullOrWhiteSpace(candidate) Then Return False
            Try
                Return File.Exists(candidate)
            Catch
                Return False
            End Try
        End Function

        Private Shared Function Load(rulesPath As String) As StandardsConversionRules
            Dim rows = ReadRows(rulesPath)
            Dim ruleHeaderRow = FindHeaderRow(rows, "Rule", "Table 1")
            Dim settingHeaderRow = FindHeaderRow(rows, "Setting", "Value")
            Dim tableHeaderRow = FindTableSectionHeaderRow(rows)

            If settingHeaderRow < 0 OrElse tableHeaderRow < 0 Then
                Throw New InvalidDataException("Rules file is missing the Setting/Value section and conversion table columns (Side, Layer, …). Export the full workbook or use the wide layout with Rule + Table sections.")
            End If

            Dim settings = BuildSettings(rows, settingHeaderRow, tableHeaderRow)
            Dim tableConfigs As Dictionary(Of Integer, StdTableConfig)
            If ruleHeaderRow >= 0 Then
                tableConfigs = BuildTableConfigs(rows, ruleHeaderRow, settingHeaderRow)
            Else
                tableConfigs = BuildDefaultTableConfigs()
            End If
            Dim rules = BuildRules(rows, tableHeaderRow)
            Dim layerDefinitions = BuildLayerDefinitions(rows)
            Return New StandardsConversionRules(rulesPath, settings, tableConfigs, rules, layerDefinitions)
        End Function

        Private Shared Function ReadRows(rulesPath As String) As List(Of String())
            Dim extension = Path.GetExtension(rulesPath)
            If extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) OrElse
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) Then
                Return ReadWorkbook(rulesPath)
            End If

            Return ReadCsv(rulesPath)
        End Function

        Private Shared Function ReadWorkbook(workbookPath As String) As List(Of String())
            Dim rows As New List(Of String())()

            Using document = SpreadsheetDocument.Open(workbookPath, False)
                Dim workbookPart = document.WorkbookPart
                If workbookPart Is Nothing Then Return rows

                For Each sheet In workbookPart.Workbook.Sheets.Elements(Of Sheet)()
                    If sheet Is Nothing OrElse sheet.Id Is Nothing Then Continue For

                    Dim worksheetPart = TryCast(workbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                    If worksheetPart Is Nothing Then Continue For

                    rows.Add(New String() {If(sheet.Name?.Value, "")})

                    For Each row In worksheetPart.Worksheet.Descendants(Of DocumentFormat.OpenXml.Spreadsheet.Row)()
                        Dim values As New SortedDictionary(Of Integer, String)()
                        For Each cell In row.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Cell)()
                            Dim columnIndex = GetColumnIndex(cell.CellReference)
                            If columnIndex >= 0 Then values(columnIndex) = GetCellText(cell, workbookPart)
                        Next

                        If values.Count = 0 Then
                            rows.Add(Array.Empty(Of String)())
                        Else
                            Dim fields(values.Keys.Max()) As String
                            For Each item In values
                                fields(item.Key) = item.Value
                            Next
                            rows.Add(fields)
                        End If
                    Next
                Next
            End Using

            Return rows
        End Function

        Private Shared Function GetColumnIndex(reference As String) As Integer
            If String.IsNullOrWhiteSpace(reference) Then Return -1

            Dim letters = Regex.Match(reference, "^[A-Za-z]+").Value
            If letters.Length = 0 Then Return -1

            Dim index = 0
            For Each ch In letters.ToUpperInvariant()
                index = (index * 26) + (AscW(ch) - AscW("A"c) + 1)
            Next

            Return index - 1
        End Function

        Private Shared Function GetCellText(cell As DocumentFormat.OpenXml.Spreadsheet.Cell, workbookPart As WorkbookPart) As String
            If cell Is Nothing Then Return ""

            If cell.DataType IsNot Nothing Then
                Select Case cell.DataType.Value
                    Case CellValues.SharedString
                        Dim sharedIndex = ToInt(cell.CellValue?.Text, -1)
                        Dim sst = workbookPart.SharedStringTablePart?.SharedStringTable
                        If sst IsNot Nothing AndAlso sharedIndex >= 0 Then
                            Dim sharedItem = sst.Elements(Of SharedStringItem)().ElementAtOrDefault(sharedIndex)
                            If sharedItem IsNot Nothing Then Return sharedItem.InnerText
                        End If
                        Return ""
                    Case CellValues.Boolean
                        Return If(cell.CellValue?.Text = "1", "TRUE", "FALSE")
                    Case CellValues.InlineString
                        Return If(cell.InlineString?.InnerText, "")
                    Case Else
                        Return If(cell.CellValue?.Text, "")
                End Select
            End If

            Return If(cell.CellValue?.Text, "")
        End Function

        Private Shared Function ReadCsv(csvPath As String) As List(Of String())
            Dim rows As New List(Of String())()
            Using parser As New TextFieldParser(csvPath)
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                While Not parser.EndOfData
                    rows.Add(parser.ReadFields())
                End While
            End Using
            Return rows
        End Function

        Private Shared Function FindHeaderRow(rows As List(Of String()), ParamArray requiredHeaders As String()) As Integer
            For i = 0 To rows.Count - 1
                Dim row = rows(i)
                If requiredHeaders.All(Function(header) FindColumn(row, header) >= 0) Then Return i
            Next
            Return -1
        End Function

        ''' <summary>Wide layout: Table + Side + Layer. Compact Excel layout: Side + Layer + Linetype (no Rule row, single implicit table).</summary>
        Private Shared Function FindTableSectionHeaderRow(rows As List(Of String())) As Integer
            Dim idx = FindHeaderRow(rows, "Table", "Side", "Layer")
            If idx >= 0 Then Return idx
            Return FindHeaderRow(rows, "Side", "Layer", "Linetype")
        End Function

        ''' <summary>Used when the workbook has no Rule/Table 1..6 matrix (compact Arcxis XLSX export).</summary>
        Private Shared Function BuildDefaultTableConfigs() As Dictionary(Of Integer, StdTableConfig)
            Dim configs As New Dictionary(Of Integer, StdTableConfig)()
            For tableId = 1 To 6
                Dim cfg = New StdTableConfig(tableId) With {
                    .BlockProcessing = False,
                    .StrictestMatchFirst = True,
                    .StairStepping = True
                }
                If tableId = 2 Then cfg.ExcludeFromTables.Add(1)
                configs(tableId) = cfg
            Next
            Return configs
        End Function

        Private Shared Function BuildSettings(rows As List(Of String()), settingHeaderRow As Integer, tableHeaderRow As Integer) As StdConversionSettings
            Dim settings As New StdConversionSettings()
            Dim header = rows(settingHeaderRow)
            Dim settingColumn = FindColumn(header, "Setting")
            Dim valueColumn = FindColumn(header, "Value")
            If settingColumn < 0 Then settingColumn = 0
            If valueColumn < 0 Then valueColumn = 1

            For i = settingHeaderRow + 1 To Math.Min(tableHeaderRow - 1, rows.Count - 1)
                Dim key = Clean(GetField(rows(i), settingColumn))
                Dim value = Clean(GetField(rows(i), valueColumn))
                If key.Length = 0 OrElse value.Length = 0 Then Continue For

                Select Case key.ToUpperInvariant()
                    Case "*STD-LOG-LEVEL*"
                        settings.LogLevel = ToInt(value, settings.LogLevel)
                    Case "*STD-DEBUG-LOOKUP-IO*"
                        settings.DebugLookup = ToBool(value)
                    Case "*STD-DEBUG-LOOKUP-LIMIT*"
                        settings.DebugLookupLimit = ToInt(value, settings.DebugLookupLimit)
                    Case "*STD-ANGLE-PRECISION-DEG*"
                        settings.AngleToleranceDegrees = ToDouble(value, settings.AngleToleranceDegrees)
                    Case "*STD-TEXT-HEIGHT-PRECISION*"
                        settings.TextHeightTolerance = ToDouble(value, settings.TextHeightTolerance)
                    Case "*STD-TEXT-ROTATION-PRECISION-DEG*"
                        settings.TextRotationToleranceDegrees = ToDouble(value, settings.TextRotationToleranceDegrees)
                    Case "*STD-MISSING-LAYER-MAKE*"
                        settings.CreateMissingLayers = ToBool(value)
                    Case "*STD-MISSING-STYLE-MAKE*"
                        settings.CreateMissingStyles = ToBool(value)
                End Select
            Next
            Return settings
        End Function

        Private Shared Function BuildTableConfigs(rows As List(Of String()), ruleHeaderRow As Integer, settingHeaderRow As Integer) As Dictionary(Of Integer, StdTableConfig)
            Dim configs As New Dictionary(Of Integer, StdTableConfig)()
            For tableId = 1 To 6
                configs(tableId) = New StdTableConfig(tableId)
            Next

            Dim header = rows(ruleHeaderRow)
            Dim ruleColumn = FindColumn(header, "Rule")
            If ruleColumn < 0 Then ruleColumn = 0
            Dim tableColumns As New Dictionary(Of Integer, Integer)()
            For tableId = 1 To 6
                tableColumns(tableId) = FindColumn(header, "Table " & tableId.ToString(CultureInfo.InvariantCulture))
            Next

            For i = ruleHeaderRow + 1 To Math.Min(settingHeaderRow - 1, rows.Count - 1)
                Dim label = Clean(GetField(rows(i), ruleColumn)).ToUpperInvariant()
                If label.Length = 0 Then Continue For

                For tableId = 1 To 6
                    Dim tableColumn = tableColumns(tableId)
                    If tableColumn < 0 Then Continue For

                    Dim enabled = ToBool(GetField(rows(i), tableColumn))
                    Dim cfg = configs(tableId)
                    Select Case label
                        Case "BLOCK PROCESSING"
                            cfg.BlockProcessing = enabled
                        Case "STRICTEST MATCH FIRST"
                            cfg.StrictestMatchFirst = enabled
                        Case "STAIR STEPPING"
                            cfg.StairStepping = enabled
                        Case "TABLE 1 EXCLUSION"
                            If enabled Then cfg.ExcludeFromTables.Add(1)
                        Case "TABLE 2 EXCLUSION"
                            If enabled Then cfg.ExcludeFromTables.Add(2)
                        Case "TABLE 3 EXCLUSION"
                            If enabled Then cfg.ExcludeFromTables.Add(3)
                        Case "TABLE 4 EXCLUSION"
                            If enabled Then cfg.ExcludeFromTables.Add(4)
                        Case "TABLE 5 EXCLUSION"
                            If enabled Then cfg.ExcludeFromTables.Add(5)
                    End Select
                Next
            Next

            Return configs
        End Function

        Private Shared Function BuildRules(rows As List(Of String()), tableHeaderRow As Integer) As List(Of StdConversionRule)
            Dim rules As New List(Of StdConversionRule)()
            Dim columns = StdRuleColumnMap.FromHeader(rows(tableHeaderRow))
            Dim i = tableHeaderRow + 1
            Dim index = 1

            While i + 1 < rows.Count
                Dim fromRow = rows(i)
                Dim toRow = rows(i + 1)
                If IsBlankRow(fromRow) OrElse IsKnownSectionHeader(fromRow) Then Exit While

                If SameToken(GetField(fromRow, columns.Side), "FROM") AndAlso SameToken(GetField(toRow, columns.Side), "TO") Then
                    Dim fromPart = EndpointFromRow(fromRow, columns)
                    Dim toPart = EndpointFromRow(toRow, columns)
                    rules.Add(New StdConversionRule(index, fromPart.TableId, fromPart, toPart))
                    index += 1
                    i += 2
                Else
                    i += 1
                End If
            End While

            Return rules
        End Function

        Private Shared Function BuildLayerDefinitions(rows As List(Of String())) As List(Of StdLayerDefinition)
            Dim definitions As New List(Of StdLayerDefinition)()
            Dim headerRow = FindLayerDefinitionHeaderRow(rows)
            If headerRow < 0 Then Return definitions

            Dim header = rows(headerRow)
            Dim layerColumn = FindColumn(header, "Layer", "LayerName", "Layer Name", "Name")
            If layerColumn < 0 Then Return definitions

            Dim descriptionColumn = FindColumn(header, "Description", "Desc")
            Dim colorColumn = FindColumn(header, "Color", "Colour", "Color/Style", "Colour/Style", "ACI", "ColorIndex", "Color Index")
            Dim linetypeColumn = FindColumn(header, "Linetype", "LineType", "Line Type")
            Dim lineweightColumn = FindColumn(header, "Lineweight", "LineWeight", "Line Weight", "Weight", "Width")
            Dim plottableColumn = FindColumn(header, "Plottable", "Plot", "Plotting")
            Dim offColumn = FindColumn(header, "Off", "IsOff")
            Dim frozenColumn = FindColumn(header, "Frozen", "Freeze", "IsFrozen")
            Dim lockedColumn = FindColumn(header, "Locked", "Lock", "IsLocked")

            For i = headerRow + 1 To rows.Count - 1
                Dim row = rows(i)
                If IsBlankRow(row) OrElse IsKnownSectionHeader(row) Then Exit For

                Dim layerName = NullableToken(GetField(row, layerColumn))
                If layerName Is Nothing Then Continue For

                definitions.Add(New StdLayerDefinition With {
                    .Name = layerName,
                    .Description = GetOptionalField(row, descriptionColumn),
                    .ColorToken = GetOptionalField(row, colorColumn),
                    .LinetypeName = GetOptionalField(row, linetypeColumn),
                    .LineweightToken = GetOptionalField(row, lineweightColumn),
                    .PlottableToken = GetOptionalField(row, plottableColumn),
                    .OffToken = GetOptionalField(row, offColumn),
                    .FrozenToken = GetOptionalField(row, frozenColumn),
                    .LockedToken = GetOptionalField(row, lockedColumn)
                })
            Next

            Return definitions
        End Function

        Private Shared Function FindLayerDefinitionHeaderRow(rows As List(Of String())) As Integer
            For i = 0 To rows.Count - 1
                Dim row = rows(i)
                For column = 0 To row.Length - 1
                    If SameToken(GetField(row, column), "tbl_layers") Then
                        If FindColumn(row, "Layer", "LayerName", "Layer Name", "Name") >= 0 Then Return i

                        For nextRow = i + 1 To rows.Count - 1
                            If IsBlankRow(rows(nextRow)) Then Continue For
                            Return nextRow
                        Next
                    End If
                Next
            Next

            For i = 0 To rows.Count - 1
                Dim row = rows(i)
                If FindColumn(row, "Layer", "LayerName", "Layer Name", "Name") >= 0 AndAlso
                   FindColumn(row, "Side") < 0 AndAlso
                   FindColumn(row, "Table") < 0 AndAlso
                   (FindColumn(row, "Color", "Colour", "Color/Style", "Colour/Style", "ACI", "ColorIndex", "Color Index") >= 0 OrElse
                    FindColumn(row, "Linetype", "LineType", "Line Type") >= 0 OrElse
                    FindColumn(row, "Lineweight", "LineWeight", "Line Weight", "Weight", "Width") >= 0 OrElse
                    FindColumn(row, "Plottable", "Plot", "Plotting") >= 0) Then
                    Return i
                End If
            Next

            Return -1
        End Function

        Private Shared Function FindColumn(header As String(), ParamArray names As String()) As Integer
            For i = 0 To header.Length - 1
                Dim normalizedHeader = NormalizeHeader(GetField(header, i))
                For Each name In names
                    If normalizedHeader = NormalizeHeader(name) Then Return i
                Next
            Next

            Return -1
        End Function

        Private Shared Function NormalizeHeader(value As String) As String
            Return Clean(value).Replace(" ", "").Replace("_", "").Replace("-", "").ToUpperInvariant()
        End Function

        Private Shared Function GetOptionalField(row As String(), column As Integer) As String
            If column < 0 Then Return Nothing
            Return NullableToken(GetField(row, column))
        End Function

        Private Shared Function IsBlankRow(row As String()) As Boolean
            If row Is Nothing Then Return True
            For Each field In row
                If Clean(field).Length > 0 Then Return False
            Next
            Return True
        End Function

        Private Shared Function IsKnownSectionHeader(row As String()) As Boolean
            Dim first = Clean(GetField(row, 0))
            Dim second = Clean(GetField(row, 1))
            If SameToken(first, "Rule") AndAlso SameToken(second, "Table 1") Then Return True
            If SameToken(first, "Setting") AndAlso SameToken(second, "Value") Then Return True
            If SameToken(first, "Table") AndAlso SameToken(second, "Side") Then Return True
            If first.StartsWith("tbl_", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function EndpointFromRow(row As String(), columns As StdRuleColumnMap) As StdRuleEndpoint
            Dim tableId = If(columns.Table >= 0, ToInt(GetField(row, columns.Table), 0), 0)
            If tableId <= 0 Then tableId = 1
            Return New StdRuleEndpoint With {
                .TableId = tableId,
                .LayerName = NullableToken(GetField(row, columns.Layer)),
                .ColorToken = NullableToken(GetField(row, columns.Color)),
                .LineType = NullableToken(GetField(row, columns.LineType)),
                .AngleDegrees = NullableDouble(GetField(row, columns.Angle)),
                .DeleteEntity = ToBool(GetField(row, columns.DeleteEntity)),
                .ForceAnnotationToDefault = ToBool(GetField(row, columns.ForceAnnotationToDefault)),
                .TextHeight = NullableDouble(GetField(row, columns.TextHeight)),
                .AnnotationStyle = NullableToken(GetField(row, columns.AnnotationStyle)),
                .Arrowhead1 = NullableToken(GetField(row, columns.Arrowhead1)),
                .Arrowhead2 = NullableToken(GetField(row, columns.Arrowhead2)),
                .TextContents = NullableToken(GetField(row, columns.TextContents)),
                .SearchWindow = NullableDouble(GetField(row, columns.SearchWindow)),
                .SearchGeometryTypes = NullableToken(GetField(row, columns.SearchGeometryTypes)),
                .SearchTargetLayer = NullableToken(GetField(row, columns.SearchTargetLayer)),
                .BlockName = NullableToken(GetField(row, columns.BlockName))
            }
        End Function

        Private Class StdRuleColumnMap
            Public Property Table As Integer
            Public Property Side As Integer
            Public Property Layer As Integer
            Public Property Color As Integer
            Public Property LineType As Integer
            Public Property Angle As Integer
            Public Property DeleteEntity As Integer
            Public Property ForceAnnotationToDefault As Integer
            Public Property TextHeight As Integer
            Public Property AnnotationStyle As Integer
            Public Property Arrowhead1 As Integer
            Public Property Arrowhead2 As Integer
            Public Property TextContents As Integer
            Public Property SearchWindow As Integer
            Public Property SearchGeometryTypes As Integer
            Public Property SearchTargetLayer As Integer
            Public Property BlockName As Integer

            Public Shared Function FromHeader(header As String()) As StdRuleColumnMap
                Return New StdRuleColumnMap With {
                    .Table = FindColumn(header, "Table"),
                    .Side = FindColumn(header, "Side"),
                    .Layer = FindColumn(header, "Layer", "LayerName", "Layer Name"),
                    .Color = FindColumn(header, "Color", "Colour", "ACI", "ColorIndex", "Color Index"),
                    .LineType = FindColumn(header, "Linetype", "LineType", "Line Type"),
                    .Angle = FindColumn(header, "Angle", "AngleDegrees", "Angle Degrees"),
                    .DeleteEntity = FindColumn(header, "Delete entity", "DeleteEntity", "Delete"),
                    .ForceAnnotationToDefault = FindColumn(header, "Force annotation to default", "ForceAnnotationToDefault", "Default Annotation"),
                    .TextHeight = FindColumn(header, "Text height", "TextHeight"),
                    .AnnotationStyle = FindColumn(header, "Anno style", "AnnotationStyle", "Annotation Style", "Text Style", "Style"),
                    .Arrowhead1 = FindColumn(header, "Arrowhead 1", "Arrowhead1"),
                    .Arrowhead2 = FindColumn(header, "Arrowhead 2", "Arrowhead2"),
                    .TextContents = FindColumn(header, "Text contents", "TextContents", "Text Content", "Text"),
                    .SearchWindow = FindColumn(header, "Search window", "SearchWindow"),
                    .SearchGeometryTypes = FindColumn(header, "Search geometry types", "SearchGeometryTypes", "Search Geometry Type"),
                    .SearchTargetLayer = FindColumn(header, "Search target layer", "SearchTargetLayer"),
                    .BlockName = FindColumn(header, "Block name", "BlockName")
                }
            End Function
        End Class

        Public Function GetTableConfig(tableId As Integer) As StdTableConfig
            If TableConfigs.ContainsKey(tableId) Then Return TableConfigs(tableId)
            Return New StdTableConfig(tableId)
        End Function

        Public Function GetRulesForTable(tableId As Integer) As List(Of StdConversionRule)
            If _rulesByTable.ContainsKey(tableId) Then Return _rulesByTable(tableId)
            Return New List(Of StdConversionRule)()
        End Function

        Public Function FindBestRule(props As StdEntityProperties, tableId As Integer, proximityIndex As StdProximityIndex, lastRuleIndex As Integer) As StdMatchResult
            Dim cfg = GetTableConfig(tableId)
            Dim best As StdConversionRule = Nothing
            Dim bestScore = -1

            For Each rule In GetRulesForTable(tableId)
                If cfg.StairStepping AndAlso lastRuleIndex > 0 AndAlso rule.Index <= lastRuleIndex Then Continue For

                Dim score As Integer = 0
                If Matches(rule.FromPart, props, proximityIndex, score) Then
                    If cfg.StrictestMatchFirst Then
                        If score > bestScore Then
                            best = rule
                            bestScore = score
                        End If
                    Else
                        Return New StdMatchResult(rule, score)
                    End If
                End If
            Next

            If best Is Nothing Then Return Nothing
            Return New StdMatchResult(best, bestScore)
        End Function

        Private Function Matches(criteria As StdRuleEndpoint, props As StdEntityProperties, proximityIndex As StdProximityIndex, ByRef score As Integer) As Boolean
            If criteria.LayerName IsNot Nothing Then
                If Not String.Equals(criteria.LayerName, props.LayerName, StringComparison.OrdinalIgnoreCase) Then Return False
                score += 1
            End If

            If criteria.ColorToken IsNot Nothing Then
                If Not ColorMatches(criteria.ColorToken, props) Then Return False
                score += 1
            End If

            If criteria.LineType IsNot Nothing Then
                If Not String.Equals(criteria.LineType, props.LineType, StringComparison.OrdinalIgnoreCase) Then Return False
                score += 1
            End If

            If criteria.AngleDegrees.HasValue Then
                If Not props.AngleDegrees.HasValue Then Return False
                If AngleDiff(criteria.AngleDegrees.Value, props.AngleDegrees.Value) > Settings.AngleToleranceDegrees Then Return False
                score += 1
            End If

            If criteria.TextHeight.HasValue Then
                If Not props.TextHeight.HasValue Then Return False
                If Math.Abs(criteria.TextHeight.Value - props.TextHeight.Value) > Settings.TextHeightTolerance Then Return False
                score += 1
            End If

            If criteria.AnnotationStyle IsNot Nothing Then
                If Not String.Equals(criteria.AnnotationStyle, props.StyleName, StringComparison.OrdinalIgnoreCase) Then Return False
                score += 1
            End If

            If criteria.TextContents IsNot Nothing Then
                If Not TextMatches(criteria.TextContents, props.TextContents) Then Return False
                score += 1
            End If

            If criteria.BlockName IsNot Nothing Then
                If Not String.Equals(criteria.BlockName, props.BlockName, StringComparison.OrdinalIgnoreCase) Then Return False
                score += 1
            End If

            If criteria.SearchWindow.HasValue AndAlso criteria.SearchWindow.Value > 0.0 Then
                If proximityIndex Is Nothing OrElse Not proximityIndex.HasNearbySelectedGeometry(props, criteria) Then Return False
                score += 1
            End If

            Return True
        End Function

        Private Shared Function ColorMatches(token As String, props As StdEntityProperties) As Boolean
            If String.Equals(token, "BYLAYER", StringComparison.OrdinalIgnoreCase) Then Return props.ColorIndex = 256
            If String.Equals(token, "BYBLOCK", StringComparison.OrdinalIgnoreCase) Then Return props.ColorIndex = 0

            If token.StartsWith("ACI:", StringComparison.OrdinalIgnoreCase) Then
                Dim aci = ToInt(token.Substring(4), Integer.MinValue)
                Return aci <> Integer.MinValue AndAlso props.ColorIndex = aci
            End If

            If token.StartsWith("RGB:", StringComparison.OrdinalIgnoreCase) Then
                Return String.Equals(token, props.RgbToken, StringComparison.OrdinalIgnoreCase)
            End If

            Dim parsed = ToInt(token, Integer.MinValue)
            Return parsed <> Integer.MinValue AndAlso props.ColorIndex = parsed
        End Function

        Private Shared Function TextMatches(findToken As String, text As String) As Boolean
            If text Is Nothing Then Return False
            Dim token = findToken.Trim()
            If token.Length >= 2 AndAlso token.StartsWith("(", StringComparison.Ordinal) AndAlso token.EndsWith(")", StringComparison.Ordinal) Then
                Dim exact = token.Substring(1, token.Length - 2)
                Return String.Equals(exact, text, StringComparison.OrdinalIgnoreCase)
            End If

            Return text.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Friend Shared Function GeometryTypeMatches(token As String, entityType As String) As Boolean
            Dim up = token.Trim().ToUpperInvariant()
            Select Case up
                Case "LINE/POLYLINE"
                    Return entityType = "LINE" OrElse entityType = "POLYLINE"
                Case "LWPOLYLINE", "2DPOLYLINE"
                    Return entityType = "POLYLINE"
                Case Else
                    Return String.Equals(up, entityType, StringComparison.OrdinalIgnoreCase)
            End Select
        End Function

        Public Shared Function GetField(row As String(), index As Integer) As String
            If row Is Nothing OrElse index < 0 OrElse index >= row.Length Then Return ""
            Return If(row(index), "")
        End Function

        Public Shared Function NullableToken(value As String) As String
            Dim cleaned = Clean(value)
            If cleaned.Length = 0 Then Return Nothing
            Select Case cleaned.ToUpperInvariant()
                Case "NA", "N/A", "NONE", "NULL"
                    Return Nothing
            End Select
            Return cleaned
        End Function

        Public Shared Function Clean(value As String) As String
            If value Is Nothing Then Return ""
            Return value.Trim()
        End Function

        Public Shared Function ToBool(value As String) As Boolean
            Select Case Clean(value).ToUpperInvariant()
                Case "TRUE", "T", "YES", "Y", "1"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function ToInt(value As String, defaultValue As Integer) As Integer
            Dim result As Integer
            If Integer.TryParse(Clean(value), NumberStyles.Integer, CultureInfo.InvariantCulture, result) Then Return result
            Return defaultValue
        End Function

        Public Shared Function ToDouble(value As String, defaultValue As Double) As Double
            Dim result As Double
            If Double.TryParse(Clean(value), NumberStyles.Float, CultureInfo.InvariantCulture, result) Then Return result
            Return defaultValue
        End Function

        Public Shared Function NullableDouble(value As String) As Double?
            Dim token = NullableToken(value)
            If token Is Nothing Then Return Nothing
            Dim result As Double
            If Double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then Return result
            Return Nothing
        End Function

        Public Shared Function SameToken(left As String, right As String) As Boolean
            Return String.Equals(Clean(left), Clean(right), StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function AngleDiff(a As Double, b As Double) As Double
            Dim diff = Math.Abs(a - b) Mod 360.0
            If diff > 180.0 Then diff = 360.0 - diff
            Return diff
        End Function

        Public Shared Function ReplaceText(source As String, findToken As String, replacement As String) As String
            If source Is Nothing Then Return source
            Dim token = Clean(findToken)
            If token.Length = 0 Then Return source

            If token.Length >= 2 AndAlso token.StartsWith("(", StringComparison.Ordinal) AndAlso token.EndsWith(")", StringComparison.Ordinal) Then
                Dim exact = token.Substring(1, token.Length - 2)
                If String.Equals(exact, source, StringComparison.OrdinalIgnoreCase) Then Return If(replacement, "")
                Return source
            End If

            Dim idx = source.IndexOf(token, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then Return source

            Dim output = source
            While idx >= 0
                output = output.Substring(0, idx) & If(replacement, "") & output.Substring(idx + token.Length)
                idx = output.IndexOf(token, idx + If(replacement, "").Length, StringComparison.OrdinalIgnoreCase)
            End While

            Return output
        End Function
    End Class

    Friend Class StdProximityIndex
        Private ReadOnly _cellSize As Double
        Private ReadOnly _cells As New Dictionary(Of String, List(Of StdEntityProperties))(StringComparer.Ordinal)

        Private Sub New(cellSize As Double)
            _cellSize = Math.Max(cellSize, 1.0)
        End Sub

        Public Shared Function Build(items As IEnumerable(Of StdEntityProperties), maxSearchWindow As Double) As StdProximityIndex
            Dim index As New StdProximityIndex(maxSearchWindow)
            For Each item In items
                If item Is Nothing OrElse Not item.Origin.HasValue Then Continue For
                index.Add(item)
            Next
            Return index
        End Function

        Private Sub Add(item As StdEntityProperties)
            Dim key = CellKey(item.Origin.Value)
            If Not _cells.ContainsKey(key) Then
                _cells(key) = New List(Of StdEntityProperties)()
            End If

            _cells(key).Add(item)
        End Sub

        Public Function HasNearbySelectedGeometry(props As StdEntityProperties, criteria As StdRuleEndpoint) As Boolean
            If props Is Nothing OrElse Not props.Origin.HasValue OrElse Not criteria.SearchWindow.HasValue Then Return False

            Dim half = criteria.SearchWindow.Value / 2.0
            Dim origin = props.Origin.Value
            Dim minX = origin.X - half
            Dim maxX = origin.X + half
            Dim minY = origin.Y - half
            Dim maxY = origin.Y + half
            Dim minCellX = CellCoordinate(minX)
            Dim maxCellX = CellCoordinate(maxX)
            Dim minCellY = CellCoordinate(minY)
            Dim maxCellY = CellCoordinate(maxY)

            For cellX = minCellX To maxCellX
                For cellY = minCellY To maxCellY
                    Dim bucketKey = CellKey(cellX, cellY)
                    If Not _cells.ContainsKey(bucketKey) Then Continue For

                    For Each candidate In _cells(bucketKey)
                        If candidate.ObjectId = props.ObjectId Then Continue For
                        If criteria.SearchTargetLayer IsNot Nothing AndAlso
                           Not String.Equals(criteria.SearchTargetLayer, candidate.LayerName, StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        If criteria.SearchGeometryTypes IsNot Nothing AndAlso
                           Not StandardsConversionRules.GeometryTypeMatches(criteria.SearchGeometryTypes, candidate.EntityType) Then
                            Continue For
                        End If

                        If Not candidate.Origin.HasValue Then Continue For
                        Dim point = candidate.Origin.Value
                        If point.X >= minX AndAlso point.X <= maxX AndAlso point.Y >= minY AndAlso point.Y <= maxY Then Return True
                    Next
                Next
            Next

            Return False
        End Function

        Private Function CellKey(point As Point3d) As String
            Return CellKey(CellCoordinate(point.X), CellCoordinate(point.Y))
        End Function

        Private Shared Function CellKey(x As Integer, y As Integer) As String
            Return x.ToString(CultureInfo.InvariantCulture) & ":" & y.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Function CellCoordinate(value As Double) As Integer
            Return CInt(Math.Floor(value / _cellSize))
        End Function
    End Class

    Friend Class StdConversionSettings
        Public Property LogLevel As Integer = 2
        Public Property DebugLookup As Boolean = False
        Public Property DebugLookupLimit As Integer = 4
        Public Property AngleToleranceDegrees As Double = 5.0
        Public Property TextHeightTolerance As Double = 0.1
        Public Property TextRotationToleranceDegrees As Double = 5.0
        Public Property CreateMissingLayers As Boolean = True
        Public Property CreateMissingStyles As Boolean = True
    End Class

    Friend Class StdLayerDefinition
        Public Property Name As String
        Public Property Description As String
        Public Property ColorToken As String
        Public Property LinetypeName As String
        Public Property LineweightToken As String
        Public Property PlottableToken As String
        Public Property OffToken As String
        Public Property FrozenToken As String
        Public Property LockedToken As String
    End Class

    Friend Class StdTableConfig
        Public Sub New(tableId As Integer)
            Me.TableId = tableId
        End Sub

        Public Property TableId As Integer
        Public Property BlockProcessing As Boolean = False
        Public Property StrictestMatchFirst As Boolean = True
        Public Property StairStepping As Boolean = True
        Public ReadOnly Property ExcludeFromTables As New HashSet(Of Integer)()
    End Class

    Friend Class StdRuleEndpoint
        Public Property TableId As Integer
        Public Property LayerName As String
        Public Property ColorToken As String
        Public Property LineType As String
        Public Property AngleDegrees As Double?
        Public Property DeleteEntity As Boolean
        Public Property ForceAnnotationToDefault As Boolean
        Public Property TextHeight As Double?
        Public Property AnnotationStyle As String
        Public Property Arrowhead1 As String
        Public Property Arrowhead2 As String
        Public Property TextContents As String
        Public Property SearchWindow As Double?
        Public Property SearchGeometryTypes As String
        Public Property SearchTargetLayer As String
        Public Property BlockName As String
    End Class

    Friend Class StdConversionRule
        Public Sub New(index As Integer, tableId As Integer, fromPart As StdRuleEndpoint, toPart As StdRuleEndpoint)
            Me.Index = index
            Me.TableId = tableId
            Me.FromPart = fromPart
            Me.ToPart = toPart
        End Sub

        Public ReadOnly Property Index As Integer
        Public ReadOnly Property TableId As Integer
        Public ReadOnly Property FromPart As StdRuleEndpoint
        Public ReadOnly Property ToPart As StdRuleEndpoint
    End Class

    Friend Class StdMatchResult
        Public Sub New(rule As StdConversionRule, score As Integer)
            Me.Rule = rule
            Me.Score = score
        End Sub

        Public ReadOnly Property Rule As StdConversionRule
        Public ReadOnly Property Score As Integer
    End Class

    Friend Class StdEntityProperties
        Public Property ObjectId As ObjectId
        Public Property EntityType As String
        Public Property LayerName As String
        Public Property ColorIndex As Integer
        Public Property RgbToken As String
        Public Property LineType As String
        Public Property AngleDegrees As Double?
        Public Property TextHeight As Double?
        Public Property StyleName As String
        Public Property TextContents As String
        Public Property BlockName As String
        Public Property Origin As Point3d?

        Public Shared Function FromEntity(ent As Entity, tr As Transaction) As StdEntityProperties
            Dim props As New StdEntityProperties With {
                .ObjectId = ent.ObjectId,
                .EntityType = NormalizeEntityType(ent),
                .LayerName = ent.Layer,
                .ColorIndex = ent.ColorIndex,
                .LineType = ent.Linetype,
                .Origin = GetOrigin(ent)
            }

            Try
                Dim c = ent.Color
                If c IsNot Nothing AndAlso c.ColorMethod = ColorMethod.ByColor Then
                    props.RgbToken = String.Format(CultureInfo.InvariantCulture, "RGB:{0},{1},{2}", CInt(c.Red), CInt(c.Green), CInt(c.Blue))
                End If
            Catch
                props.RgbToken = Nothing
            End Try

            Dim ln = TryCast(ent, Line)
            If ln IsNot Nothing Then
                props.AngleDegrees = RadiansToDegrees(Math.Atan2(ln.EndPoint.Y - ln.StartPoint.Y, ln.EndPoint.X - ln.StartPoint.X))
            End If

            Dim txt = TryCast(ent, DBText)
            If txt IsNot Nothing Then
                props.TextHeight = txt.Height
                props.StyleName = ResolveTextStyleName(txt.TextStyleId, tr)
                props.TextContents = txt.TextString
            End If

            Dim mt = TryCast(ent, MText)
            If mt IsNot Nothing Then
                props.TextHeight = mt.TextHeight
                props.StyleName = ResolveTextStyleName(mt.TextStyleId, tr)
                props.TextContents = mt.Contents
            End If

            Dim br = TryCast(ent, BlockReference)
            If br IsNot Nothing Then
                props.BlockName = ResolveBlockName(br, tr)
            End If

            Return props
        End Function

        Private Shared Function NormalizeEntityType(ent As Entity) As String
            If TypeOf ent Is Line Then Return "LINE"
            If TypeOf ent Is Circle Then Return "CIRCLE"
            If TypeOf ent Is Arc Then Return "ARC"
            If TypeOf ent Is DBText Then Return "TEXT"
            If TypeOf ent Is MText Then Return "MTEXT"
            If TypeOf ent Is BlockReference Then Return "INSERT"
            If TypeOf ent Is Polyline OrElse TypeOf ent Is Polyline2d Then Return "POLYLINE"
            If TypeOf ent Is Hatch Then Return "HATCH"
            If ent.GetType().Name.IndexOf("Mline", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "MLINE"
            If ent.GetType().Name.IndexOf("MLeader", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "MLEADER"
            If ent.GetType().Name.IndexOf("Leader", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "LEADER"
            If ent.GetType().Name.IndexOf("Dimension", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "DIMENSION"
            Return ent.GetType().Name.ToUpperInvariant()
        End Function

        Private Shared Function GetOrigin(ent As Entity) As Point3d?
            Dim ln = TryCast(ent, Line)
            If ln IsNot Nothing Then Return ln.StartPoint

            Dim cir = TryCast(ent, Circle)
            If cir IsNot Nothing Then Return cir.Center

            Dim arcEnt = TryCast(ent, Arc)
            If arcEnt IsNot Nothing Then Return arcEnt.StartPoint

            Dim txt = TryCast(ent, DBText)
            If txt IsNot Nothing Then Return txt.Position

            Dim mt = TryCast(ent, MText)
            If mt IsNot Nothing Then Return mt.Location

            Dim br = TryCast(ent, BlockReference)
            If br IsNot Nothing Then Return br.Position

            Dim pl = TryCast(ent, Polyline)
            If pl IsNot Nothing AndAlso pl.NumberOfVertices > 0 Then Return pl.GetPoint3dAt(0)

            Try
                Dim ext = ent.GeometricExtents
                Return New Point3d((ext.MinPoint.X + ext.MaxPoint.X) / 2.0, (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0, (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ResolveTextStyleName(styleId As ObjectId, tr As Transaction) As String
            If styleId.IsNull Then Return Nothing
            Try
                Dim rec = TryCast(tr.GetObject(styleId, OpenMode.ForRead, False), TextStyleTableRecord)
                If rec IsNot Nothing Then Return rec.Name
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function ResolveBlockName(blockRef As BlockReference, tr As Transaction) As String
            Try
                Dim btr = TryCast(tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead, False), BlockTableRecord)
                If btr IsNot Nothing Then Return btr.Name
            Catch
            End Try

            Try
                Dim btr = TryCast(tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead, False), BlockTableRecord)
                If btr IsNot Nothing Then Return btr.Name
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function RadiansToDegrees(radians As Double) As Double
            Dim degrees = radians * 180.0 / Math.PI
            If degrees < 0.0 Then degrees += 360.0
            Return degrees
        End Function
    End Class
End Namespace
