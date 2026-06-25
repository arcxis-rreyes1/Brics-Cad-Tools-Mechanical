Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports PdfSharp.Pdf.IO
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document
Imports Exception = Teigha.Runtime.Exception

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.PdfInsertCommands))>
Namespace Arcxis_Cad_Tools

    Public Class PdfInsertCommands

        <CommandMethod("APA", CommandFlags.Modal)>
        Public Sub AttachPdfCommand()
            RunInteractivePdfCommand(isImport:=False)
        End Sub

        <CommandMethod("APP", CommandFlags.Modal)>
        Public Sub ImportPdfCommand()
            RunInteractivePdfCommand(isImport:=True)
        End Sub

        Public Shared Function GetPdfPageCount(pdfPath As String) As Integer
            If String.IsNullOrWhiteSpace(pdfPath) OrElse Not File.Exists(pdfPath) Then
                Throw New FileNotFoundException("PDF file not found.", pdfPath)
            End If

            Using doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly)
                Return doc.PageCount
            End Using
        End Function

        Public Shared Function AttachPdfPage(
            db As Database,
            pdfFile As String,
            pageNumber As Integer,
            insPoint As Point3d,
            Optional scale As Double = 1.0,
            Optional rotationDegrees As Double = 0.0
        ) As ObjectId

            If pageNumber < 1 Then
                Throw New ArgumentOutOfRangeException(NameOf(pageNumber), "Page number must be >= 1.")
            End If

            Using tr = db.TransactionManager.StartTransaction()
                Dim defId = CreatePdfDefinition(db, pdfFile, pageNumber, tr)
                Dim refId = CreatePdfReference(defId, db, pageNumber, insPoint, scale, rotationDegrees, tr)
                tr.Commit()
                Return refId
            End Using
        End Function

        Public Shared Sub AttachPdfPages(
            db As Database,
            pdfFile As String,
            pageNumbers As IEnumerable(Of Integer),
            firstInsPoint As Point3d,
            verticalSpacing As Double,
            Optional scale As Double = 1.0,
            Optional rotationDegrees As Double = 0.0
        )
            Dim pageList = pageNumbers.ToList()
            If pageList.Count = 0 Then Return

            Using docLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()
                    Dim index As Integer = 0
                    For Each pageNumber In pageList
                        Dim insPoint = New Point3d(
                            firstInsPoint.X,
                            firstInsPoint.Y - (index * verticalSpacing),
                            firstInsPoint.Z)

                        Dim defId = CreatePdfDefinition(db, pdfFile, pageNumber, tr)
                        CreatePdfReference(defId, db, pageNumber, insPoint, scale, rotationDegrees, tr)
                        index += 1
                    Next

                    tr.Commit()
                End Using
            End Using
        End Sub

        Public Shared Sub ImportPdfPage(
            doc As Document,
            pdfFile As String,
            pageNumber As Integer,
            insPoint As Point3d,
            Optional scale As Double = 1.0,
            Optional rotationDegrees As Double = 0.0
        )
            doc.SendStringToExecute(
                BuildPdfImportCommand(pdfFile, pageNumber, insPoint, scale, rotationDegrees),
                True, False, False)
        End Sub

        Public Shared Sub ImportPdfPages(
            doc As Document,
            pdfFile As String,
            pageNumbers As IEnumerable(Of Integer),
            firstInsPoint As Point3d,
            verticalSpacing As Double,
            Optional scale As Double = 1.0,
            Optional rotationDegrees As Double = 0.0
        )
            Dim cmd As New System.Text.StringBuilder()
            Dim index As Integer = 0

            For Each pageNumber In pageNumbers
                Dim insPoint = New Point3d(
                    firstInsPoint.X,
                    firstInsPoint.Y - (index * verticalSpacing),
                    firstInsPoint.Z)

                cmd.Append(BuildPdfImportCommand(pdfFile, pageNumber, insPoint, scale, rotationDegrees))
                index += 1
            Next

            If cmd.Length > 0 Then
                doc.SendStringToExecute(cmd.ToString(), True, False, False)
            End If
        End Sub

        Private Shared Sub RunInteractivePdfCommand(isImport As Boolean)
            Dim doc = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then Return

            Dim ed = doc.Editor
            Dim db = doc.Database

            Dim pdfPath = PromptForPdfFile(ed)
            If String.IsNullOrWhiteSpace(pdfPath) Then Return

            Dim pageCount As Integer
            Try
                pageCount = GetPdfPageCount(pdfPath)
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Could not read PDF: " & ex.Message)
                Return
            End Try

            Dim pages = PromptForPages(ed, pageCount)
            If pages Is Nothing OrElse pages.Count = 0 Then Return

            Dim ppr = ed.GetPoint(vbLf & "Specify insertion point:")
            If ppr.Status <> PromptStatus.OK Then Return

            Dim scale = PromptForScale(ed)
            If Not scale.HasValue Then Return

            Dim rotation = PromptForRotation(ed)
            If Not rotation.HasValue Then Return

            Dim spacing As Double = 0.0
            If pages.Count > 1 Then
                Dim spacingResult = PromptForVerticalSpacing(ed)
                If Not spacingResult.HasValue Then Return
                spacing = spacingResult.Value
            End If

            Dim actionName = If(isImport, "import", "attach")
            Try
                If isImport Then
                    ImportPdfPages(doc, pdfPath, pages, ppr.Value, spacing, scale.Value, rotation.Value)
                Else
                    AttachPdfPages(db, pdfPath, pages, ppr.Value, spacing, scale.Value, rotation.Value)
                End If

                ed.WriteMessage(vbLf & $"PDF {actionName} queued for {pages.Count} page(s) from {Path.GetFileName(pdfPath)}.")
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & $"PDF {actionName} failed: " & ex.Message)
            End Try
        End Sub

        Private Shared Function PromptForPdfFile(ed As Editor) As String
            Dim ofd As New OpenFileDialog() With {
                .Title = "Select PDF file",
                .Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                .Multiselect = False
            }

            If ofd.ShowDialog() <> DialogResult.OK Then
                ed.WriteMessage(vbLf & "Cancelled.")
                Return Nothing
            End If

            Return ofd.FileName
        End Function

        Private Shared Function PromptForPages(ed As Editor, pageCount As Integer) As List(Of Integer)
            If pageCount <= 1 Then
                Return New List(Of Integer) From {1}
            End If

            Dim pko As New PromptKeywordOptions(vbLf & $"PDF has {pageCount} pages. Import [Single/All]:", "Single All")
            pko.Keywords.Default = "Single"
            pko.AllowNone = False

            Dim pkr = ed.GetKeywords(pko)
            If pkr.Status <> PromptStatus.OK Then Return Nothing

            If pkr.StringResult.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                Return Enumerable.Range(1, pageCount).ToList()
            End If

            Dim pio As New PromptIntegerOptions(vbLf & $"Enter page number <1-{pageCount}>:")
            pio.LowerLimit = 1
            pio.UpperLimit = pageCount
            pio.DefaultValue = 1

            Dim pir = ed.GetInteger(pio)
            If pir.Status <> PromptStatus.OK Then Return Nothing

            Return New List(Of Integer) From {pir.Value}
        End Function

        Private Shared Function PromptForScale(ed As Editor) As Double?
            Dim pdo As New PromptDoubleOptions(vbLf & "Specify scale factor <1>:")
            pdo.DefaultValue = 1.0
            pdo.AllowNegative = False
            pdo.AllowZero = False
            pdo.AllowNone = True

            Dim pdr = ed.GetDouble(pdo)
            If pdr.Status = PromptStatus.None Then Return 1.0
            If pdr.Status <> PromptStatus.OK Then Return Nothing
            Return pdr.Value
        End Function

        Private Shared Function PromptForRotation(ed As Editor) As Double?
            Dim pdo As New PromptDoubleOptions(vbLf & "Specify rotation angle in degrees <0>:")
            pdo.DefaultValue = 0.0
            pdo.AllowNone = True

            Dim pdr = ed.GetDouble(pdo)
            If pdr.Status = PromptStatus.None Then Return 0.0
            If pdr.Status <> PromptStatus.OK Then Return Nothing
            Return pdr.Value
        End Function

        Private Shared Function PromptForVerticalSpacing(ed As Editor) As Double?
            Dim pdo As New PromptDoubleOptions(vbLf & "Vertical spacing between pages <0>:")
            pdo.DefaultValue = 0.0
            pdo.AllowNegative = False
            pdo.AllowNone = True

            Dim pdr = ed.GetDouble(pdo)
            If pdr.Status = PromptStatus.None Then Return 0.0
            If pdr.Status <> PromptStatus.OK Then Return Nothing
            Return pdr.Value
        End Function

        Private Shared Function CreatePdfDefinition(
            db As Database,
            pdfFileName As String,
            pageNumber As Integer,
            tr As Transaction
        ) As ObjectId

            Dim dic = GetPdfDefinitionDictionary(db, tr)
            Dim baseName = Path.GetFileNameWithoutExtension(pdfFileName)
            Dim defKey = $"{baseName} - {pageNumber}"
            Dim suffix As Integer = 0

            While dic.Contains(defKey)
                suffix += 1
                defKey = $"{baseName}_{suffix:D2} - {pageNumber}"
            End While

            Dim pdfDef As New PdfDefinition()
            pdfDef.SourceFileName = pdfFileName

            Dim defId = dic.SetAt(defKey, pdfDef)
            tr.AddNewlyCreatedDBObject(pdfDef, True)
            Return defId
        End Function

        Private Shared Function CreatePdfReference(
            defId As ObjectId,
            db As Database,
            pageNumber As Integer,
            insPoint As Point3d,
            scale As Double,
            rotationDegrees As Double,
            tr As Transaction
        ) As ObjectId

            Dim space = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            Dim pdfRef As New PdfReference()
            pdfRef.DefinitionId = defId
            pdfRef.SetDatabaseDefaults(db)
            pdfRef.Position = insPoint
            pdfRef.NameOfSheet = pageNumber.ToString(CultureInfo.InvariantCulture)
            pdfRef.ScaleFactors = New Scale3d(scale, scale, scale)
            pdfRef.Rotation = rotationDegrees * (Math.PI / 180.0)

            Dim refId = space.AppendEntity(pdfRef)
            tr.AddNewlyCreatedDBObject(pdfRef, True)
            Return refId
        End Function

        Private Shared Function GetPdfDefinitionDictionary(db As Database, tr As Transaction) As DBDictionary
            Dim namedDic = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
            Dim pdfDicKey = UnderlayDefinition.GetDictionaryKey(GetType(PdfDefinition))

            If Not namedDic.Contains(pdfDicKey) Then
                namedDic.UpgradeOpen()
                Dim newDic As New DBDictionary()
                Dim dicId = namedDic.SetAt(pdfDicKey, newDic)
                tr.AddNewlyCreatedDBObject(newDic, True)
                Return CType(tr.GetObject(dicId, OpenMode.ForWrite), DBDictionary)
            End If

            Return CType(tr.GetObject(namedDic.GetAt(pdfDicKey), OpenMode.ForWrite), DBDictionary)
        End Function

        Private Shared Function BuildPdfImportCommand(
            pdfPath As String,
            pageNumber As Integer,
            insPoint As Point3d,
            scale As Double,
            rotationDegrees As Double
        ) As String

            Return "_.-PDFIMPORT" & vbLf &
                   "_F" & vbLf &
                   pdfPath & vbLf &
                   pageNumber.ToString(CultureInfo.InvariantCulture) & vbLf &
                   PointToCommandString(insPoint) & vbLf &
                   scale.ToString(CultureInfo.InvariantCulture) & vbLf &
                   rotationDegrees.ToString(CultureInfo.InvariantCulture) & vbLf
        End Function

        Private Shared Function PointToCommandString(pt As Point3d) As String
            Return pt.X.ToString(CultureInfo.InvariantCulture) & "," &
                   pt.Y.ToString(CultureInfo.InvariantCulture) & "," &
                   pt.Z.ToString(CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
