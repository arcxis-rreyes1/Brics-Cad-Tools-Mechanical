Imports System
Imports System.Reflection
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Imports Exception = Teigha.Runtime.Exception

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

<Assembly: CommandClass(GetType(Arcxis_Cad_Tools.MechanicalFunctions))>

Namespace Arcxis_Cad_Tools
    Public Class MechanicalFunctions
        <CommandMethod("EXCEL2TABLE")>
        Public Sub ExcelSelectionToAcadTable()

            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim db = doc.Database

            ' --- Get Excel selection ---
            Dim xlApp As Excel.Application = Nothing
            Try
                ' Use VB Interaction.GetObject to bind to the running Excel instance (COM ROT)
                xlApp = CType(Interaction.GetObject("", "Excel.Application"), Excel.Application)
            Catch ex As COMException
                ed.WriteMessage(vbLf & "Excel is not running or no selection found.")
                Return
            Catch
                ed.WriteMessage(vbLf & "Excel is not running or no selection found.")
                Return
            End Try

            Dim rng As Excel.Range = TryCast(xlApp.Selection, Excel.Range)
            If rng Is Nothing Then
                ed.WriteMessage(vbLf & "Please select a cell range in Excel before running EXCEL2TABLE.")
                Return
            End If

            ' --------------------------------------------------------------------
            '   Build lists of visible rows/columns
            ' --------------------------------------------------------------------
            Dim visibleRows As New List(Of Excel.Range)
            Dim visibleCols As New List(Of Excel.Range)

            For r As Integer = 1 To rng.Rows.Count
                Dim row As Excel.Range = rng.Rows(r)
                If Not row.EntireRow.Hidden Then
                    visibleRows.Add(row)
                End If
            Next

            For c As Integer = 1 To rng.Columns.Count
                Dim col As Excel.Range = rng.Columns(c)
                If Not col.EntireColumn.Hidden Then
                    visibleCols.Add(col)
                End If
            Next

            If visibleRows.Count = 0 OrElse visibleCols.Count = 0 Then
                ed.WriteMessage(vbLf & "Selection contains no visible rows or columns.")
                Return
            End If

            ' --- Ask insertion point ---
            Dim ppr = ed.GetPoint(vbLf & "Specify insertion point:")
            If ppr.Status <> PromptStatus.OK Then Return

            Using doc.LockDocument()
                Using tr = db.TransactionManager.StartTransaction()

                    Dim bt = CType(tr.GetObject(db.BlockTableId, Teigha.DatabaseServices.OpenMode.ForRead), BlockTable)
                    Dim btr = CType(tr.GetObject(db.CurrentSpaceId, Teigha.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)

                    ' --- Create AutoCAD table with visible row/column count ---
                    Dim tbl As New Table()
                    tbl.TableStyle = db.Tablestyle

                    Dim nRows As Integer = visibleRows.Count
                    Dim nCols As Integer = visibleCols.Count

                    tbl.SetSize(nRows, nCols)

                    ' Default sizes
                    For r = 0 To nRows - 1
                        tbl.SetRowHeight(r, 3.0)
                    Next
                    For c = 0 To nCols - 1
                        tbl.SetColumnWidth(c, 15.0)
                    Next

                    ' --------------------------------------------------------------------
                    '   Fill visible rows & columns ONLY
                    ' --------------------------------------------------------------------
                    Dim acRow As Integer = 0

                    For Each excelRow As Excel.Range In visibleRows
                        Dim acCol As Integer = 0

                        For Each excelCol As Excel.Range In visibleCols
                            ' Get cell
                            Dim cell As Excel.Range = rng.Cells(excelRow.Row - rng.Row + 1,
                                                        excelCol.Column - rng.Column + 1)

                            Dim txt As String = CStr(cell.Text)
                            Dim acCell = tbl.Cells(acRow, acCol)

                            acCell.TextString = txt
                            acCell.TextHeight = 1.5

                            ' Alignment
                            Try
                                Select Case CInt(cell.HorizontalAlignment)
                                    Case Excel.XlHAlign.xlHAlignLeft
                                        acCell.Alignment = CellAlignment.MiddleLeft
                                    Case Excel.XlHAlign.xlHAlignRight
                                        acCell.Alignment = CellAlignment.MiddleRight
                                    Case Excel.XlHAlign.xlHAlignCenter
                                        acCell.Alignment = CellAlignment.MiddleCenter
                                End Select
                            Catch
                            End Try

                            ' Bold = slightly larger text
                            Try
                                If CBool(cell.Font.Bold) Then
                                    acCell.TextHeight = 1.7
                                End If
                            Catch
                            End Try

                            acCol += 1
                        Next

                        acRow += 1
                    Next

                    ' Place table
                    tbl.Position = ppr.Value
                    btr.AppendEntity(tbl)
                    tr.AddNewlyCreatedDBObject(tbl, True)

                    tr.Commit()
                End Using
            End Using

        End Sub

    End Class

End Namespace
