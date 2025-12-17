Imports System
Imports System.IO
Imports System.Drawing.Printing
Imports System.Windows.Documents
Imports System.Linq
Imports System.Text
Imports System.Drawing.Color
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices
Module Module_Arcxis_TB
    Public ATB_OldName As String
    Public ATB_RevNum1 As String
    Public ATB_REVDATE1 As String
    Public ATB_REVDESCRIPTION1 As String
    Public ATB_INITIAL1 As String
    Public ATB_RevNum2 As String
    Public ATB_REVDATE2 As String
    Public ATB_REVDESCRIPTION2 As String
    Public ATB_INITIAL2 As String
    Public ATB_RevNum3 As String
    Public ATB_REVDATE3 As String
    Public ATB_REVDESCRIPTION3 As String
    Public ATB_INITIAL3 As String
    Public ATB_RevNum4 As String
    Public ATB_REVDATE4 As String
    Public ATB_REVDESCRIPTION4 As String
    Public ATB_INITIAL4 As String
    Public ATB_RevNum5 As String
    Public ATB_REVDATE5 As String
    Public ATB_REVDESCRIPTION5 As String
    Public ATB_INITIAL5 As String
    Public ATB_RevNum6 As String
    Public ATB_REVDATE6 As String
    Public ATB_REVDESCRIPTION6 As String
    Public INITIAL6 As String
    Public ATB_RevNum7 As String
    Public ATB_REVDATE7 As String
    Public ATB_REVDESCRIPTION7 As String
    Public ATB_INITIAL7 As String
    Public ATB_lastselected As String

    Public ATB_RevList(7, 4) As String
    Public ATB_blockName As String

    Public ATB_PlotterName As String
    Public ATB_PlotSize As String
    Public ATB_PlotStyleName As String
    Public ATB_PlotSetUpName As String
    Public ATB_WaterMarkText As String

    Public ATB_CurLayoutName As String
    Public ATB_FirLayoutName As String
    Public ATB_SecLayoutName As String
    Public ATB_UserSelect As String
    Public ATB_UserSelect2 As String

    Public ATB_PlanNumber As String
    Public ATB_Elevation As String
    Public ATB_Opt As String
    Public ATB_SubDivision As String
    Public ATB_CustomerName As String
    Public Lot As String
    Public ATB_Block As String
    Public ATB_Section As String
    Public ATB_Drawn As String
    Public ATB_Design As String
    Public ATB_Check As String

    Public ATB_NewName As String
    Public ATB_CustomLayoutLetter As String
    Public ATB_CheckState As String

    Public ATB_CustomLayoutList As New ArrayList
    Public ATB_LayoutList As New List(Of String)
    Public ATB_FirstLayoutName As String

    Public DpisOldName As String
    Public DpisNewName As String

    Public NetworkLetterForEgnyte As String

    Public LispCompany As String
    Public LispTeam As String

    Public Declare Function WNetGetConnection Lib "mpr.dll" Alias _
         "WNetGetConnectionA" (ByVal lpszLocalName As String,
         ByVal lpszRemoteName As StringBuilder, ByRef cbRemoteName As Integer) As Integer

    Public Sub UNCPath()
        Dim answer As String
        For Each drv In IO.DriveInfo.GetDrives()
            If drv.DriveType = IO.DriveType.Network Then
                Dim UncPath As New System.Text.StringBuilder(255)
                WNetGetConnection(drv.Name.Replace("\", ""), UncPath, UncPath.Capacity)
                answer = UncPath.ToString
                answer = LCase(answer)
                If answer = "\\egnytedrive\energyinspectors" Then
                    NetworkLetterForEgnyte = drv.Name
                End If
            End If
        Next
    End Sub

    Private Function GetPaths()

        Dim CurTrustPath As String = Application.GetSystemVariable("SRCHPATH")
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
        GetPaths = CleanedRemCurTrustPathList

    End Function

    <CommandMethod("SupplementalPaths", CommandFlags.Modal)>
    Public Sub SupplementalPaths(Optional ByVal CurTrustPathList As List(Of String) = Nothing)


        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim aced As Editor = acDoc.Editor
        'Dim acpref As AcadPreferences = Application.Preferences
        Dim CleanedTrustPathList As List(Of String) = New List(Of String)

        ' Ensure we have a list to work with
        If CurTrustPathList Is Nothing OrElse CurTrustPathList.Count = 0 Then
            CurTrustPathList = GetPaths()
        End If

        For Each curpath As String In CurTrustPathList
            If String.IsNullOrWhiteSpace(curpath) Then Continue For

            ' Normalize and validate absolute paths only
            Dim p As String = curpath.Trim()
            Dim hasDpis As Boolean = p.IndexOf("dpis", StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasSeal As Boolean = p.IndexOf("seal$", StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasSDrive As Boolean = p.IndexOf("s:\", StringComparison.OrdinalIgnoreCase) >= 0

            If hasDpis OrElse hasSeal OrElse hasSDrive Then
                ' Skip paths containing DPIS, SEAL$, or S:\ (case-insensitive)
            Else
                If Global.System.IO.Path.IsPathRooted(p) Then
                    ' Add only if not already present (case-insensitive)
                    If Not CleanedTrustPathList.Any(Function(x) String.Equals(x, p, StringComparison.OrdinalIgnoreCase)) Then
                        CleanedTrustPathList.Add(p)
                    End If
                End If
            End If
        Next

        UNCPath()

        Dim TrustPathListAdd As List(Of String) = New List(Of String)
        Dim FinalTrustPathList As List(Of String) = New List(Of String)

        ' Known good UNC locations (keep as absolute paths)

        ' Add mapped drive path only if we actually found a mapped letter
        If Not String.IsNullOrEmpty(NetworkLetterForEgnyte) Then
            Dim mappedEgnyte = Global.System.IO.Path.Combine(NetworkLetterForEgnyte, "shared\arcxis\engineering\drafting standards")
            TrustPathListAdd.Add(mappedEgnyte)
        End If
        TrustPathListAdd.Add("\\egnytedrive\energyinspectors\shared\arcxis\engineering\drafting standards\cad lisp routines\BricsCad")
        TrustPathListAdd.Add("\\egnytedrive\energyinspectors\shared\arcxis\engineering\drafting standards")
        TrustPathListAdd.Add("\\egnytedrive\energyinspectors\shared\arcxis\engineering\drafting standards\cad lisp routines\BricsCad\Support Files")
        TrustPathListAdd.Add("\\egnytedrive\energyinspectors\shared\onyx file system\templates\engineering\sealsoriginal")


        For Each candidate In CleanedTrustPathList
            If Not String.IsNullOrWhiteSpace(candidate) AndAlso Global.System.IO.Path.IsPathRooted(candidate) Then
                If Not FinalTrustPathList.Any(Function(pth) String.Equals(pth, candidate, StringComparison.OrdinalIgnoreCase)) Then
                    FinalTrustPathList.Add(candidate)
                End If
            End If
        Next

        For Each candidate As String In TrustPathListAdd
            If Not String.IsNullOrWhiteSpace(candidate) AndAlso Global.System.IO.Path.IsPathRooted(candidate) Then
                If Not FinalTrustPathList.Any(Function(pth) String.Equals(pth, candidate, StringComparison.OrdinalIgnoreCase)) Then
                    FinalTrustPathList.Add(candidate)
                End If
            End If
        Next

        If FinalTrustPathList.Count > 0 Then
            Dim NewTrustPath As String = String.Join(";", FinalTrustPathList)
            ' Use correct system variable name and avoid altering path casing
            Application.SetSystemVariable("SRCHPATH", NewTrustPath)
        End If


        Dim acadApp As Object = Application.AcadApplication


        Dim NewPrinterStyleSheetDir As String = "\\EgnyteDrive\energyinspectors\Shared\Arcxis\Engineering\Drafting Standards\CAD Plot Styles"
        Dim NewPrinterConfigDir As String = "\\EgnyteDrive\energyinspectors\Shared\Arcxis\Engineering\Drafting Standards\CAD Plot Styles"
        Dim NewPrinterDescDir As String = "\\EgnyteDrive\energyinspectors\Shared\Arcxis\Engineering\Drafting Standards\CAD Plot Styles"

        acadApp.Preferences.Files.PrinterConfigPath = NewPrinterConfigDir
        acadApp.Preferences.Files.PrinterStyleSheetPath = NewPrinterStyleSheetDir

    End Sub

End Module
