Imports System
Imports System.IO
Imports System.Drawing.Printing
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
Imports exception = Teigha.Runtime.Exception
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
    Public NetworkUNCPathForEgnyte As String ' Store the full UNC base path

    Public LispCompany As String
    Public LispTeam As String

    Public PlotStyleName As String


    Public Declare Function WNetGetConnection Lib "mpr.dll" Alias _
         "WNetGetConnectionA" (ByVal lpszLocalName As String,
         ByVal lpszRemoteName As StringBuilder, ByRef cbRemoteName As Integer) As Integer

    Public Sub InitializeArcxisPaths()
        ' Automatically configure Arcxis paths if not already present
        Dim currentPaths As List(Of String) = GetPaths()
        Dim hasArcxisPath As Boolean = False

        ' Check if ANY path contains the Arcxis engineering path
        For Each curpath As String In currentPaths
            If Not String.IsNullOrWhiteSpace(curpath) AndAlso
               curpath.IndexOf("arcxis\engineering", StringComparison.OrdinalIgnoreCase) >= 0 Then
                hasArcxisPath = True
                Exit For
            End If
        Next

        ' Only call SupplementalPaths if no Arcxis path exists or list is empty
        If Not hasArcxisPath OrElse currentPaths.Count = 0 Then
            SupplementalPaths(currentPaths)
        End If
    End Sub

    Public Sub UNCPath()
        Dim answer As String
        For Each drv In IO.DriveInfo.GetDrives()
            If drv.DriveType = IO.DriveType.Network Then
                Dim UncPath As New System.Text.StringBuilder(255)
                WNetGetConnection(drv.Name.Replace("\", ""), UncPath, UncPath.Capacity)
                answer = UncPath.ToString
                answer = LCase(answer)

                ' Check if the UNC path starts with \\egnytedrive\
                If answer.StartsWith("\\egnytedrive\") Then
                    ' Now check if \shared exists under this path
                    Dim sharedPath As String = Global.System.IO.Path.Combine(answer, "shared")
                    If Directory.Exists(sharedPath) Then
                        NetworkLetterForEgnyte = drv.Name
                        NetworkUNCPathForEgnyte = sharedPath ' Store the full path to \shared
                        Exit For ' Found it, no need to continue
                    End If
                End If
            End If
        Next
    End Sub

    Private Function ValidateSealsAccess() As Boolean
        ' Check if the seals directory exists
        If String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            Return False
        End If

        Dim sealsPath As String = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\templates\engineering\sealsoriginal")

        If Not Directory.Exists(sealsPath) Then
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acEd As Editor = acDoc.Editor
            Dim errorMsg As String = vbCrLf & "Active access to seals drive does not exist." & vbCrLf &
                                     "Please reach out to IT for access to:" & vbCrLf & vbCrLf &
                                     Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\Templates\Engineering\SealsOriginal")
            acEd.WriteMessage(errorMsg)
            Return False
        End If

        Return True
    End Function

    Private Function GetPaths() As List(Of String)
        Dim CurTrustPath As String = Application.GetSystemVariable("SRCHPATH")
        Dim CurPaths As String = LCase(CurTrustPath)
        Dim RemCurTrustPathList As List(Of String) = New List(Of String)
        Dim isittrue As Integer = 0
        RemCurTrustPathList = CurPaths.Split(";").ToList
        Dim CleanedRemCurTrustPathList As New List(Of String)

        For Each entry In RemCurTrustPathList
            If entry <> "" Then
                If Not CleanedRemCurTrustPathList.Contains(entry) Then
                    CleanedRemCurTrustPathList.Add(entry)
                End If
            End If
        Next
        Return CleanedRemCurTrustPathList

    End Function

    ' Function to validate if a path is reachable
    Private Function IsPathValid(ByVal path As String) As Boolean
        Try
            ' Normalize the path
            Dim normalizedPath As String = Global.System.IO.Path.GetFullPath(path)

            ' Check if the path exists
            Return Directory.Exists(normalizedPath) OrElse File.Exists(normalizedPath)
        Catch ex As Exception
            ' Handle any exception (e.g., path format is invalid)
            Return False
        End Try
    End Function

    ' Subroutine to show a message box for path errors
    Private Sub ShowPathErrorMessage(ByVal invalidPaths As List(Of String))
        Try
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acEd As Editor = acDoc.Editor

            ' Build the error message
            Dim errorMessage As String = "The following paths are invalid or unreachable:" & Environment.NewLine
            For Each path In invalidPaths
                errorMessage &= "- " & path & Environment.NewLine
            Next
            errorMessage &= "Please check your network connections or contact support."

            ' Show the message box
            acEd.WriteMessage(errorMessage)
        Catch ex As exception
            ' Ignore any exception in the error handling code
        End Try
    End Sub

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

        ' Build dynamic paths based on discovered Egnyte configuration
        ' Add mapped drive path only if we actually found a mapped letter
        If Not String.IsNullOrEmpty(NetworkLetterForEgnyte) Then
            Dim mappedEgnyte = Global.System.IO.Path.Combine(NetworkLetterForEgnyte, "shared\arcxis\engineering\drafting standards\CAD Lisp Routines\BricsCad")
            TrustPathListAdd.Add(mappedEgnyte)
        End If

        ' Add UNC paths using the discovered base path
        If Not String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            TrustPathListAdd.Add(Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "arcxis\engineering\drafting standards\cad lisp routines\BricsCad"))
            TrustPathListAdd.Add(Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "arcxis\engineering\drafting standards\cad lisp routines\BricsCad\Support Files"))

            ' Only add seals path if it exists
            Dim sealsPath As String = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\templates\engineering\sealsoriginal")
            If Directory.Exists(sealsPath) Then
                TrustPathListAdd.Add(sealsPath)
            Else
                ' Notify user that seals access is not available
                Dim sealsFullPath As String = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\Templates\Engineering\SealsOriginal")
                Dim errorMsg As String = vbCrLf & "Active access to seals drive does not exist." & vbCrLf &
                                         "Please reach out to IT for access to:" & vbCrLf & vbCrLf & sealsFullPath
                aced.WriteMessage(errorMsg)
            End If
        End If

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

        ' Build dynamic printer paths using discovered Egnyte configuration
        Dim NewPrinterStyleSheetDir As String = ""
        Dim NewPrinterConfigDir As String = ""

        If Not String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            NewPrinterStyleSheetDir = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "Arcxis\Engineering\Drafting Standards\CAD Plot Styles")
            NewPrinterConfigDir = NewPrinterStyleSheetDir
        End If

        If Not String.IsNullOrEmpty(NewPrinterConfigDir) Then
            acadApp.Preferences.Files.PrinterConfigPath = NewPrinterConfigDir
            acadApp.Preferences.Files.PrinterStyleSheetPath = NewPrinterStyleSheetDir
        End If

    End Sub
    Public Sub AddNewPaths(Optional ByVal CurTrustPathList As List(Of String) = Nothing)


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
            Dim hasegnyte As Boolean = p.IndexOf("\\egnytedrive\", StringComparison.OrdinalIgnoreCase) >= 0
            Dim HasEngyteDriveLetter As Boolean = p.IndexOf(":\Shared\Arcxis\", StringComparison.OrdinalIgnoreCase) >= 0

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

        ' Build dynamic paths based on discovered Egnyte configuration
        ' Add mapped drive path only if we actually found a mapped letter
        If Not String.IsNullOrEmpty(NetworkLetterForEgnyte) Then
            Dim mappedEgnyte = Global.System.IO.Path.Combine(NetworkLetterForEgnyte, "shared\arcxis\engineering\drafting standards\CAD Lisp Routines\BricsCad")
            TrustPathListAdd.Add(mappedEgnyte)
        End If

        ' Add UNC paths using the discovered base path
        If Not String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            TrustPathListAdd.Add(Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "arcxis\engineering\drafting standards\cad lisp routines\BricsCad"))
            TrustPathListAdd.Add(Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "arcxis\engineering\drafting standards\cad lisp routines\BricsCad\Support Files"))

            ' Only add seals path if it exists
            Dim sealsPath As String = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\templates\engineering\sealsoriginal")
            If Directory.Exists(sealsPath) Then
                TrustPathListAdd.Add(sealsPath)
            Else
                ' Notify user that seals access is not available
                Dim sealsFullPath As String = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "onyx file system\Templates\Engineering\SealsOriginal")
                Dim errorMsg As String = vbCrLf & "Active access to seals drive does not exist." & vbCrLf &
                                         "Please reach out to IT for access to:" & vbCrLf & vbCrLf & sealsFullPath
                aced.WriteMessage(errorMsg)
            End If
        End If

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

        ' Build dynamic printer paths using discovered Egnyte configuration
        Dim NewPrinterStyleSheetDir As String = ""
        Dim NewPrinterConfigDir As String = ""

        If Not String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            NewPrinterStyleSheetDir = Global.System.IO.Path.Combine(NetworkUNCPathForEgnyte, "Arcxis\Engineering\Drafting Standards\CAD Plot Styles")
            NewPrinterConfigDir = NewPrinterStyleSheetDir
        End If

        If Not String.IsNullOrEmpty(NewPrinterConfigDir) Then
            acadApp.Preferences.Files.PrinterConfigPath = NewPrinterConfigDir
            acadApp.Preferences.Files.PrinterStyleSheetPath = NewPrinterStyleSheetDir
        End If

    End Sub

End Module
