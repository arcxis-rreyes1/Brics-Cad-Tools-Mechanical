Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports Bricscad.ApplicationServices
Imports Bricscad.EditorInput
Imports Application = Bricscad.ApplicationServices.Application

''' <summary>
''' Egnyte discovery, support file search paths (SRCHPATH), and plot style paths.
''' </summary>
Public Module ArcxisPaths
    Public NetworkLetterForEgnyte As String
    Public NetworkUNCPathForEgnyte As String

    Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (
        ByVal lpszLocalName As String,
        ByVal lpszRemoteName As StringBuilder,
        ByRef cbRemoteName As Integer) As Integer

    Public Sub InitializeArcxisPaths()
        EnsureDllSupportPaths()

        Dim currentPaths = GetPathsPreserveCase()
        Dim hasArcxisPath = currentPaths.Any(
            Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso
                          p.IndexOf("arcxis\engineering", StringComparison.OrdinalIgnoreCase) >= 0)

        If Not hasArcxisPath OrElse currentPaths.Count = 0 Then
            RefreshSupportPaths(currentPaths)
        End If
    End Sub

    Public Sub AddNewPaths()
        EnsureDllSupportPaths()
        RefreshSupportPaths()
    End Sub

    Public Sub EnsureDllSupportPaths()
        Try
            Dim dllDir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            If String.IsNullOrWhiteSpace(dllDir) OrElse Not Directory.Exists(dllDir) Then Return

            DiscoverEgnytePaths()

            Dim currentPaths = GetPathsPreserveCase()
            Dim added = False

            For Each pathVariant In PathNormalizer.GetEquivalentPaths(dllDir)
                If Not System.IO.Path.IsPathRooted(pathVariant) Then Continue For
                If Not Directory.Exists(pathVariant) Then Continue For
                If PathNormalizer.ListContainsEquivalent(currentPaths, pathVariant) Then Continue For

                currentPaths.Add(pathVariant)
                added = True
            Next

            If added AndAlso currentPaths.Count > 0 Then
                Application.SetSystemVariable("SRCHPATH", String.Join(";", currentPaths))
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Discover mapped Egnyte drive and UNC base path (\\egnytedrive\...\shared).</summary>
    Public Sub DiscoverEgnytePaths()
        For Each drv In DriveInfo.GetDrives()
            If drv.DriveType <> DriveType.Network Then Continue For

            Dim uncBuilder As New StringBuilder(255)
            WNetGetConnection(drv.Name.Replace("\", ""), uncBuilder, uncBuilder.Capacity)
            Dim answer = uncBuilder.ToString().ToLowerInvariant()

            If Not answer.StartsWith("\\egnytedrive\") Then Continue For

            Dim sharedPath = System.IO.Path.Combine(answer, "shared")
            If Not Directory.Exists(sharedPath) Then Continue For

            NetworkLetterForEgnyte = drv.Name
            NetworkUNCPathForEgnyte = sharedPath
            Exit For
        Next
    End Sub

    ''' <summary>Backward-compatible alias.</summary>
    Public Sub UNCPath()
        DiscoverEgnytePaths()
    End Sub

    Public Sub RefreshSupportPaths(Optional existingPaths As List(Of String) = Nothing)
        Dim doc = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return
        Dim ed = doc.Editor

        If existingPaths Is Nothing OrElse existingPaths.Count = 0 Then
            existingPaths = GetPathsPreserveCase()
        End If

        Dim keptPaths As New List(Of String)()
        For Each curpath In existingPaths
            If String.IsNullOrWhiteSpace(curpath) Then Continue For

            Dim p = curpath.Trim()
            If p.IndexOf("dpis", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
            If p.IndexOf("seal$", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
            If p.IndexOf("s:\", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
            If Not System.IO.Path.IsPathRooted(p) Then Continue For

            If Not keptPaths.Any(Function(x) String.Equals(x, p, StringComparison.OrdinalIgnoreCase)) Then
                keptPaths.Add(p)
            End If
        Next

        DiscoverEgnytePaths()

        Dim pathsToAdd As New List(Of String)()

        If Not String.IsNullOrEmpty(NetworkLetterForEgnyte) Then
            pathsToAdd.Add(System.IO.Path.Combine(
                NetworkLetterForEgnyte,
                "shared\arcxis\engineering\drafting standards\CAD Lisp Routines\BricsCad"))
        End If

        If Not String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then
            pathsToAdd.Add(System.IO.Path.Combine(
                NetworkUNCPathForEgnyte,
                "arcxis\engineering\drafting standards\cad lisp routines\BricsCad"))
            pathsToAdd.Add(System.IO.Path.Combine(
                NetworkUNCPathForEgnyte,
                "arcxis\engineering\drafting standards\cad lisp routines\BricsCad\Support Files"))

            Dim sealsPath = System.IO.Path.Combine(
                NetworkUNCPathForEgnyte,
                "onyx file system\templates\engineering\sealsoriginal")
            If Directory.Exists(sealsPath) Then
                pathsToAdd.Add(sealsPath)
            Else
                Dim sealsFullPath = System.IO.Path.Combine(
                    NetworkUNCPathForEgnyte,
                    "onyx file system\Templates\Engineering\SealsOriginal")
                ed.WriteMessage(vbCrLf & "Active access to seals drive does not exist." & vbCrLf &
                                "Please reach out to IT for access to:" & vbCrLf & vbCrLf & sealsFullPath)
            End If
        End If

        AppendPluginDirectoryPaths(pathsToAdd)

        Dim finalPaths As New List(Of String)()
        For Each candidate In keptPaths
            If PathNormalizer.ListContainsEquivalent(finalPaths, candidate) Then Continue For
            finalPaths.Add(candidate)
        Next
        For Each candidate In pathsToAdd
            If String.IsNullOrWhiteSpace(candidate) OrElse Not System.IO.Path.IsPathRooted(candidate) Then Continue For
            If PathNormalizer.ListContainsEquivalent(finalPaths, candidate) Then Continue For
            finalPaths.Add(candidate)
        Next

        If finalPaths.Count > 0 Then
            Application.SetSystemVariable("SRCHPATH", String.Join(";", finalPaths))
        End If

        ApplyPlotStylePaths()
    End Sub

    Private Sub AppendPluginDirectoryPaths(target As List(Of String))
        Try
            Dim dllDir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            If String.IsNullOrWhiteSpace(dllDir) Then Return

            For Each pathVariant In PathNormalizer.GetEquivalentPaths(dllDir)
                If Not System.IO.Path.IsPathRooted(pathVariant) Then Continue For
                If Not Directory.Exists(pathVariant) Then Continue For
                If PathNormalizer.ListContainsEquivalent(target, pathVariant) Then Continue For
                target.Add(pathVariant)
            Next
        Catch
        End Try
    End Sub

    Private Sub ApplyPlotStylePaths()
        If String.IsNullOrEmpty(NetworkUNCPathForEgnyte) Then Return

        Dim plotStyles = System.IO.Path.Combine(
            NetworkUNCPathForEgnyte,
            "Arcxis\Engineering\Drafting Standards\CAD Plot Styles")
        If String.IsNullOrEmpty(plotStyles) Then Return

        Try
            Dim acadApp = Application.AcadApplication
            acadApp.Preferences.Files.PrinterConfigPath = plotStyles
            acadApp.Preferences.Files.PrinterStyleSheetPath = plotStyles
        Catch
        End Try
    End Sub

    Private Function GetPathsPreserveCase() As List(Of String)
        Dim curTrustPath = CStr(Application.GetSystemVariable("SRCHPATH"))
        If String.IsNullOrWhiteSpace(curTrustPath) Then Return New List(Of String)()

        Dim cleaned As New List(Of String)()
        For Each entry In curTrustPath.Split(";"c)
            Dim trimmed = entry.Trim()
            If trimmed = "" Then Continue For
            If Not cleaned.Any(Function(p) String.Equals(p, trimmed, StringComparison.OrdinalIgnoreCase)) Then
                cleaned.Add(trimmed)
            End If
        Next

        Return cleaned
    End Function
End Module
