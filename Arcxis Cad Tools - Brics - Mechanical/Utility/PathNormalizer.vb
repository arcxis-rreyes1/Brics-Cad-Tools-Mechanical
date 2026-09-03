Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public NotInheritable Class PathNormalizer
    Private Sub New()
    End Sub

    Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (
        ByVal lpszLocalName As String,
        ByVal lpszRemoteName As StringBuilder,
        ByRef cbRemoteName As Integer) As Integer

    Public Shared Function NormalizePath(path As String) As String
        If String.IsNullOrWhiteSpace(path) Then Return Nothing

        Try
            Dim normalized As String = System.IO.Path.GetFullPath(path.Trim())
            If normalized.EndsWith("\") AndAlso normalized.Length > 3 Then
                normalized = normalized.TrimEnd("\"c)
            End If
            Return normalized
        Catch
            Return path.Trim()
        End Try
    End Function

    Public Shared Function TryGetUncPath(localPath As String) As String
        If String.IsNullOrWhiteSpace(localPath) Then Return Nothing

        Dim normalized As String = NormalizePath(localPath)
        If String.IsNullOrWhiteSpace(normalized) Then Return Nothing
        If normalized.StartsWith("\\") Then Return normalized

        Dim root As String = System.IO.Path.GetPathRoot(normalized)
        If String.IsNullOrEmpty(root) OrElse root.Length < 2 Then Return Nothing

        Dim driveLetter As String = root.TrimEnd("\"c)
        If driveLetter.Length < 2 OrElse driveLetter(1) <> ":"c Then Return Nothing

        Try
            Dim unc As New StringBuilder(1024)
            Dim size As Integer = unc.Capacity
            If WNetGetConnection(driveLetter, unc, size) <> 0 OrElse unc.Length = 0 Then
                Return Nothing
            End If

            Dim uncRoot As String = unc.ToString().TrimEnd("\"c)
            Dim relative As String = normalized.Substring(root.Length).TrimStart("\"c)
            If String.IsNullOrEmpty(relative) Then Return uncRoot
            Return NormalizePath(System.IO.Path.Combine(uncRoot, relative))
        Catch
            Return Nothing
        End Try
    End Function

    Public Shared Function TryGetMappedDrivePath(anyPath As String) As String
        If String.IsNullOrWhiteSpace(anyPath) Then Return Nothing

        Dim normalized As String = NormalizePath(anyPath)
        If String.IsNullOrWhiteSpace(normalized) Then Return Nothing

        If normalized.Length >= 2 AndAlso normalized(1) = ":"c AndAlso Not normalized.StartsWith("\\") Then
            Return normalized
        End If

        If Not normalized.StartsWith("\\") Then Return Nothing

        For Each drv In DriveInfo.GetDrives()
            If drv.DriveType <> DriveType.Network Then Continue For

            Try
                Dim letter As String = drv.Name.TrimEnd("\"c)
                Dim unc As New StringBuilder(1024)
                Dim size As Integer = unc.Capacity
                If WNetGetConnection(letter, unc, size) <> 0 OrElse unc.Length = 0 Then Continue For

                Dim uncRoot As String = unc.ToString().TrimEnd("\"c)
                If Not normalized.StartsWith(uncRoot, StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim relative As String = normalized.Substring(uncRoot.Length).TrimStart("\"c)
                If String.IsNullOrEmpty(relative) Then Return NormalizePath(drv.Name)
                Return NormalizePath(System.IO.Path.Combine(drv.Name, relative))
            Catch
            End Try
        Next

        Return Nothing
    End Function

    Public Shared Function GetEquivalentPaths(path As String) As List(Of String)
        Dim results As New List(Of String)()
        Dim normalized As String = NormalizePath(path)
        If String.IsNullOrWhiteSpace(normalized) Then Return results

        AddUnique(results, normalized)

        Dim unc As String = TryGetUncPath(normalized)
        If Not String.IsNullOrWhiteSpace(unc) Then AddUnique(results, unc)

        Dim mapped As String = TryGetMappedDrivePath(normalized)
        If Not String.IsNullOrWhiteSpace(mapped) Then AddUnique(results, mapped)

        Return results
    End Function

    Public Shared Function IsEquivalent(pathA As String, pathB As String) As Boolean
        If String.IsNullOrWhiteSpace(pathA) OrElse String.IsNullOrWhiteSpace(pathB) Then Return False

        ' Build each side once; AddUnique must not call back into IsEquivalent.
        Dim variantsA = GetEquivalentPaths(pathA)
        Dim variantsB = GetEquivalentPaths(pathB)
        For Each variantA In variantsA
            Dim normalizedA = NormalizePath(variantA)
            For Each variantB In variantsB
                If String.Equals(normalizedA, NormalizePath(variantB), StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

    Public Shared Function ListContainsEquivalent(paths As IEnumerable(Of String), candidate As String) As Boolean
        If paths Is Nothing OrElse String.IsNullOrWhiteSpace(candidate) Then Return False

        For Each existing In paths
            If IsEquivalent(existing, candidate) Then Return True
        Next

        Return False
    End Function

    Private Shared Sub AddUnique(target As List(Of String), path As String)
        If String.IsNullOrWhiteSpace(path) Then Return

        Dim normalized = NormalizePath(path)
        If String.IsNullOrWhiteSpace(normalized) Then Return

        For Each existing In target
            If String.Equals(NormalizePath(existing), normalized, StringComparison.OrdinalIgnoreCase) Then Return
        Next

        target.Add(normalized)
    End Sub
End Class
