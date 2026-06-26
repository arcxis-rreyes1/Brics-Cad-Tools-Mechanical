Imports System.IO

Public Class NetworkHelpers
    Public Shared Function IsNetworkPathAccessible(pathToCheck As String) As Boolean
        Try
            If String.IsNullOrEmpty(pathToCheck) Then Return False

            ' Use System.IO.Path to get the root
            Dim root As String = System.IO.Path.GetPathRoot(pathToCheck)
            If String.IsNullOrEmpty(root) Then Return False

            ' Validate root exists and is readable
            If Not Directory.Exists(root) Then Return False
            Directory.GetDirectories(root)

            Return True
        Catch
            Return False
        End Try
    End Function

    Public Shared Function CreateDirectoryWithRetry(pathToCreate As String, Optional maxRetries As Integer = 3) As Boolean
        For attempt As Integer = 1 To maxRetries
            Try
                If Directory.Exists(pathToCreate) Then Return True
                Directory.CreateDirectory(pathToCreate)
                Return True
            Catch ex As IOException When attempt < maxRetries
                Threading.Thread.Sleep(500)
            Catch
                Return False
            End Try
        Next
        Return False
    End Function
End Class