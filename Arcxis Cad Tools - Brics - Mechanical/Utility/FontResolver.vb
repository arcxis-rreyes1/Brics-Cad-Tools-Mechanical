Imports PdfSharp.Fonts
Imports System.IO

Public Class SystemFontResolver
    Implements IFontResolver

    Public Function ResolveTypeface(familyName As String, bold As Boolean, italic As Boolean) As FontResolverInfo Implements IFontResolver.ResolveTypeface
        ' Map common font names to actual font files based on style
        Dim fontFile As String = Nothing

        Select Case familyName.ToLowerInvariant()
            Case "arial"
                If bold AndAlso italic Then
                    fontFile = "arialbi.ttf"
                ElseIf bold Then
                    fontFile = "arialbd.ttf"
                ElseIf italic Then
                    fontFile = "ariali.ttf"
                Else
                    fontFile = "arial.ttf"
                End If
            Case "times new roman", "times"
                If bold AndAlso italic Then
                    fontFile = "timesbi.ttf"
                ElseIf bold Then
                    fontFile = "timesbd.ttf"
                ElseIf italic Then
                    fontFile = "timesi.ttf"
                Else
                    fontFile = "times.ttf"
                End If
            Case "courier new", "courier"
                If bold AndAlso italic Then
                    fontFile = "courbi.ttf"
                ElseIf bold Then
                    fontFile = "courbd.ttf"
                ElseIf italic Then
                    fontFile = "couri.ttf"
                Else
                    fontFile = "cour.ttf"
                End If
            Case Else
                ' Fallback to Arial with appropriate style
                If bold AndAlso italic Then
                    fontFile = "arialbi.ttf"
                ElseIf bold Then
                    fontFile = "arialbd.ttf"
                ElseIf italic Then
                    fontFile = "ariali.ttf"
                Else
                    fontFile = "arial.ttf"
                End If
        End Select

        ' Return font info with the file name as the key
        Return New FontResolverInfo(fontFile)
    End Function

    Public Function GetFont(faceName As String) As Byte() Implements IFontResolver.GetFont
        ' Get font file path from Windows Fonts directory
        Dim fontPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), faceName)

        ' Check if file exists, if not try system32\fonts
        If Not File.Exists(fontPath) Then
            fontPath = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot"), "Fonts", faceName)
        End If

        ' Read and return font file bytes
        If File.Exists(fontPath) Then
            Try
                Return File.ReadAllBytes(fontPath)
            Catch
                ' If we can't read the font, return nothing and let PdfSharp handle it
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function
End Class