Imports System
Imports System.Collections.Generic
Imports System.Drawing.Printing
Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Bricscad.PlottingServices


Public Class Form_Arcxis_TB2

    ' Per-layout viewport clone tracking and template fingerprints
    Private ReadOnly _layoutViewportData As New Dictionary(Of String, LayoutViewportData)(StringComparer.OrdinalIgnoreCase)

    Private Class LayoutViewportData
        Public ClonedModelViewportIds As HashSet(Of ObjectId)
        Public TemplateFingerprints As List(Of ModelViewportFingerprint)
    End Class

    Private Class ModelViewportFingerprint
        Public CenterX As Double
        Public CenterY As Double
        Public Width As Double
        Public Height As Double
    End Class

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim NumFramePages As Integer = CInt(Val(TextBox1.Text))
        Dim NewLayoutName As String = ComboBox1.Text
        Dim NumElevations As Integer = CInt(Val(TextBox2.Text))
        Dim wallbracepagenum110 As String = TextBox3.Text
        Dim wallbracepagenum100 As String = TextBox6.Text
        Dim wallbracepagenum90 As String = TextBox7.Text
        Dim WSIpagenum As String = TextBox4.Text
        Dim WSIIpagenum As String = TextBox5.Text
        Dim elevtotal As Integer = NumFramePages * NumElevations
        Dim WB110total As Integer = wallbracepagenum110 * NumElevations
        Dim WB100total As Integer = wallbracepagenum100 * NumElevations
        Dim WB90total As Integer = wallbracepagenum90 * NumElevations
        Dim WSItotal As Integer = WSIpagenum * NumElevations
        Dim WSIItotal As Integer = WSIIpagenum * NumElevations
        Dim ExistlayoutName As String = ATB_CurLayoutName
        Dim frm As New Form_Arcxis_TB1
        Dim frm2 As New Form_Arcxis_TB2
        Dim frm4 As New Form_Arcxis_TB4
        ATB_CustomLayoutList.Clear()
        _layoutViewportData.Clear()

        Me.Close()

        If NewLayoutName = "Custom" Then

            For elnum As Integer = 1 To NumElevations

                frm4.Label1.Text = "Enter Custom Layout Letter for Elevation #" & elnum
                frm4.TextBox1.Text = ""
                frm4.ShowDialog()

                ATB_CustomLayoutList.Add(ATB_CustomLayoutLetter)

                If elnum = 1 Then

                    NewLayoutName = "FR-1" & ATB_CustomLayoutLetter

                    RenameFramingLayout(ExistlayoutName, NewLayoutName)

                End If

            Next

            AddLayoutNameCustom(NewLayoutName, NumElevations, NumFramePages, wallbracepagenum110, wallbracepagenum100, wallbracepagenum90, WSIpagenum, WSIIpagenum)

        Else

            If NumFramePages = 0 Then

                If NewLayoutName = "FR-1.1" Then

                    If wallbracepagenum110 <> 0 Then

                        NewLayoutName = "WB-1.1 (110 MPH)"

                    ElseIf wallbracepagenum100 <> 0 Then

                        NewLayoutName = "WB-1.1 (100 MPH)"


                    ElseIf wallbracepagenum90 <> 0 Then

                        NewLayoutName = "WB-1.1 (90 MPH)"

                    End If

                Else

                    If wallbracepagenum110 <> 0 Then

                        NewLayoutName = "WB-1A (110 MPH)"

                    ElseIf wallbracepagenum100 <> 0 Then

                        NewLayoutName = "WB-1A (100 MPH)"

                    ElseIf wallbracepagenum90 <> 0 Then

                        NewLayoutName = "WB-1A (90 MPH)"

                    End If

                End If

                RenameFramingLayout(ExistlayoutName, NewLayoutName)

                AddLayoutNameWallBracingOnly(NewLayoutName, NumElevations, NumFramePages, wallbracepagenum110, wallbracepagenum100, wallbracepagenum90, WSIpagenum, WSIIpagenum)

            Else

                RenameFramingLayout(ExistlayoutName, NewLayoutName)

                AddLayoutName(NewLayoutName, NumElevations, NumFramePages, wallbracepagenum110, wallbracepagenum100, wallbracepagenum90, WSIpagenum, WSIIpagenum)


            End If
        End If

        SetPSLTScale()

        ' BricsCAD lazily injects a default model viewport when a layout is first activated
        ' (e.g. during SetPSLTScale's CurrentLayout switches). Clean those up after activation.
        FinalizeNewLayoutViewports()

        frm2.Close()

        If frm.CheckBox1.Checked = True Then
            For Each Lay1 In frm.ListBox1.Items
                frm.ListBox1.SelectedItems.Add(Lay1)
            Next
        End If


    End Sub

    Private Sub AddLayoutName(NewLayoutName As String, NumElevations As Integer, NumFramePages As Integer, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

        Dim panover As Integer
        Dim pandown As Integer
        Dim NextlayoutName As String
        Dim pgltr As Char
        Dim dcnm As Integer
        Dim elnt As String
        Dim dnum As String

        Dim i As Integer = Asc("A")
        Dim x As Char = Chr(i + 48)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)

                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To NumFramePages

                    panover = 1800 * (pagenum - 1)

                    NextlayoutName = "FR-" & pagenum & pgltr & "-R"
                    dnum = "FR-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To NumFramePages

                    panover = 1800 * (pagenum - 1)

                    NextlayoutName = "FR-" & pagenum & "." & dcnm
                    dnum = "FR-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)

                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To NumFramePages

                    panover = 1800 * (pagenum - 1)

                    NextlayoutName = "FR-" & pagenum & pgltr
                    dnum = "FR-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''FR-1A-R (LC)

        If ATB_CheckState = "YES" Then

            For frvalue As Integer = 0 To NumElevations - 1

                pandown = (NumElevations * 1200) + (1200 * frvalue)

                If NewLayoutName = "FR-1A-R" Then

                    pgltr = Chr(65 + frvalue)

                    elnt = "ELEVATION " & pgltr

                    For pagenum As Integer = 1 To NumFramePages

                        panover = 1800 * (pagenum - 1)
                        NextlayoutName = "FR-" & pagenum & pgltr & "-R (LC)"
                        dnum = "FR-" & pagenum & pgltr

                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                    Next

                ElseIf NewLayoutName = "FR-1.1" Then

                    dcnm = 1 + frvalue
                    elnt = "ELEVATION " & dcnm

                    For pagenum As Integer = 1 To NumFramePages

                        panover = 1800 * (pagenum - 1)
                        NextlayoutName = "FR-" & pagenum & "." & dcnm & " (LC)"
                        dnum = "FR-" & pagenum & "." & dcnm

                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                    Next

                ElseIf NewLayoutName = "FR-1A" Then


                    pgltr = Chr(65 + frvalue)

                    elnt = "ELEVATION " & pgltr

                    For pagenum As Integer = 1 To NumFramePages

                        panover = 1800 * (pagenum - 1)
                        NextlayoutName = "FR-" & pagenum & pgltr & " (LC)"
                        dnum = "FR-" & pagenum & pgltr

                        CreateFramingLayout(NewLayoutName, NextlayoutName)
                        ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                    Next

                End If

            Next

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (110)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum110

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                    NextlayoutName = "WB-" & pagenum & pgltr & "-R (110 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum110

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (110 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum110

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                    NextlayoutName = "WB-" & pagenum & pgltr & " (110 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (100)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum100

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110)
                    NextlayoutName = "WB-" & pagenum & pgltr & "-R (100 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum100

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110)
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (100 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum100

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110)
                    NextlayoutName = "WB-" & pagenum & pgltr & " (100 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (90)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue


            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum90

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100)
                    NextlayoutName = "WB-" & pagenum & pgltr & "-R (90 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum90

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100)
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (90 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum90

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100)
                    NextlayoutName = "WB-" & pagenum & pgltr & " (90 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(I)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & "-R (I)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (I)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (I)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(II)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName = "FR-1A-R" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & "-R (II)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1.1" Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (II)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName = "FR-1A" Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (II)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next


    End Sub

    Private Sub AddLayoutNameWallBracingOnly(NewLayoutName As String, NumElevations As Integer, NumFramePages As Integer, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

        Dim panover As Integer
        Dim pandown As Integer
        Dim NextlayoutName As String
        Dim pgltr As Char
        Dim dcnm As Integer
        Dim elnt As String
        Dim dnum As String

        Dim i As Integer = Asc("A")
        Dim x As Char = Chr(i + 48)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A (110)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName.Contains("WB-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum110

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (110 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("WB-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum110

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                    NextlayoutName = "WB-" & pagenum & pgltr & " (110 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A (100)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName.Contains("WB-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum100

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110)
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (100 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("WB-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum100

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110)
                    NextlayoutName = "WB-" & pagenum & pgltr & " (100 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (90)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName.Contains("WB-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To wallbracepagenum90

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100)
                    NextlayoutName = "WB-" & pagenum & "." & dcnm & " (90 MPH)"
                    dnum = "WB-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("WB-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To wallbracepagenum90

                    panover = (1800 * NumFramePages) + (1800 * (pagenum - 1)) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100)
                    NextlayoutName = "WB-" & pagenum & pgltr & " (90 MPH)"
                    dnum = "WB-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(I)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName.Contains("WB-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (I)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("WB-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (I)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

            If NewLayoutName.Contains("FR-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (I)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("FR-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (I)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(II)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            If NewLayoutName.Contains("WB-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (II)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("WB-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (II)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

            If NewLayoutName.Contains("FR-1.1") Then

                dcnm = 1 + frvalue
                elnt = "ELEVATION " & dcnm

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & "." & dcnm & " (II)"
                    dnum = "W-" & pagenum & "." & dcnm

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            ElseIf NewLayoutName.Contains("FR-1A") Then

                pgltr = Chr(65 + frvalue)
                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To WSIIpagenum

                    panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                    NextlayoutName = "W-" & pagenum & pgltr & " (II)"
                    dnum = "W-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            End If

        Next

    End Sub



    Private Sub AddLayoutNameCustom(NewLayoutName As String, NumElevations As Integer, NumFramePages As Integer, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

        Dim panover As Integer
        Dim pandown As Integer
        Dim NextlayoutName As String
        Dim pgltr As String
        Dim elnt As String
        Dim dnum As String

        For frvalue As Integer = 0 To NumElevations - 1
            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)

            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To NumFramePages

                panover = 1800 * (pagenum - 1)
                NextlayoutName = "FR-" & pagenum & pgltr
                dnum = "FR-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next
        'Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''FR-1A-R (LC)

        If ATB_CheckState = "YES" Then

            For frvalue As Integer = 0 To NumElevations - 1

                pandown = (NumElevations * 1200) + (1200 * frvalue)

                pgltr = ATB_CustomLayoutList(frvalue)

                elnt = "ELEVATION " & pgltr

                For pagenum As Integer = 1 To NumFramePages

                    panover = 1800 * (pagenum - 1)
                    NextlayoutName = "FR-" & pagenum & pgltr & " (LC)"
                    dnum = "FR-" & pagenum & pgltr

                    CreateFramingLayout(NewLayoutName, NextlayoutName)
                    ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

                Next

            Next

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (110)
        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)
            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To wallbracepagenum110

                panover = (1800 * NumFramePages) + (1800 * (pagenum - 1))
                NextlayoutName = "WB-" & pagenum & pgltr & " (110 MPH)"
                dnum = "WB-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (100)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)
            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To wallbracepagenum100

                panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * (pagenum - 1))
                NextlayoutName = "WB-" & pagenum & pgltr & " (100 MPH)"
                dnum = "WB-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''WB-1A-R (90)

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)
            elnt = "ELEVATION " & pgltr
            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To wallbracepagenum90

                panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * (pagenum - 1))
                NextlayoutName = "WB-" & pagenum & pgltr & " (90 MPH)"
                dnum = "WB-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(I)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)
            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To WSIpagenum

                panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + (1800 * wallbracepagenum90) + (1800 * (pagenum - 1))
                NextlayoutName = "W-" & pagenum & pgltr & " (I)"
                dnum = "W-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''(II)W-1A-R

        For frvalue As Integer = 0 To NumElevations - 1

            pandown = 1200 * frvalue

            pgltr = ATB_CustomLayoutList(frvalue)
            elnt = "ELEVATION " & pgltr

            For pagenum As Integer = 1 To WSIIpagenum

                panover = (1800 * NumFramePages) + (1800 * wallbracepagenum110) + (1800 * wallbracepagenum100) + +(1800 * wallbracepagenum90) + (1800 * WSIpagenum) + (1800 * (pagenum - 1))
                NextlayoutName = "W-" & pagenum & pgltr & " (II)"
                dnum = "W-" & pagenum & pgltr

                CreateFramingLayout(NewLayoutName, NextlayoutName)
                ModifyViewPortCenter(NextlayoutName, panover, pandown, elnt, dnum)

            Next
        Next

    End Sub

    Private Shared Function IsOverallViewport(vp As Viewport) As Boolean
        If vp Is Nothing Then Return False
        Return vp.Number = 1
    End Function

    Private Shared Function CaptureSourceModelViewportIds(sourceBtr As BlockTableRecord, acTrans As Transaction) As List(Of ObjectId)
        Dim ids As New List(Of ObjectId)
        For Each entId As ObjectId In sourceBtr
            Dim sourceVp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If sourceVp Is Nothing Then Continue For
            If IsOverallViewport(sourceVp) OrElse IsKnownBricsCadDefaultViewport(sourceVp) Then Continue For
            ids.Add(entId)
        Next
        Return ids
    End Function

    Private Shared Function CaptureTemplateModelViewportFingerprints(sourceBtr As BlockTableRecord, acTrans As Transaction) As List(Of ModelViewportFingerprint)
        Dim fingerprints As New List(Of ModelViewportFingerprint)
        For Each entId As ObjectId In sourceBtr
            Dim sourceVp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If sourceVp Is Nothing Then Continue For
            If IsOverallViewport(sourceVp) OrElse IsKnownBricsCadDefaultViewport(sourceVp) Then Continue For
            fingerprints.Add(New ModelViewportFingerprint With {
                .CenterX = sourceVp.CenterPoint.X,
                .CenterY = sourceVp.CenterPoint.Y,
                .Width = sourceVp.Width,
                .Height = sourceVp.Height
            })
        Next
        Return fingerprints
    End Function

    Private Shared Function MatchesTemplateModelViewport(vp As Viewport, fingerprints As List(Of ModelViewportFingerprint)) As Boolean
        If vp Is Nothing OrElse fingerprints Is Nothing OrElse fingerprints.Count = 0 Then Return False
        Const tol As Double = 0.15
        For Each fp In fingerprints
            If Math.Abs(vp.CenterPoint.X - fp.CenterX) <= tol AndAlso
               Math.Abs(vp.CenterPoint.Y - fp.CenterY) <= tol AndAlso
               Math.Abs(vp.Width - fp.Width) <= tol AndAlso
               Math.Abs(vp.Height - fp.Height) <= tol Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Function ShouldKeepModelViewport(vp As Viewport, entId As ObjectId, clonedModelViewportIds As HashSet(Of ObjectId), templateFingerprints As List(Of ModelViewportFingerprint)) As Boolean
        If vp Is Nothing OrElse IsOverallViewport(vp) Then Return False
        If IsKnownBricsCadDefaultViewport(vp) Then Return False
        If clonedModelViewportIds IsNot Nothing AndAlso clonedModelViewportIds.Contains(entId) Then Return True
        Return MatchesTemplateModelViewport(vp, templateFingerprints)
    End Function

    Private Shared Sub SafeEraseViewport(vp As Viewport)
        If vp Is Nothing OrElse vp.IsErased OrElse IsOverallViewport(vp) Then Return
        If vp.Locked Then vp.Locked = False
        Try
            vp.Erase()
        Catch ex As Teigha.Runtime.Exception
            Try
                vp.On = False
                vp.Erase()
            Catch
            End Try
        End Try
    End Sub

    Private Shared Sub RemoveAllModelViewports(btr As BlockTableRecord, acTrans As Transaction)
        For Each entId As ObjectId In btr
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForWrite), Viewport)
            If vp Is Nothing OrElse vp.IsErased OrElse IsOverallViewport(vp) Then Continue For
            SafeEraseViewport(vp)
        Next
    End Sub

    Private Shared Function IsKnownBricsCadDefaultViewport(vp As Viewport) As Boolean
        If vp Is Nothing OrElse IsOverallViewport(vp) Then Return False
        Const centerTol As Double = 0.25
        Const sizeTol As Double = 0.35

        ' BricsCAD-injected small model viewport at ~5.25,4 — do NOT flag large template VPs at same center
        If Math.Abs(vp.CenterPoint.X - 5.25) <= centerTol AndAlso Math.Abs(vp.CenterPoint.Y - 4.0) <= centerTol Then
            Dim minDim As Double = Math.Min(vp.Width, vp.Height)
            Dim maxDim As Double = Math.Max(vp.Width, vp.Height)
            If Math.Abs(minDim - 6.375) <= sizeTol AndAlso Math.Abs(maxDim - 8.375) <= sizeTol Then Return True
            If Math.Abs(minDim - 6.4) <= sizeTol AndAlso Math.Abs(maxDim - 8.4) <= sizeTol Then Return True
            If String.Equals(vp.Layer, "S-ANNO-AUTOMATION", StringComparison.OrdinalIgnoreCase) AndAlso maxDim < 12.0 AndAlso minDim < 10.0 Then Return True
        End If

        Return False
    End Function

    Private Sub StoreLayoutViewportData(layoutName As String, clonedIds As HashSet(Of ObjectId), fingerprints As List(Of ModelViewportFingerprint))
        _layoutViewportData(layoutName) = New LayoutViewportData With {
            .ClonedModelViewportIds = New HashSet(Of ObjectId)(clonedIds),
            .TemplateFingerprints = fingerprints
        }
    End Sub

    Private Function GetLayoutViewportData(layoutName As String) As LayoutViewportData
        Dim data As LayoutViewportData = Nothing
        _layoutViewportData.TryGetValue(layoutName, data)
        Return data
    End Function

    ' DeepClone IdMapping is the authoritative clone list: only mapped ObjectIds came from the template
    Private Shared Function BuildClonedViewportKeepSet(idmap As IdMapping, sourceViewportIds As List(Of ObjectId), acTrans As Transaction) As HashSet(Of ObjectId)
        Dim kept As New HashSet(Of ObjectId)

        For Each sourceId As ObjectId In sourceViewportIds
            If Not idmap.Contains(sourceId) Then Continue For
            Dim pair As IdPair = idmap(sourceId)
            If pair.IsCloned AndAlso Not pair.Value.IsNull Then kept.Add(pair.Value)
        Next

        For Each pair As IdPair In idmap
            If Not pair.IsCloned OrElse pair.Value.IsNull OrElse kept.Contains(pair.Value) Then Continue For
            Dim sourceVp = TryCast(acTrans.GetObject(pair.Key, OpenMode.ForRead), Viewport)
            If sourceVp Is Nothing OrElse IsKnownBricsCadDefaultViewport(sourceVp) Then Continue For
            kept.Add(pair.Value)
        Next

        Return kept
    End Function

    Private Shared Sub PrepareLayoutViewports(acDoc As Document, acCurDb As Database, layoutName As String, layoutData As LayoutViewportData)
        If layoutData Is Nothing Then Return

        Dim lm As LayoutManager = LayoutManager.Current
        Dim previousLayout As String = lm.CurrentLayout

        Try
            lm.CurrentLayout = layoutName
        Catch
            Return
        End Try

        Try
            acDoc.Editor.Regen()
        Catch
        End Try

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                If layouts.Contains(layoutName) Then
                    Dim layId As ObjectId = layouts.GetAt(layoutName)
                    Dim lay As Layout = TryCast(acTrans.GetObject(layId, OpenMode.ForRead), Layout)
                    If lay IsNot Nothing Then
                        Dim btr As BlockTableRecord = TryCast(acTrans.GetObject(lay.BlockTableRecordId, OpenMode.ForWrite), BlockTableRecord)
                        If btr IsNot Nothing Then
                            SyncLayoutModelViewports(btr, acTrans, layoutData.ClonedModelViewportIds, layoutData.TemplateFingerprints)
                        End If
                    End If
                End If
                acTrans.Commit()
            End Using
        End Using

        Try
            If Not String.IsNullOrEmpty(previousLayout) AndAlso previousLayout <> layoutName Then
                lm.CurrentLayout = previousLayout
            End If
        Catch
        End Try
    End Sub

    Private Sub FinalizeNewLayoutViewports()
        If _layoutViewportData.Count = 0 Then Return

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        For Each kvp In _layoutViewportData
            PrepareLayoutViewports(acDoc, acCurDb, kvp.Key, kvp.Value)
        Next
    End Sub

    Private Shared Function HasOverallViewport(btr As BlockTableRecord, acTrans As Transaction) As Boolean
        For Each entId As ObjectId In btr
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If vp IsNot Nothing AndAlso Not vp.IsErased AndAlso IsOverallViewport(vp) Then Return True
        Next
        Return False
    End Function

    Private Shared Sub SyncLayoutModelViewports(btr As BlockTableRecord, acTrans As Transaction, clonedModelViewportIds As HashSet(Of ObjectId), templateFingerprints As List(Of ModelViewportFingerprint))
        If Not HasOverallViewport(btr, acTrans) Then Return

        Dim modelVpIds As New List(Of ObjectId)
        For Each entId As ObjectId In btr
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If vp Is Nothing OrElse vp.IsErased OrElse IsOverallViewport(vp) Then Continue For
            modelVpIds.Add(entId)
        Next

        Dim keepIds As New HashSet(Of ObjectId)
        For Each entId As ObjectId In modelVpIds
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If vp IsNot Nothing AndAlso ShouldKeepModelViewport(vp, entId, clonedModelViewportIds, templateFingerprints) Then
                keepIds.Add(entId)
            End If
        Next

        ' Safety: never wipe every model viewport when the template had at least one
        If keepIds.Count = 0 AndAlso modelVpIds.Count > 0 AndAlso templateFingerprints IsNot Nothing AndAlso templateFingerprints.Count > 0 Then
            For Each entId As ObjectId In modelVpIds
                Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
                If vp IsNot Nothing AndAlso Not IsKnownBricsCadDefaultViewport(vp) AndAlso MatchesTemplateModelViewport(vp, templateFingerprints) Then
                    keepIds.Add(entId)
                End If
            Next
        End If

        For Each entId As ObjectId In modelVpIds
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForWrite), Viewport)
            If vp Is Nothing Then Continue For
            If keepIds.Contains(entId) Then
                vp.On = True
            Else
                SafeEraseViewport(vp)
            End If
        Next
    End Sub

    Private Shared Sub ClearLayoutBlockExceptOverall(btr As BlockTableRecord, acTrans As Transaction)
        Dim toErase As New List(Of ObjectId)
        For Each entId As ObjectId In btr
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If vp IsNot Nothing AndAlso IsOverallViewport(vp) Then Continue For
            toErase.Add(entId)
        Next
        For Each entId As ObjectId In toErase
            Dim obj = acTrans.GetObject(entId, OpenMode.ForWrite)
            If obj Is Nothing OrElse obj.IsErased Then Continue For
            Dim vp = TryCast(obj, Viewport)
            If vp IsNot Nothing Then
                SafeEraseViewport(vp)
            Else
                obj.Erase()
            End If
        Next
    End Sub

    Private Sub RegisterExistingLayoutViewports(layoutName As String, targetBtr As BlockTableRecord, acTrans As Transaction, templateFingerprints As List(Of ModelViewportFingerprint))
        Dim keepIds As New HashSet(Of ObjectId)
        For Each entId As ObjectId In targetBtr
            Dim vp = TryCast(acTrans.GetObject(entId, OpenMode.ForRead), Viewport)
            If vp Is Nothing Then Continue For
            If ShouldKeepModelViewport(vp, entId, Nothing, templateFingerprints) Then keepIds.Add(entId)
        Next
        StoreLayoutViewportData(layoutName, keepIds, templateFingerprints)
    End Sub

    Private Sub CloneTemplateViewports(sourceBtr As BlockTableRecord, targetBtr As BlockTableRecord, acTrans As Transaction, layoutName As String, templateFingerprints As List(Of ModelViewportFingerprint))
        Dim sourceModelViewportIds As List(Of ObjectId) = CaptureSourceModelViewportIds(sourceBtr, acTrans)
        RemoveAllModelViewports(targetBtr, acTrans)

        Dim viewportIdmap As New IdMapping()
        Dim modelVpCol As New ObjectIdCollection()
        For Each entId As ObjectId In sourceModelViewportIds
            modelVpCol.Add(entId)
        Next
        If modelVpCol.Count > 0 Then
            sourceBtr.Database.DeepCloneObjects(modelVpCol, targetBtr.ObjectId, viewportIdmap, False)
        End If

        Dim clonedModelViewportIds As HashSet(Of ObjectId) = BuildClonedViewportKeepSet(viewportIdmap, sourceModelViewportIds, acTrans)
        StoreLayoutViewportData(layoutName, clonedModelViewportIds, templateFingerprints)
    End Sub

    Private Sub CreateFramingLayout(NewLayoutName As String, NextlayoutName As String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim needsPostRegenSync As Boolean = False

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

                If Not layouts.Contains(NewLayoutName) Then
                    acDoc.Editor.WriteMessage(vbLf & "Template layout '" & NewLayoutName & "' was not found.")
                    acTrans.Commit()
                    Return
                End If

                Dim isRenamedTemplateLayout As Boolean = String.Equals(NextlayoutName, NewLayoutName, StringComparison.OrdinalIgnoreCase)
                If layouts.Contains(NextlayoutName) AndAlso Not isRenamedTemplateLayout Then
                    acDoc.Editor.WriteMessage(vbLf & "Layout '" & NextlayoutName & "' already exists; skipping.")
                    acTrans.Commit()
                    Return
                End If

                Dim sourceLayId As ObjectId = layouts.GetAt(NewLayoutName)
                Dim layex As Layout = TryCast(acTrans.GetObject(sourceLayId, OpenMode.ForRead), Layout)
                Dim sourceBtr As BlockTableRecord = acTrans.GetObject(layex.BlockTableRecordId, OpenMode.ForRead)
                Dim templateFingerprints As List(Of ModelViewportFingerprint) = CaptureTemplateModelViewportFingerprints(sourceBtr, acTrans)

                If Application.GetSystemVariable("LAYOUTREGENCTL") <> 0 Then
                    Application.SetSystemVariable("LAYOUTREGENCTL", 0)
                End If

                If isRenamedTemplateLayout Then
                    Dim existingLayoutBtr As BlockTableRecord = acTrans.GetObject(sourceBtr.ObjectId, OpenMode.ForWrite)
                    RegisterExistingLayoutViewports(NextlayoutName, existingLayoutBtr, acTrans, templateFingerprints)
                Else
                    Dim geometryCol As New ObjectIdCollection()
                    For Each id As ObjectId In sourceBtr
                        Dim sourceVp = TryCast(acTrans.GetObject(id, OpenMode.ForRead), Viewport)
                        If sourceVp IsNot Nothing Then Continue For
                        geometryCol.Add(id)
                    Next

                    Dim lm As LayoutManager = LayoutManager.Current
                    Dim targetLayId As ObjectId = lm.CreateLayout(NextlayoutName)
                    Dim targetLay As Layout = TryCast(acTrans.GetObject(targetLayId, OpenMode.ForWrite), Layout)
                    targetLay.CopyFrom(layex)

                    Dim targetBtrId As ObjectId = targetLay.BlockTableRecordId
                    Dim targetBtr As BlockTableRecord = acTrans.GetObject(targetBtrId, OpenMode.ForWrite)
                    ClearLayoutBlockExceptOverall(targetBtr, acTrans)

                    Dim geometryIdmap As New IdMapping()
                    If geometryCol.Count > 0 Then
                        acCurDb.DeepCloneObjects(geometryCol, targetBtrId, geometryIdmap, False)
                    End If

                    CloneTemplateViewports(sourceBtr, targetBtr, acTrans, NextlayoutName, templateFingerprints)
                End If

                needsPostRegenSync = True
                acTrans.Commit()
            End Using
        End Using

        If needsPostRegenSync Then
            PrepareLayoutViewports(acDoc, acCurDb, NextlayoutName, GetLayoutViewportData(NextlayoutName))
        End If

    End Sub


    Private Sub ModifyViewPortCenter(NextlayoutName As String, panover As Integer, pandown As Integer, elnt As String, dnum As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed = acDoc.Editor

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Create a transaction for the external drawing
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)


                ' Check to see if the layout exists in the external drawing
                If layouts.Contains(NextlayoutName) = True Then

                    Dim Layid As ObjectId = layouts.GetAt(NextlayoutName)
                    Dim lay As Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForRead), Layout)

                    '' Open the Block table for read
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                    '' Open the Block table record Paper space for write
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                    Dim blkBlkRec As BlockTableRecord = acTrans.GetObject(lay.BlockTableRecordId, OpenMode.ForWrite)

                    Dim layoutData As LayoutViewportData = GetLayoutViewportData(NextlayoutName)
                    Dim clonedViewportIds As HashSet(Of ObjectId) = If(layoutData IsNot Nothing, layoutData.ClonedModelViewportIds, Nothing)
                    Dim templateFingerprints As List(Of ModelViewportFingerprint) = If(layoutData IsNot Nothing, layoutData.TemplateFingerprints, Nothing)
                    SyncLayoutModelViewports(blkBlkRec, acTrans, clonedViewportIds, templateFingerprints)

                    Dim vpIds As ObjectIdCollection = New ObjectIdCollection()

                    For Each objID As ObjectId In blkBlkRec

                        If (objID.ObjectClass.DxfName.ToUpper = "VIEWPORT") Then

                            Dim layoutVp = TryCast(acTrans.GetObject(objID, OpenMode.ForRead), Viewport)
                            If layoutVp Is Nothing Then Continue For
                            If Not ShouldKeepModelViewport(layoutVp, objID, clonedViewportIds, templateFingerprints) Then Continue For
                            vpIds.Add(objID)

                        ElseIf (objID.ObjectClass.DxfName.ToUpper = "INSERT") Then

                            ' Open the block reference
                            Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(objID, OpenMode.ForRead), BlockReference)

                            Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

                            Dim RevvblockName As String = RevTblRec.Name

                            If RevvblockName.Contains("Arcxis Title Block") Then

                                ' Iterate the attribute collection
                                For Each attId As ObjectId In RevBlockRef.AttributeCollection

                                    ' Open the attribute reference
                                    Dim attref As AttributeReference = DirectCast(acTrans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                                    Dim tagvalue As String = attref.Tag
                                    Dim textvalue As String = attref.TextString


                                    If tagvalue.Contains("ELEVATION") Then

                                        If textvalue.Contains("-") Then

                                            Dim elstr As String = attref.TextString
                                            Dim elstrsplit As String() = elstr.Split("-"c)
                                            If elstrsplit.Length > 1 Then
                                                Dim eldes As String = elstrsplit(1)
                                                Dim nelstr As String = elnt & " -" & eldes
                                                attref.TextString = nelstr
                                            Else
                                                attref.TextString = elnt & " - RIGHT SWING"
                                            End If

                                        Else

                                            attref.TextString = elnt & " - RIGHT SWING"

                                        End If

                                    ElseIf tagvalue.Contains("FR-1") Then

                                        attref.TextString = dnum

                                    ElseIf tagvalue.Contains("PLAN") Then

                                        Dim PLN As String = attref.TextString

                                        If NextlayoutName.Contains("(LC)") Then

                                            attref.TextString = PLN & " – LEAGUE CITY"

                                        End If

                                        If dnum.Contains("WB") Then

                                            If PLN.Contains("FRAMING PLAN") Then

                                                Dim NEWPLN As String = PLN.Replace("FRAMING PLAN", "WALL BRACING PLAN")
                                                attref.TextString = NEWPLN

                                            Else

                                                Dim NEWPLN As String = PLN.Replace("PLAN", "WALL BRACING PLAN")
                                                attref.TextString = NEWPLN

                                            End If

                                        ElseIf dnum.Contains("W-") Then

                                            If PLN.Contains("FRAMING PLAN") Then

                                                Dim NEWPLN As String = PLN.Replace("FRAMING PLAN", "PLAN")
                                                attref.TextString = NEWPLN

                                            End If

                                        End If

                                    ElseIf tagvalue.Contains("OPTIONAL") Then

                                        If NextlayoutName.Contains("(I)") Then

                                            ' attref.TextString = "WINDSTORM PLAN - INLAND I - 120 MPH 3-SECOND GUST - EXP. C"
                                            'attref.TextString = "WINDSTORM PLAN - 155 MPH ULTIMATE 3-SECOND GUST - EXP. C"
                                            attref.TextString = "WINDSTORM PLAN - 150 MPH ULTIMATE 3-SECOND GUST - EXP. C"

                                        ElseIf NextlayoutName.Contains("(II)") Then

                                            'attref.TextString = "WINDSTORM PLAN - INLAND II - 110 MPH 3-SECOND GUST - EXP. C"
                                            'attref.TextString = "WINDSTORM PLAN - 142 MPH ULTIMATE 3-SECOND GUST - EXP. C"
                                            attref.TextString = "WINDSTORM PLAN - 143 MPH ULTIMATE 3-SECOND GUST - EXP. C"

                                        End If

                                    End If

                                Next

                            End If

                        End If

                    Next

                    For Each vpId As ObjectId In vpIds

                        Dim layoutviewport = TryCast(acTrans.GetObject(vpId, OpenMode.ForWrite), Viewport)

                        If layoutviewport IsNot Nothing Then

                            Dim wasLocked As Boolean = layoutviewport.Locked
                            layoutviewport.Locked = False

                            Dim vcp As Point2d = layoutviewport.ViewCenter
                            Dim cp As Point3d = layoutviewport.CenterPoint
                            Dim wvp As Double = layoutviewport.Width
                            Dim hvp As Double = layoutviewport.Height

                            layoutviewport.ViewCenter = New Point2d(vcp.X + panover, vcp.Y - pandown)
                            layoutviewport.CenterPoint = New Point3d(cp.X, cp.Y, cp.Z)
                            layoutviewport.Width = wvp
                            layoutviewport.Height = hvp
                            layoutviewport.On = True

                            layoutviewport.Locked = wasLocked

                        End If

                    Next

                    SyncLayoutModelViewports(blkBlkRec, acTrans, clonedViewportIds, templateFingerprints)

                Else
                    ' Display a message if the layout could not be found in the specified drawing
                    acDoc.Editor.WriteMessage(vbLf & "Layout '" & NextlayoutName & "' could not be imported '" & "'.")
                End If

                ' Discard the changes made to the external drawing file
                acTrans.Commit()
            End Using
        End Using

    End Sub

    Private Sub RenameFramingLayout(ExistlayoutName As String, NewlayoutName As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Create a transaction for the external drawing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim layouts As DBDictionary = TryCast(acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

            ' Check to see if the layout exists in the external drawing
            If layouts.Contains(NewlayoutName) = False Then

                Dim Layid As ObjectId = layouts.GetAt(ExistlayoutName)
                Dim lay As Layout = TryCast(acTrans.GetObject(Layid, OpenMode.ForWrite), Layout)

                lay.LayoutName = NewlayoutName

            End If

            ' Discard the changes made to the external drawing file
            acTrans.Commit()

        End Using

    End Sub

    Private Sub SetPSLTScale()

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                Dim layDict As DBDictionary = acCurDb.LayoutDictionaryId.GetObject(OpenMode.ForRead)
                For Each entry As DBDictionaryEntry In layDict
                    Dim lay As Layout = CType(entry.Value.GetObject(OpenMode.ForRead), Layout)
                    layAndTab.Add(lay.TabOrder, lay.LayoutName)
                Next

                Dim curtab As String = acLayoutMgr.CurrentLayout

                For Each layStr In layAndTab.Values
                    'ed.WriteMessage(v & vbCrLf)

                    If layStr <> "Model" Then


                        acLayoutMgr.CurrentLayout = layStr
                        Application.SetSystemVariable("Psltscale", 0)

                        Dim bt = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)


                        If bt.Has("Arcxis Title Block") Then

                            Dim blkname As String = "Arcxis Title Block"
                            GetTB(blkname)

                        End If

                        If bt.Has("Arcxis Title Block 24x36") Then

                            Dim blkname As String = "Arcxis Title Block 24x36"
                            GetTB(blkname)

                        End If

                    End If

                Next

                acLayoutMgr.CurrentLayout = curtab

                ' Abort the changes to the database
                acTrans.Commit()

            End Using
        End Using

    End Sub

    Private Sub GetTB(blkname As String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

        ' Get the layout dictionary of the current database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acLayoutMgr As LayoutManager = LayoutManager.Current

            Dim curtab As String = acLayoutMgr.CurrentLayout

            Dim acTypValAr() As TypedValue = {New TypedValue(DxfCode.LayoutName, curtab)}
            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
            Dim prSelRes As PromptSelectionResult = acDoc.Editor.SelectAll(acSelFtr)

            Dim Rev1List As New ArrayList



            If prSelRes.Status = PromptStatus.OK Then

                Dim SS As SelectionSet = prSelRes.Value
                Dim sscount As Integer = SS.Count

                If SS IsNot Nothing Then

                    For Each brId As ObjectId In prSelRes.Value.GetObjectIds()
                        If brId.ObjectClass.Name = "AcDbBlockReference" Then

                            Dim strutid As String = brId.ToString

                            ' Open the block reference
                            Dim TBBlockRef As BlockReference = DirectCast(acTrans.GetObject(brId, OpenMode.ForRead), BlockReference)

                            Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                            Dim TBName As String = TBTblRec.Name

                            Dim layidd As String = TBTblRec.LayoutId.ToString

                            Dim BLKEntity As Entity = acTrans.GetObject(brId, OpenMode.ForRead)

                            If TBTblRec.Name = "Arcxis Title Block" Then
                                ZoomObjects(BLKEntity)
                            End If

                        End If

                    Next

                End If

            End If

            ' Abort the changes to the database
            acTrans.Commit()

        End Using

    End Sub

    Private Sub ZoomObjects(BLKEntity As Entity)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Using view As ViewTableRecord = ed.GetCurrentView()
                Dim WCS2DCS As Matrix3d = Matrix3d.PlaneToWorld(view.ViewDirection)
                WCS2DCS = Matrix3d.Displacement(view.Target - Point3d.Origin) * WCS2DCS
                WCS2DCS = Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * WCS2DCS
                WCS2DCS = WCS2DCS.Inverse()
                Dim ext As Extents3d = BLKEntity.GeometricExtents
                Dim tmp As Extents3d = BLKEntity.GeometricExtents
                ext.AddExtents(tmp)
                ext.TransformBy(WCS2DCS)
                view.Width = ext.MaxPoint.X - ext.MinPoint.X
                view.Height = ext.MaxPoint.Y - ext.MinPoint.Y
                view.CenterPoint = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2.0, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0)
                ed.SetCurrentView(view)
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked Then

            Module_Arcxis_TB.ATB_CheckState = "YES"

        Else

            Module_Arcxis_TB.ATB_CheckState = "No"

        End If

    End Sub


End Class