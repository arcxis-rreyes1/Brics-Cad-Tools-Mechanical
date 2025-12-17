Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.PlottingServices
Imports System.Drawing.Printing
Imports System.IO
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.GraphicsInterface


Public Class Form_Arcxis_TB2

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim NumFramePages As Integer = TextBox1.Text
        Dim NewLayoutName As String = ComboBox1.Text
        Dim NumElevations As String = TextBox2.Text
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

        frm2.Close()

        If frm.CheckBox1.Checked = True Then
            For Each Lay1 In frm.ListBox1.Items
                frm.ListBox1.SelectedItems.Add(Lay1)
            Next
        End If


    End Sub

    Private Sub AddLayoutName(NewLayoutName As String, NumElevations As String, NumFramePages As String, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

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

    Private Sub AddLayoutNameWallBracingOnly(NewLayoutName As String, NumElevations As String, NumFramePages As String, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

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



    Private Sub AddLayoutNameCustom(NewLayoutName As String, NumElevations As String, NumFramePages As String, wallbracepagenum110 As String, wallbracepagenum100 As String, wallbracepagenum90 As String, WSIpagenum As String, WSIIpagenum As String)

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

    Private Sub CreateFramingLayout(NewLayoutName As String, NextlayoutName As String)

        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Get the current document and database
        Dim acDocex As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDbex As Database = acDoc.Database

        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            ' Create a transaction for the external drawing
            Using acTransEx As Transaction = acCurDbex.TransactionManager.StartTransaction()

                Dim layoutsEx As DBDictionary = TryCast(acTransEx.GetObject(acCurDbex.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)

                ' Check to see if the layout exists in the external drawing
                If layoutsEx.Contains(NextlayoutName) = False Then

                    Dim Layid As ObjectId = layoutsEx.GetAt(NewLayoutName)
                    Dim layex As Layout = TryCast(acTransEx.GetObject(Layid, OpenMode.ForRead), Layout)

                    'layex.LayoutName = "FR-1A-R"

                    Dim blkBlkRecEx As BlockTableRecord = acTransEx.GetObject(layex.BlockTableRecordId, OpenMode.ForRead)

                    ' Get the objects from the block associated with the layout
                    Dim idCol As ObjectIdCollection = New ObjectIdCollection()
                    For Each id As ObjectId In blkBlkRecEx

                        idCol.Add(id)

                    Next
                    ' Create a transaction for the current drawing
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                        ' Get the block table and create a new block then copy the objects between drawings
                        Dim blkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)


                        Using blkBlkRec As New BlockTableRecord
                            blkBlkRec.Name = "*Paper_Space" & CStr(layoutsEx.Count() - 1)

                            blkTbl.Add(blkBlkRec)
                            acTrans.AddNewlyCreatedDBObject(blkBlkRec, True)
                            ' acCurDbex.WblockCloneObjects(idCol, blkBlkRec.ObjectId, New IdMapping(), DuplicateRecordCloning.Ignore, False)
                            Dim idmap As New IdMapping
                            acCurDb.DeepCloneObjects(idCol, blkBlkRec.ObjectId, idmap, False)


                            ' Create a new layout and then copy properties between drawings
                            Dim layouts As DBDictionary = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForWrite)

                            Using lay As New Layout
                                lay.LayoutName = NextlayoutName
                                lay.AddToLayoutDictionary(acCurDb, blkBlkRec.ObjectId)
                                acTrans.AddNewlyCreatedDBObject(lay, True)
                                lay.CopyFrom(layex)

                            End Using

                        End Using

                        ' Regen the drawing to get the layout tab to display
                        acDoc.Editor.Regen()

                        ' Save the changes made
                        acTrans.Commit()

                    End Using
                Else
                    ' Display a message if the layout could not be found in the specified drawing
                    acDoc.Editor.WriteMessage(vbLf & "Layout '" & NewLayoutName & "' could not be imported '" & "'.")
                End If

                ' Discard the changes made to the external drawing file
                acTransEx.Commit()
            End Using
        End Using


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

                    Dim blkBlkRec As BlockTableRecord = acTrans.GetObject(lay.BlockTableRecordId, OpenMode.ForRead)

                    Dim vpIds As ObjectIdCollection = New ObjectIdCollection()

                    For Each objID As ObjectId In blkBlkRec


                        If (objID.ObjectClass.DxfName.ToUpper = "VIEWPORT") Then

                            vpIds.Add(objID)

                        ElseIf (objID.ObjectClass.DxfName.ToUpper = "INSERT") Then

                            ' Open the block reference
                            Dim RevBlockRef As BlockReference = DirectCast(acTrans.GetObject(objID, OpenMode.ForRead), BlockReference)

                            Dim RevTblRec As BlockTableRecord = TryCast(acTrans.GetObject(RevBlockRef.DynamicBlockTableRecord, OpenMode.ForWrite), BlockTableRecord)

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
                                            Dim elstrsplit As String() = elstr.Split("-")
                                            Dim eldes As String = elstrsplit(1)
                                            Dim nelstr As String = elnt & " -" & eldes
                                            attref.TextString = nelstr

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

                        Dim layoutviewport = TryCast(acTrans.GetObject(vpId, OpenMode.ForWrite), Autodesk.AutoCAD.DatabaseServices.Viewport)

                        If layoutviewport IsNot Nothing Then

                            Dim vcp As Point2d = layoutviewport.ViewCenter
                            Dim vtp As Point3d = layoutviewport.ViewTarget
                            Dim cp As Point3d = layoutviewport.CenterPoint
                            Dim wvp As Double = layoutviewport.Width
                            Dim hvp As Double = layoutviewport.Height

                            layoutviewport.ViewCenter = New Point2d(vcp.X + panover, vcp.Y - pandown)
                            layoutviewport.CenterPoint = New Point3d(cp.X, cp.Y, cp.Z)
                            layoutviewport.Width = wvp
                            layoutviewport.Height = hvp

                        End If

                    Next

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

                            Dim TBTblRec As BlockTableRecord = TryCast(acTrans.GetObject(TBBlockRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)

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