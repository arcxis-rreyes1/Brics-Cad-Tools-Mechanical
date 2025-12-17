

Public Class Form_TakeOffTag

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button3.Click

        If RadioButton32.Checked = True Then

            Module_TakeOffTag.LabelType = "Tags"

        ElseIf RadioButton33.Checked = True Then

            Module_TakeOffTag.LabelType = "TakeOff"

        End If

        If RadioButton1.Checked = True Then

            Module_TakeOffTag.TagType = "Rafter"

        ElseIf RadioButton2.Checked = True Then

            Module_TakeOffTag.TagType = "CJoist"

        ElseIf RadioButton3.Checked = True Then

            Module_TakeOffTag.TagType = "TJI"

        ElseIf RadioButton4.Checked = True Then

            Module_TakeOffTag.TagType = "BCI"

        ElseIf RadioButton5.Checked = True Then

            Module_TakeOffTag.TagType = "Purlin"

        ElseIf RadioButton6.Checked = True Then

            Module_TakeOffTag.TagType = "Clip"

        End If

        If CheckBox1.Checked = True Then

            Module_TakeOffTag.FenceTag = "Yes"

        Else

            Module_TakeOffTag.FenceTag = "No"

        End If

        If RadioButton7.Checked = True Then

            Module_TakeOffTag.RafterTagLocation = "Left"

        ElseIf RadioButton8.Checked = True Then

            Module_TakeOffTag.RafterTagLocation = "Center"

        ElseIf RadioButton9.Checked = True Then

            Module_TakeOffTag.RafterTagLocation = "Right"

        End If


        If RadioButton10.Checked = True Then

            Module_TakeOffTag.Grade = "#2"

        ElseIf RadioButton11.Checked = True Then

            Module_TakeOffTag.Grade = "#3"

        End If

        If RadioButton12.Checked = True Then

            Module_TakeOffTag.Storage = "No Storage"

        ElseIf RadioButton13.Checked = True Then

            Module_TakeOffTag.Storage = "With Storage"

        End If

        If RadioButton14.Checked = True Then

            Module_TakeOffTag.MemberSize = "2x6"
            Module_TakeOffTag.MemberSizeButton = "2x6"

        ElseIf RadioButton15.Checked = True Then

            Module_TakeOffTag.MemberSize = "2x8"
            Module_TakeOffTag.MemberSizeButton = "2x8"

        ElseIf RadioButton16.Checked = True Then

            Module_TakeOffTag.MemberSize = "2x10"
            Module_TakeOffTag.MemberSizeButton = "2x10"

        ElseIf RadioButton17.Checked = True Then

            Module_TakeOffTag.MemberSize = "2x12"
            Module_TakeOffTag.MemberSizeButton = "2x12"

        ElseIf RadioButton18.Checked = True Then

            Module_TakeOffTag.MemberSize = "Auto"
            Module_TakeOffTag.MemberSizeButton = "Auto"

        End If

        If RadioButton19.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "110"

        ElseIf RadioButton20.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "210"

        ElseIf RadioButton21.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "230"

        ElseIf RadioButton22.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "360"

        ElseIf RadioButton23.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "560"

        ElseIf RadioButton24.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "4500"

        ElseIf RadioButton25.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "5000"

        ElseIf RadioButton26.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "6000"

        ElseIf RadioButton27.Checked = True Then

            Module_TakeOffTag.IJoistSeries = "6500"

        End If

        If RadioButton28.Checked = True Then

            Module_TakeOffTag.IJoistSize = "11 1/4"

        ElseIf RadioButton29.Checked = True Then

            Module_TakeOffTag.IJoistSize = "11 7/8"

        ElseIf RadioButton30.Checked = True Then

            Module_TakeOffTag.IJoistSize = "14"

        ElseIf RadioButton31.Checked = True Then

            Module_TakeOffTag.IJoistSize = "16"

        End If

        Module_TakeOffTag.RafterSlope = ComboBox1.Text
        Module_TakeOffTag.JoistSpace = ComboBox2.Text

        Module_TakeOffTag.TagCancel = "OK"
        Me.Close()



    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        GroupBox2.Enabled = True
        GroupBox3.Enabled = False
        GroupBox4.Enabled = True
        GroupBox5.Enabled = True
        GroupBox6.Enabled = False
        GroupBox7.Enabled = True
        GroupBox8.Enabled = False
        GroupBox9.Enabled = False
        GroupBox10.Enabled = True
        'GroupBox11.Enabled = True
        RadioButton18.Enabled = False

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        GroupBox2.Enabled = False
        GroupBox3.Enabled = True
        GroupBox4.Enabled = False
        GroupBox5.Enabled = True
        GroupBox6.Enabled = True
        GroupBox7.Enabled = True
        GroupBox8.Enabled = False
        GroupBox9.Enabled = False
        GroupBox10.Enabled = True
        'GroupBox11.Enabled = False
        RadioButton18.Enabled = True

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = False
        GroupBox8.Enabled = True
        GroupBox9.Enabled = True
        GroupBox10.Enabled = False
        'GroupBox11.Enabled = False
        RadioButton19.Enabled = True
        RadioButton20.Enabled = True
        RadioButton21.Enabled = True
        RadioButton22.Enabled = True
        RadioButton23.Enabled = True
        RadioButton24.Enabled = False
        RadioButton25.Enabled = False
        RadioButton26.Enabled = False
        RadioButton27.Enabled = False


    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = False
        GroupBox8.Enabled = True
        GroupBox9.Enabled = True
        GroupBox10.Enabled = False
        'GroupBox11.Enabled = False
        RadioButton19.Enabled = False
        RadioButton20.Enabled = False
        RadioButton21.Enabled = False
        RadioButton22.Enabled = False
        RadioButton23.Enabled = False
        RadioButton24.Enabled = True
        RadioButton25.Enabled = True
        RadioButton26.Enabled = True
        RadioButton27.Enabled = True

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = True
        GroupBox8.Enabled = False
        GroupBox9.Enabled = False
        GroupBox10.Enabled = False
        'GroupBox11.Enabled = False
        RadioButton18.Enabled = False


    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged

        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = False
        GroupBox8.Enabled = False
        GroupBox9.Enabled = False
        GroupBox10.Enabled = False
        RadioButton18.Enabled = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Module_TakeOffTag.TagCancel = "Cancel"

        Me.Close()


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then

            GroupBox4.Enabled = False
            'Module_TakeOffTag.FenceTag = "Yes"

        ElseIf CheckBox1.Checked = False Then

            GroupBox4.Enabled = True
            'Module_TakeOffTag.FenceTag = "No"

        End If

        'GroupBox4.Enabled = False

    End Sub

    Private Sub RadioButton32_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton32.CheckedChanged

        RadioButton3.Enabled = False
        RadioButton4.Enabled = False
        RadioButton5.Enabled = False
        RadioButton6.Enabled = False

    End Sub

    Private Sub RadioButton33_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton33.CheckedChanged

        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        RadioButton4.Enabled = True
        RadioButton5.Enabled = True
        RadioButton3.Enabled = True

    End Sub
End Class