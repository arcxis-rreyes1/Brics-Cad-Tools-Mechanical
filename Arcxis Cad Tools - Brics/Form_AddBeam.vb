Public Class Form_AddBeam

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If RadioButton1.Checked = True Then

            Module_AddBeams.HoriBeamSpaceOffset = "120"

        ElseIf RadioButton2.Checked = True Then

            Module_AddBeams.HoriBeamSpaceOffset = "144"

        ElseIf RadioButton3.Checked = True Then

            Module_AddBeams.HoriBeamSpaceOffset = "168"

        ElseIf RadioButton4.Checked = True Then

            'Dim bso As Double = TextBox1.Text

            Module_AddBeams.HoriBeamSpaceOffset = TextBox1.Text

        ElseIf RadioButton15.Checked = True Then

            Module_AddBeams.UseNumHoriBeams = "Yes"

            Module_AddBeams.NumHoriBeams = TextBox5.Text - 2


        End If


        If RadioButton5.Checked = True Then

            Module_AddBeams.HoriBeamWidth = "10"

        ElseIf RadioButton6.Checked = True Then

            Module_AddBeams.HoriBeamWidth = "12"

        ElseIf RadioButton7.Checked = True Then

            Module_AddBeams.HoriBeamWidth = TextBox2.Text

        End If


        If RadioButton8.Checked = True Then

            Module_AddBeams.VertBeamSpaceOffset = "120"

        ElseIf RadioButton9.Checked = True Then

            Module_AddBeams.VertBeamSpaceOffset = "144"

        ElseIf RadioButton10.Checked = True Then

            Module_AddBeams.VertBeamSpaceOffset = "168"

        ElseIf RadioButton11.Checked = True Then

            Module_AddBeams.VertBeamSpaceOffset = TextBox3.Text

        ElseIf RadioButton16.Checked = True Then

            Module_AddBeams.UseNumVertBeams = "Yes"

            Module_AddBeams.NumVertBeams = TextBox6.Text - 2

        End If


        If RadioButton12.Checked = True Then

            Module_AddBeams.VertBeamWidth = "10"

        ElseIf RadioButton13.Checked = True Then

            Module_AddBeams.VertBeamWidth = "12"

        ElseIf RadioButton14.Checked = True Then

            Module_AddBeams.VertBeamWidth = TextBox4.Text

        End If

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Module_AddBeams.BeamSpaceCancel = "Cancel"

        Me.Close()

    End Sub





    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        TextBox1.Enabled = True
        TextBox5.Enabled = False
        Module_AddBeams.UseNumHoriBeams = "No"

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

        TextBox2.Enabled = True

    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged

        TextBox3.Enabled = True
        TextBox6.Enabled = False
        Module_AddBeams.UseNumVertBeams = "No"

    End Sub


    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged

        TextBox4.Enabled = True

    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged

        TextBox1.Enabled = False
        TextBox5.Enabled = True
        Module_AddBeams.UseNumHoriBeams = "Yes"

    End Sub

    Private Sub RadioButton16_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton16.CheckedChanged

        TextBox3.Enabled = False
        TextBox6.Enabled = True
        Module_AddBeams.UseNumVertBeams = "Yes"

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        TextBox1.Enabled = False
        TextBox5.Enabled = False
        Module_AddBeams.UseNumHoriBeams = "No"

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        TextBox1.Enabled = False
        TextBox5.Enabled = False
        Module_AddBeams.UseNumHoriBeams = "No"

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        TextBox1.Enabled = False
        TextBox5.Enabled = False
        Module_AddBeams.UseNumHoriBeams = "No"

    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged

        TextBox3.Enabled = False
        TextBox6.Enabled = False
        Module_AddBeams.UseNumVertBeams = "No"

    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

        TextBox3.Enabled = False
        TextBox6.Enabled = False
        Module_AddBeams.UseNumVertBeams = "No"

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        TextBox3.Enabled = False
        TextBox6.Enabled = False
        Module_AddBeams.UseNumVertBeams = "No"

    End Sub
End Class