Public Class Form_Arcxis_TB4


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Module_Arcxis_TB.ATB_CustomLayoutLetter = TextBox1.Text

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form_Arcxis_TB4_Load(sender As Object, e As EventArgs) Handles Me.Load

        TextBox1.Focus() ' or MyTextBox.Focus()

    End Sub
End Class