Public Class Form_PlanType
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.Text = "" Then
            MsgBox("Please enter a plan type.")
            Exit Sub
        ElseIf Listbox1.Text = "Framing" Then
            Dim frm As New Form_FramingProperties
            frm.Show()
        End If

    End Sub
End Class