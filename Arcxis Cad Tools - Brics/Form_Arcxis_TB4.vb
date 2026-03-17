Public Class Form_Arcxis_TB4

    Private Sub ResizeDivisionSection()
        Me.SuspendLayout()

        Dim hasItems As Boolean = CheckedListBox2.Items.Count > 0
        GroupBox9.Visible = hasItems

        If hasItems Then
            Dim visibleCount As Integer = Math.Max(CheckedListBox2.Items.Count, 1)
            CheckedListBox2.IntegralHeight = False
            CheckedListBox2.Height = (visibleCount * CheckedListBox2.ItemHeight) + 4

            Dim gbTitleHeight As Integer = GroupBox9.Height - GroupBox9.ClientSize.Height
            GroupBox9.Height = CheckedListBox2.Top + CheckedListBox2.Height + 6 + gbTitleHeight
        End If

        Dim requiredBottom As Integer = Math.Max(Math.Max(Button1.Bottom, Button2.Bottom), If(GroupBox9.Visible, GroupBox9.Bottom, 0)) + 12
        Me.ClientSize = New System.Drawing.Size(Me.ClientSize.Width, requiredBottom)

        Me.ResumeLayout()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Module_Arcxis_TB.ATB_CustomLayoutLetter = TextBox1.Text

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form_Arcxis_TB4_Load(sender As Object, e As EventArgs) Handles Me.Load

        TextBox1.Focus() ' or MyTextBox.Focus()
        ResizeDivisionSection()

    End Sub
End Class