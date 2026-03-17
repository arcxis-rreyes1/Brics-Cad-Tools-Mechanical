<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Arcxis_TB4
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        TextBox1 = New System.Windows.Forms.TextBox()
        Label1 = New System.Windows.Forms.Label()
        Button2 = New System.Windows.Forms.Button()
        Button1 = New System.Windows.Forms.Button()
        GroupBox9 = New System.Windows.Forms.GroupBox()
        CheckedListBox2 = New System.Windows.Forms.CheckedListBox()
        GroupBox9.SuspendLayout()
        SuspendLayout()
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New System.Drawing.Point(282, 22)
        TextBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New System.Drawing.Size(116, 23)
        TextBox1.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(14, 25)
        Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(235, 15)
        Label1.TabIndex = 3
        Label1.Text = "Enter Custom Layout Letter for Elevation #?"
        ' 
        ' Button2
        ' 
        Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Button2.Location = New System.Drawing.Point(310, 73)
        Button2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Button2.Name = "Button2"
        Button2.Size = New System.Drawing.Size(88, 27)
        Button2.TabIndex = 3
        Button2.Text = "Cancel"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Location = New System.Drawing.Point(215, 73)
        Button1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Button1.Name = "Button1"
        Button1.Size = New System.Drawing.Size(88, 27)
        Button1.TabIndex = 2
        Button1.Text = "OK"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' GroupBox9
        ' 
        GroupBox9.Controls.Add(CheckedListBox2)
        GroupBox9.Location = New System.Drawing.Point(14, 56)
        GroupBox9.Name = "GroupBox9"
        GroupBox9.Size = New System.Drawing.Size(194, 57)
        GroupBox9.TabIndex = 11
        GroupBox9.TabStop = False
        GroupBox9.Text = "Multi Market Option"
        ' 
        ' CheckedListBox2
        ' 
        CheckedListBox2.FormattingEnabled = True
        CheckedListBox2.Location = New System.Drawing.Point(6, 22)
        CheckedListBox2.Name = "CheckedListBox2"
        CheckedListBox2.Size = New System.Drawing.Size(182, 22)
        CheckedListBox2.TabIndex = 0
        ' 
        ' Form_Arcxis_TB4
        ' 
        AcceptButton = Button1
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        AutoSize = True
        CancelButton = Button2
        ClientSize = New System.Drawing.Size(430, 140)
        Controls.Add(GroupBox9)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(TextBox1)
        Controls.Add(Label1)
        Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Name = "Form_Arcxis_TB4"
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Text = "Custom Layout"
        GroupBox9.ResumeLayout(False)
        ResumeLayout(False)
        PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckedListBox2 As System.Windows.Forms.CheckedListBox
End Class
