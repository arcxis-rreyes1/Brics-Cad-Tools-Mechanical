<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_FramingProperties
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.SealsList = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.SheetLabels = New System.Windows.Forms.ComboBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.RFR = New System.Windows.Forms.RadioButton()
        Me.WSFW = New System.Windows.Forms.RadioButton()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 19)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(242, 49)
        Me.GroupBox2.TabIndex = 23
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Builder Name"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(6, 17)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(230, 20)
        Me.TextBox1.TabIndex = 1
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.SealsList)
        Me.GroupBox4.Location = New System.Drawing.Point(136, 171)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(115, 107)
        Me.GroupBox4.TabIndex = 25
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Seals To Print"
        '
        'SealsList
        '
        Me.SealsList.CheckOnClick = True
        Me.SealsList.FormattingEnabled = True
        Me.SealsList.IntegralHeight = False
        Me.SealsList.Items.AddRange(New Object() {"TML-TX"})
        Me.SealsList.Location = New System.Drawing.Point(6, 17)
        Me.SealsList.Name = "SealsList"
        Me.SealsList.Size = New System.Drawing.Size(103, 82)
        Me.SealsList.TabIndex = 3
        Me.SealsList.Tag = "TML-TX"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox8)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.GroupBox7)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(261, 315)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Available Properties"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.TextBox3)
        Me.GroupBox5.Location = New System.Drawing.Point(9, 117)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(132, 43)
        Me.GroupBox5.TabIndex = 29
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Project Number"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(6, 17)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(120, 20)
        Me.TextBox3.TabIndex = 2
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.SheetLabels)
        Me.GroupBox8.Location = New System.Drawing.Point(159, 117)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(92, 43)
        Me.GroupBox8.TabIndex = 28
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Sheet Labeling"
        '
        'SheetLabels
        '
        Me.SheetLabels.BackColor = System.Drawing.SystemColors.Window
        Me.SheetLabels.FormattingEnabled = True
        Me.SheetLabels.Items.AddRange(New Object() {"S", "FR"})
        Me.SheetLabels.Location = New System.Drawing.Point(6, 17)
        Me.SheetLabels.Name = "SheetLabels"
        Me.SheetLabels.Size = New System.Drawing.Size(75, 21)
        Me.SheetLabels.TabIndex = 28
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 68)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(242, 43)
        Me.GroupBox3.TabIndex = 27
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Plan Name/Number"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(6, 17)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(230, 20)
        Me.TextBox2.TabIndex = 2
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(170, 284)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Location = New System.Drawing.Point(88, 284)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.RFR)
        Me.GroupBox7.Controls.Add(Me.WSFW)
        Me.GroupBox7.Location = New System.Drawing.Point(9, 171)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(115, 107)
        Me.GroupBox7.TabIndex = 26
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Builder Specific Info"
        '
        'RFR
        '
        Me.RFR.AutoCheck = False
        Me.RFR.AutoSize = True
        Me.RFR.Location = New System.Drawing.Point(6, 42)
        Me.RFR.Name = "RFR"
        Me.RFR.Size = New System.Drawing.Size(47, 17)
        Me.RFR.TabIndex = 3
        Me.RFR.TabStop = True
        Me.RFR.Text = "RFR"
        Me.RFR.UseVisualStyleBackColor = True
        '
        'WSFW
        '
        Me.WSFW.AutoCheck = False
        Me.WSFW.AutoSize = True
        Me.WSFW.Location = New System.Drawing.Point(6, 19)
        Me.WSFW.Name = "WSFW"
        Me.WSFW.Size = New System.Drawing.Size(60, 17)
        Me.WSFW.TabIndex = 2
        Me.WSFW.TabStop = True
        Me.WSFW.Text = "WSFW"
        Me.WSFW.UseVisualStyleBackColor = True
        '
        'Form_FramingProperties
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(286, 335)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form_FramingProperties"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form_DrawingProperties"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents SealsList As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents SheetLabels As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents RFR As System.Windows.Forms.RadioButton
    Friend WithEvents WSFW As System.Windows.Forms.RadioButton
End Class
