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
        GroupBox2 = New System.Windows.Forms.GroupBox()
        TextBox1 = New System.Windows.Forms.TextBox()
        GroupBox4 = New System.Windows.Forms.GroupBox()
        SealsList = New System.Windows.Forms.CheckedListBox()
        GroupBox1 = New System.Windows.Forms.GroupBox()
        GroupBox5 = New System.Windows.Forms.GroupBox()
        TextBox3 = New System.Windows.Forms.TextBox()
        GroupBox8 = New System.Windows.Forms.GroupBox()
        SheetLabels = New System.Windows.Forms.ComboBox()
        GroupBox3 = New System.Windows.Forms.GroupBox()
        TextBox2 = New System.Windows.Forms.TextBox()
        Button2 = New System.Windows.Forms.Button()
        Button1 = New System.Windows.Forms.Button()
        GroupBox7 = New System.Windows.Forms.GroupBox()
        RFR = New System.Windows.Forms.RadioButton()
        WSFW = New System.Windows.Forms.RadioButton()
        DWF = New System.Windows.Forms.RadioButton()
        GroupBox2.SuspendLayout()
        GroupBox4.SuspendLayout()
        GroupBox1.SuspendLayout()
        GroupBox5.SuspendLayout()
        GroupBox8.SuspendLayout()
        GroupBox3.SuspendLayout()
        GroupBox7.SuspendLayout()
        SuspendLayout()
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(TextBox1)
        GroupBox2.Location = New System.Drawing.Point(10, 22)
        GroupBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox2.Size = New System.Drawing.Size(282, 57)
        GroupBox2.TabIndex = 23
        GroupBox2.TabStop = False
        GroupBox2.Text = "Builder Name"
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New System.Drawing.Point(7, 20)
        TextBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New System.Drawing.Size(268, 23)
        TextBox1.TabIndex = 1
        ' 
        ' GroupBox4
        ' 
        GroupBox4.Controls.Add(SealsList)
        GroupBox4.Location = New System.Drawing.Point(159, 197)
        GroupBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox4.Name = "GroupBox4"
        GroupBox4.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox4.Size = New System.Drawing.Size(134, 123)
        GroupBox4.TabIndex = 25
        GroupBox4.TabStop = False
        GroupBox4.Text = "Seals To Print"
        ' 
        ' SealsList
        ' 
        SealsList.CheckOnClick = True
        SealsList.FormattingEnabled = True
        SealsList.IntegralHeight = False
        SealsList.Items.AddRange(New Object() {"TML-TX"})
        SealsList.Location = New System.Drawing.Point(7, 20)
        SealsList.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        SealsList.Name = "SealsList"
        SealsList.Size = New System.Drawing.Size(119, 94)
        SealsList.TabIndex = 3
        SealsList.Tag = "TML-TX"
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(GroupBox5)
        GroupBox1.Controls.Add(GroupBox8)
        GroupBox1.Controls.Add(GroupBox3)
        GroupBox1.Controls.Add(Button2)
        GroupBox1.Controls.Add(Button1)
        GroupBox1.Controls.Add(GroupBox7)
        GroupBox1.Controls.Add(GroupBox4)
        GroupBox1.Controls.Add(GroupBox2)
        GroupBox1.Location = New System.Drawing.Point(14, 14)
        GroupBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox1.Size = New System.Drawing.Size(304, 363)
        GroupBox1.TabIndex = 5
        GroupBox1.TabStop = False
        GroupBox1.Text = "Available Properties"
        ' 
        ' GroupBox5
        ' 
        GroupBox5.Controls.Add(TextBox3)
        GroupBox5.Location = New System.Drawing.Point(10, 135)
        GroupBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox5.Name = "GroupBox5"
        GroupBox5.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox5.Size = New System.Drawing.Size(154, 50)
        GroupBox5.TabIndex = 29
        GroupBox5.TabStop = False
        GroupBox5.Text = "Project Number"
        ' 
        ' TextBox3
        ' 
        TextBox3.Location = New System.Drawing.Point(7, 20)
        TextBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New System.Drawing.Size(139, 23)
        TextBox3.TabIndex = 2
        ' 
        ' GroupBox8
        ' 
        GroupBox8.Controls.Add(SheetLabels)
        GroupBox8.Location = New System.Drawing.Point(186, 135)
        GroupBox8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox8.Name = "GroupBox8"
        GroupBox8.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox8.Size = New System.Drawing.Size(107, 50)
        GroupBox8.TabIndex = 28
        GroupBox8.TabStop = False
        GroupBox8.Text = "Sheet Labeling"
        ' 
        ' SheetLabels
        ' 
        SheetLabels.BackColor = Drawing.SystemColors.Window
        SheetLabels.FormattingEnabled = True
        SheetLabels.Items.AddRange(New Object() {"S", "FR"})
        SheetLabels.Location = New System.Drawing.Point(7, 20)
        SheetLabels.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        SheetLabels.Name = "SheetLabels"
        SheetLabels.Size = New System.Drawing.Size(87, 23)
        SheetLabels.TabIndex = 28
        ' 
        ' GroupBox3
        ' 
        GroupBox3.Controls.Add(TextBox2)
        GroupBox3.Location = New System.Drawing.Point(10, 78)
        GroupBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox3.Size = New System.Drawing.Size(282, 50)
        GroupBox3.TabIndex = 27
        GroupBox3.TabStop = False
        GroupBox3.Text = "Plan Name/Number"
        ' 
        ' TextBox2
        ' 
        TextBox2.Location = New System.Drawing.Point(7, 20)
        TextBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New System.Drawing.Size(268, 23)
        TextBox2.TabIndex = 2
        ' 
        ' Button2
        ' 
        Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Button2.Location = New System.Drawing.Point(198, 328)
        Button2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Button2.Name = "Button2"
        Button2.Size = New System.Drawing.Size(88, 27)
        Button2.TabIndex = 10
        Button2.Text = "Cancel"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Button1.Location = New System.Drawing.Point(103, 328)
        Button1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Button1.Name = "Button1"
        Button1.Size = New System.Drawing.Size(88, 27)
        Button1.TabIndex = 9
        Button1.Text = "OK"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' GroupBox7
        ' 
        GroupBox7.Controls.Add(DWF)
        GroupBox7.Controls.Add(RFR)
        GroupBox7.Controls.Add(WSFW)
        GroupBox7.Location = New System.Drawing.Point(10, 197)
        GroupBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox7.Name = "GroupBox7"
        GroupBox7.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        GroupBox7.Size = New System.Drawing.Size(134, 123)
        GroupBox7.TabIndex = 26
        GroupBox7.TabStop = False
        GroupBox7.Text = "Builder Specific Info"
        ' 
        ' RFR
        ' 
        RFR.AutoCheck = False
        RFR.AutoSize = True
        RFR.Location = New System.Drawing.Point(7, 48)
        RFR.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        RFR.Name = "RFR"
        RFR.Size = New System.Drawing.Size(45, 19)
        RFR.TabIndex = 3
        RFR.TabStop = True
        RFR.Text = "RFR"
        RFR.UseVisualStyleBackColor = True
        ' 
        ' WSFW
        ' 
        WSFW.AutoCheck = False
        WSFW.AutoSize = True
        WSFW.Location = New System.Drawing.Point(7, 22)
        WSFW.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        WSFW.Name = "WSFW"
        WSFW.Size = New System.Drawing.Size(59, 19)
        WSFW.TabIndex = 2
        WSFW.TabStop = True
        WSFW.Text = "WSFW"
        WSFW.UseVisualStyleBackColor = True
        ' 
        ' DWF
        ' 
        DWF.AutoSize = True
        DWF.Location = New System.Drawing.Point(7, 73)
        DWF.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        DWF.Name = "DWF"
        DWF.Size = New System.Drawing.Size(50, 19)
        DWF.TabIndex = 4
        DWF.TabStop = True
        DWF.Text = "DWF"
        DWF.UseVisualStyleBackColor = True
        ' 
        ' Form_FramingProperties
        ' 
        AcceptButton = Button1
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        AutoSize = True
        CancelButton = Button2
        ClientSize = New System.Drawing.Size(334, 387)
        Controls.Add(GroupBox1)
        Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Name = "Form_FramingProperties"
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Text = "Form_DrawingProperties"
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        GroupBox4.ResumeLayout(False)
        GroupBox1.ResumeLayout(False)
        GroupBox5.ResumeLayout(False)
        GroupBox5.PerformLayout()
        GroupBox8.ResumeLayout(False)
        GroupBox3.ResumeLayout(False)
        GroupBox3.PerformLayout()
        GroupBox7.ResumeLayout(False)
        GroupBox7.PerformLayout()
        ResumeLayout(False)

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
    Friend WithEvents DWF As System.Windows.Forms.RadioButton
End Class
