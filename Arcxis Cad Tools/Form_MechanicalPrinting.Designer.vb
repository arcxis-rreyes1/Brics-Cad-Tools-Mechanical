<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_MechanicalPrinting
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Counties = New System.Windows.Forms.GroupBox()
        Me.CountyList = New System.Windows.Forms.CheckedListBox()
        Me.Manufacturers = New System.Windows.Forms.GroupBox()
        Me.ManList = New System.Windows.Forms.CheckedListBox()
        Me.GasType = New System.Windows.Forms.GroupBox()
        Me.GasList = New System.Windows.Forms.CheckedListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.Counties.SuspendLayout()
        Me.Manufacturers.SuspendLayout()
        Me.GasType.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Counties)
        Me.GroupBox1.Controls.Add(Me.Manufacturers)
        Me.GroupBox1.Controls.Add(Me.GasType)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(222, 431)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Automated Printing"
        '
        'Counties
        '
        Me.Counties.Controls.Add(Me.CountyList)
        Me.Counties.Location = New System.Drawing.Point(8, 295)
        Me.Counties.Name = "Counties"
        Me.Counties.Size = New System.Drawing.Size(202, 96)
        Me.Counties.TabIndex = 9
        Me.Counties.TabStop = False
        Me.Counties.Text = "Counties"
        '
        'CountyList
        '
        Me.CountyList.CheckOnClick = True
        Me.CountyList.FormattingEnabled = True
        Me.CountyList.IntegralHeight = False
        Me.CountyList.Items.AddRange(New Object() {"All"})
        Me.CountyList.Location = New System.Drawing.Point(6, 19)
        Me.CountyList.Name = "CountyList"
        Me.CountyList.Size = New System.Drawing.Size(190, 66)
        Me.CountyList.TabIndex = 5
        Me.CountyList.Tag = "TML-TX"
        '
        'Manufacturers
        '
        Me.Manufacturers.Controls.Add(Me.ManList)
        Me.Manufacturers.Location = New System.Drawing.Point(8, 117)
        Me.Manufacturers.Name = "Manufacturers"
        Me.Manufacturers.Size = New System.Drawing.Size(202, 83)
        Me.Manufacturers.TabIndex = 8
        Me.Manufacturers.TabStop = False
        Me.Manufacturers.Text = "Manufacturer(s)"
        '
        'ManList
        '
        Me.ManList.CheckOnClick = True
        Me.ManList.FormattingEnabled = True
        Me.ManList.IntegralHeight = False
        Me.ManList.Items.AddRange(New Object() {"All"})
        Me.ManList.Location = New System.Drawing.Point(6, 19)
        Me.ManList.Name = "ManList"
        Me.ManList.Size = New System.Drawing.Size(190, 51)
        Me.ManList.TabIndex = 3
        Me.ManList.Tag = "All"
        '
        'GasType
        '
        Me.GasType.Controls.Add(Me.GasList)
        Me.GasType.Location = New System.Drawing.Point(8, 206)
        Me.GasType.Name = "GasType"
        Me.GasType.Size = New System.Drawing.Size(202, 83)
        Me.GasType.TabIndex = 8
        Me.GasType.TabStop = False
        Me.GasType.Text = "Gas Type"
        '
        'GasList
        '
        Me.GasList.CheckOnClick = True
        Me.GasList.FormattingEnabled = True
        Me.GasList.IntegralHeight = False
        Me.GasList.Items.AddRange(New Object() {"All"})
        Me.GasList.Location = New System.Drawing.Point(6, 19)
        Me.GasList.Name = "GasList"
        Me.GasList.Size = New System.Drawing.Size(190, 51)
        Me.GasList.TabIndex = 4
        Me.GasList.Tag = "TML-TX"
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(129, 397)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Location = New System.Drawing.Point(47, 397)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Location = New System.Drawing.Point(7, 68)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(204, 43)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Plan Name/Number"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(6, 17)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(192, 20)
        Me.TextBox2.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 19)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(204, 43)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Builder Name"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(6, 17)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(192, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Form_MechanicalPrinting
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(246, 453)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form_MechanicalPrinting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form_MechanicalPrinting"
        Me.GroupBox1.ResumeLayout(False)
        Me.Counties.ResumeLayout(False)
        Me.Manufacturers.ResumeLayout(False)
        Me.GasType.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Manufacturers As Windows.Forms.GroupBox
    Friend WithEvents ManList As Windows.Forms.CheckedListBox
    Friend WithEvents GasType As Windows.Forms.GroupBox
    Friend WithEvents GasList As Windows.Forms.CheckedListBox
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents Counties As Windows.Forms.GroupBox
    Friend WithEvents CountyList As Windows.Forms.CheckedListBox
End Class
