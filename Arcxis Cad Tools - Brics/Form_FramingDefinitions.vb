Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports TextBox = System.Windows.Forms.TextBox
Public Class Form_FramingDefinitions

    Private ReadOnly _textBoxes As TextBox()
    Private _originalLabelValue As Integer

    Public Sub New(originalValue As Integer)
        InitializeComponent()
        _textBoxes = New TextBox() {TextBox1, TextBox3, TextBox4, TextBox2}
        _originalLabelValue = originalValue
        Label3.Text = originalValue.ToString()

        For Each tb In _textBoxes
            AddHandler tb.TextChanged, AddressOf TextBoxes_TextChanged
        Next
    End Sub

    Private Sub TextBoxes_TextChanged(sender As Object, e As EventArgs)
        Dim sum As Integer = 0
        For Each tb In _textBoxes
            Dim val As Integer
            If Integer.TryParse(tb.Text, val) Then
                sum += val
            End If
        Next

        Dim result As Integer = _originalLabelValue - sum
        Label3.Text = result.ToString()

        ' Highlight red if less than 0, else default color
        If result < 0 Then
            Label3.ForeColor = Color.Red
        Else
            Label3.ForeColor = Color.Black
        End If

        ' Hide Continue button if result is exactly 0, show otherwise
        'Button1.Visible = (result = 0)

        ' Define your mappings: (TextBox, ListBox)
        Dim mappings = {
        (TextBox1, ListBox1),
        (TextBox2, ListBox2),
        (TextBox3, ListBox3),
        (TextBox4, ListBox4)
    }

        For Each m In mappings
            Dim tb = m.Item1
            Dim lb = m.Item2
            Dim tbText = tb.Text.Trim()

            ' Enable/disable ListBox based on TextBox value
            If tbText <> "0" And tbText <> "" Then
                lb.BackColor = Color.Red
                lb.Enabled = True
            Else
                lb.BackColor = SystemColors.Window
                lb.Enabled = False
                lb.ClearSelected() ' Unselect all items if locking
            End If

            ' Enable/disable TextBox if value is 0 or empty and _originalLabelValue is 0
            If tbText = "" AndAlso Label3.Text = 0 Then
                tb.Enabled = False
            Else
                tb.Enabled = True
            End If
        Next

        UpdateContinueButtonVisibility()

    End Sub

    Public ReadOnly Property GroupedEntries As List(Of Tuple(Of String, String))
        Get
            Dim result As New List(Of Tuple(Of String, String))()

            ' Define your mappings: (TextBox, GroupBox, ListBox)
            Dim mappings = {
            (TextBox1, GroupBox1, ListBox1),
            (TextBox2, GroupBox2, ListBox2),
            (TextBox3, GroupBox3, ListBox3),
            (TextBox4, GroupBox4, ListBox4)
        }

            For Each m In mappings
                Dim count As Integer
                If Integer.TryParse(m.Item1.Text, count) AndAlso count > 0 Then
                    Dim groupName As String = m.Item2.Text
                    Dim listValue As String = If(m.Item3.SelectedItem?.ToString(), "")
                    For i = 1 To count
                        result.Add(Tuple.Create(groupName, listValue))
                    Next
                End If
            Next

            Return result
        End Get
    End Property

    Private Sub UpdateContinueButtonVisibility()
        ' Calculate the sum of all textbox values
        Dim sum As Integer = 0
        For Each tb In _textBoxes
            Dim val As Integer
            If Integer.TryParse(tb.Text, val) Then
                sum += val
            End If
        Next

        Dim result As Integer = _originalLabelValue - sum

        ' Only allow continue if result is 0 and all required listboxes are selected
        If result = 0 Then
            ' Check that for each textbox > 0, the corresponding listbox has a selection
            Dim mappings = {
            (TextBox1, ListBox1),
            (TextBox2, ListBox2),
            (TextBox3, ListBox3),
            (TextBox4, ListBox4)
        }
            For Each m In mappings
                Dim tb = m.Item1
                Dim lb = m.Item2
                Dim tbVal As Integer
                If Integer.TryParse(tb.Text, tbVal) AndAlso tbVal > 0 Then
                    If lb.SelectedIndex < 0 Then
                        Button1.Visible = False
                        Return
                    End If
                End If
            Next
            Button1.Visible = True
        Else
            Button1.Visible = False
        End If
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ListBox1.BackColor = SystemColors.Window
        UpdateContinueButtonVisibility()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        ListBox2.BackColor = SystemColors.Window
        UpdateContinueButtonVisibility()
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        ListBox3.BackColor = SystemColors.Window
        UpdateContinueButtonVisibility()
    End Sub

    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        ListBox4.BackColor = SystemColors.Window
        UpdateContinueButtonVisibility()
    End Sub

End Class