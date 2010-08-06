'ScreenWatcher
'Copyright (C) 2010 John Moore and Andre Wiggins
'
'Homepage: http://www.screenwatcher.net
'
'ScreenWatcher, composed of the client application, server application, and Class Designer,
'   is © (copyright) of and was originally written by John Moore and André Wiggins.
'ScreenWatcher has been released under the Reciprocal Public License. 
'   You should have received a copy of this license with the source code.
'   If not, you can find a copy of it at http://www.programiscellaneous.com/programming-projects/screenwatcher/licensing-copyright/.
'   Additionally, the license template can be found at http://www.opensource.org/licenses/rpl1.5.txt.
'ScreenWatcher is protected under U.S. law from infringement of any rights not granted to the end user
'   of the program by this license, including the acquisition of profit from redistribution
'   of the original source or any variant of it except as permitted by the license itself.

Public Class PeriodDialog
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblRowsPrompt As System.Windows.Forms.Label
    Friend WithEvents lblSeatsPerRowPrompt As System.Windows.Forms.Label
    Friend WithEvents lblPeriodNumberPrompt As System.Windows.Forms.Label
    Friend WithEvents txtPeriodNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtRows As System.Windows.Forms.TextBox
    Friend WithEvents txtSeatsPerRow As System.Windows.Forms.TextBox
    Friend WithEvents lblPeriodTitlePrompt As System.Windows.Forms.Label
    Friend WithEvents txtPeriodTitle As System.Windows.Forms.TextBox
    Friend WithEvents btnSubmit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblRowsPrompt = New System.Windows.Forms.Label
        Me.lblSeatsPerRowPrompt = New System.Windows.Forms.Label
        Me.lblPeriodNumberPrompt = New System.Windows.Forms.Label
        Me.txtPeriodNumber = New System.Windows.Forms.TextBox
        Me.txtRows = New System.Windows.Forms.TextBox
        Me.txtSeatsPerRow = New System.Windows.Forms.TextBox
        Me.lblPeriodTitlePrompt = New System.Windows.Forms.Label
        Me.txtPeriodTitle = New System.Windows.Forms.TextBox
        Me.btnSubmit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblRowsPrompt
        '
        Me.lblRowsPrompt.AutoSize = True
        Me.lblRowsPrompt.Location = New System.Drawing.Point(8, 9)
        Me.lblRowsPrompt.Name = "lblRowsPrompt"
        Me.lblRowsPrompt.Size = New System.Drawing.Size(36, 16)
        Me.lblRowsPrompt.TabIndex = 0
        Me.lblRowsPrompt.Text = "Rows:"
        '
        'lblSeatsPerRowPrompt
        '
        Me.lblSeatsPerRowPrompt.AutoSize = True
        Me.lblSeatsPerRowPrompt.Location = New System.Drawing.Point(8, 35)
        Me.lblSeatsPerRowPrompt.Name = "lblSeatsPerRowPrompt"
        Me.lblSeatsPerRowPrompt.Size = New System.Drawing.Size(77, 16)
        Me.lblSeatsPerRowPrompt.TabIndex = 1
        Me.lblSeatsPerRowPrompt.Text = "Seats per row:"
        '
        'lblPeriodNumberPrompt
        '
        Me.lblPeriodNumberPrompt.AutoSize = True
        Me.lblPeriodNumberPrompt.Location = New System.Drawing.Point(8, 61)
        Me.lblPeriodNumberPrompt.Name = "lblPeriodNumberPrompt"
        Me.lblPeriodNumberPrompt.Size = New System.Drawing.Size(83, 16)
        Me.lblPeriodNumberPrompt.TabIndex = 2
        Me.lblPeriodNumberPrompt.Text = "Period Number:"
        '
        'txtPeriodNumber
        '
        Me.txtPeriodNumber.Location = New System.Drawing.Point(89, 61)
        Me.txtPeriodNumber.Name = "txtPeriodNumber"
        Me.txtPeriodNumber.Size = New System.Drawing.Size(36, 20)
        Me.txtPeriodNumber.TabIndex = 2
        Me.txtPeriodNumber.Text = ""
        '
        'txtRows
        '
        Me.txtRows.Location = New System.Drawing.Point(89, 9)
        Me.txtRows.Name = "txtRows"
        Me.txtRows.Size = New System.Drawing.Size(36, 20)
        Me.txtRows.TabIndex = 0
        Me.txtRows.Text = ""
        '
        'txtSeatsPerRow
        '
        Me.txtSeatsPerRow.Location = New System.Drawing.Point(89, 35)
        Me.txtSeatsPerRow.Name = "txtSeatsPerRow"
        Me.txtSeatsPerRow.Size = New System.Drawing.Size(36, 20)
        Me.txtSeatsPerRow.TabIndex = 1
        Me.txtSeatsPerRow.Text = ""
        '
        'lblPeriodTitlePrompt
        '
        Me.lblPeriodTitlePrompt.AutoSize = True
        Me.lblPeriodTitlePrompt.Location = New System.Drawing.Point(8, 86)
        Me.lblPeriodTitlePrompt.Name = "lblPeriodTitlePrompt"
        Me.lblPeriodTitlePrompt.Size = New System.Drawing.Size(65, 16)
        Me.lblPeriodTitlePrompt.TabIndex = 6
        Me.lblPeriodTitlePrompt.Text = "Period Title:"
        '
        'txtPeriodTitle
        '
        Me.txtPeriodTitle.Location = New System.Drawing.Point(89, 87)
        Me.txtPeriodTitle.Name = "txtPeriodTitle"
        Me.txtPeriodTitle.Size = New System.Drawing.Size(120, 20)
        Me.txtPeriodTitle.TabIndex = 3
        Me.txtPeriodTitle.Text = ""
        '
        'btnSubmit
        '
        Me.btnSubmit.Location = New System.Drawing.Point(153, 113)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(61, 20)
        Me.btnSubmit.TabIndex = 4
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(89, 113)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(61, 20)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'PeriodDialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(219, 139)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSubmit)
        Me.Controls.Add(Me.txtPeriodTitle)
        Me.Controls.Add(Me.lblPeriodTitlePrompt)
        Me.Controls.Add(Me.txtSeatsPerRow)
        Me.Controls.Add(Me.txtRows)
        Me.Controls.Add(Me.txtPeriodNumber)
        Me.Controls.Add(Me.lblPeriodNumberPrompt)
        Me.Controls.Add(Me.lblSeatsPerRowPrompt)
        Me.Controls.Add(Me.lblRowsPrompt)
        Me.Name = "PeriodDialog"
        Me.Text = "Period Information"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Enum PeriodMode
        NewPeriod = 0
        EditPeriod = 1
    End Enum

    Dim ClassDesigner As ClassDesigner
    Public Mode As PeriodMode

    Private Sub PeriodInformation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ClassDesigner = Me.Owner
        Select Case Mode
            Case PeriodMode.EditPeriod
                LoadEditInterface()
            Case PeriodMode.NewPeriod
                LoadNewInterface()
        End Select
    End Sub

    Private Sub LoadEditInterface()
        btnSubmit.Text = "Save"
        txtRows.Enabled = False
        txtSeatsPerRow.Enabled = False
        txtPeriodTitle.Text = ClassDesigner.pPeriodInfo.Title
        txtRows.Text = ClassDesigner.pPeriodInfo.Dimensions.X
        txtSeatsPerRow.Text = ClassDesigner.pPeriodInfo.Dimensions.Y
        txtPeriodNumber.Text = ClassDesigner.pPeriodInfo.PeriodNumber
    End Sub

    Private Sub LoadNewInterface()
        btnSubmit.Text = "Create"
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Select Case Mode
            Case PeriodMode.EditPeriod
                If Val(txtRows.Text) <> ClassDesigner.pPeriodInfo.Dimensions.X Then
                    MessageBox.Show("You must use the Class Designer main window to change the number of rows.", "Error")
                    Exit Sub
                ElseIf Val(txtSeatsPerRow.Text) <> ClassDesigner.pPeriodInfo.Dimensions.Y Then
                    MessageBox.Show("You must use the Class Designer main window to change the number of seats per row.", "Error")
                    Exit Sub
                ElseIf txtPeriodNumber.Text.Length < 1 Or IsNumeric(txtPeriodNumber.Text) = False Then
                    MessageBox.Show("You must enter a numerical value for the period number.", "Error")
                    Exit Sub
                ElseIf txtPeriodTitle.Text.Length < 1 Then
                    MessageBox.Show("You must enter a period title.", "Error")
                    Exit Sub
                Else
                    ClassDesigner.EditPeriodInfo(Val(txtPeriodNumber.Text), txtPeriodTitle.Text)
                End If
                ClassDesigner.PeriodDialogBolOpen(False)
                Me.Close()
            Case PeriodMode.NewPeriod
                If txtRows.Text.Length < 1 Or IsNumeric(txtRows.Text) = False Then
                    MessageBox.Show("You must enter a numerical value for the number of rows.", "Error")
                    Exit Sub
                ElseIf txtSeatsPerRow.Text.Length < 1 Or IsNumeric(txtSeatsPerRow.Text) = False Then
                    MessageBox.Show("You must enter a numerical value for the number of seats per row.", "Error")
                    Exit Sub
                ElseIf txtPeriodNumber.Text.Length < 1 Or IsNumeric(txtPeriodNumber.Text) = False Then
                    MessageBox.Show("You must enter a numerical value for the period number.", "Error")
                    Exit Sub
                ElseIf txtPeriodTitle.Text.Length < 1 Then
                    MessageBox.Show("You must enter a period title.", "Error")
                    Exit Sub
                Else
                    ClassDesigner.NewPeriodInfo(New Point(Val(txtRows.Text), Val(txtSeatsPerRow.Text)), Val(txtPeriodNumber.Text), txtPeriodTitle.Text)
                End If
                ClassDesigner.NewPeriod()
                ClassDesigner.PeriodDialogBolOpen(False)
                Me.Close()
        End Select
    End Sub
End Class
