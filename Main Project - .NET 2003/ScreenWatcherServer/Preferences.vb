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

Public Class Preferences
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
    Friend WithEvents trkTansparency As System.Windows.Forms.TrackBar
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkStayOnTop As System.Windows.Forms.CheckBox
    Friend WithEvents lblTransparency As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.trkTansparency = New System.Windows.Forms.TrackBar
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkStayOnTop = New System.Windows.Forms.CheckBox
        Me.lblTransparency = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        CType(Me.trkTansparency, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'trkTansparency
        '
        Me.trkTansparency.Location = New System.Drawing.Point(10, 28)
        Me.trkTansparency.Maximum = 100
        Me.trkTansparency.Minimum = 10
        Me.trkTansparency.Name = "trkTansparency"
        Me.trkTansparency.Size = New System.Drawing.Size(249, 56)
        Me.trkTansparency.TabIndex = 0
        Me.trkTansparency.TickFrequency = 10
        Me.trkTansparency.Value = 100
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkStayOnTop)
        Me.GroupBox1.Controls.Add(Me.trkTansparency)
        Me.GroupBox1.Controls.Add(Me.lblTransparency)
        Me.GroupBox1.Location = New System.Drawing.Point(19, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(269, 130)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Opacity"
        '
        'chkStayOnTop
        '
        Me.chkStayOnTop.Location = New System.Drawing.Point(134, 92)
        Me.chkStayOnTop.Name = "chkStayOnTop"
        Me.chkStayOnTop.Size = New System.Drawing.Size(125, 28)
        Me.chkStayOnTop.TabIndex = 1
        Me.chkStayOnTop.Text = "Stay On Top"
        '
        'lblTransparency
        '
        Me.lblTransparency.Location = New System.Drawing.Point(10, 92)
        Me.lblTransparency.Name = "lblTransparency"
        Me.lblTransparency.Size = New System.Drawing.Size(67, 27)
        Me.lblTransparency.TabIndex = 2
        Me.lblTransparency.Text = "100%"
        Me.lblTransparency.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(192, 166)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(90, 27)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(86, 166)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(90, 27)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'Preferences
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(307, 209)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Preferences"
        Me.Text = "Preferences"
        CType(Me.trkTansparency, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim SWServer As SWServer
    Dim PreviousOpac As Double
    Dim PreviousTopMost As Boolean

    Private Sub Preferences_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SWServer = Me.Owner
        PreviousOpac = SWServer.Opacity
        PreviousTopMost = SWServer.TopMost

        Me.trkTansparency.Value = Int(PreviousOpac * 100)
        Me.lblTransparency.Text = PreviousOpac * 100 & "%"

        Me.chkStayOnTop.Checked = PreviousTopMost
    End Sub

    Private Sub trkTansparency_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkTansparency.Scroll
        Dim NewOpacValue As Double = Me.trkTansparency.Value
        Me.lblTransparency.Text = NewOpacValue.ToString & "%"

        SWServer.Opacity = NewOpacValue / 100

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SWServer.TopMost = Me.chkStayOnTop.Checked
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        SWServer.Opacity = PreviousOpac
        SWServer.TopMost = PreviousTopMost
        Me.Close()
    End Sub

End Class