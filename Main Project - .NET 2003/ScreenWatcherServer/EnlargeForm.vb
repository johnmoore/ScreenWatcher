'ScreenWatcher
'Copyright (C) 2010 John Moore and Andre Wiggins
'
'Homepage: http://www.screenwatcher.net
'
'ScreenWatcher, composed of the client application, server application, and Class Designer,
'   is � (copyright) of and was originally written by John Moore and Andr� Wiggins.
'ScreenWatcher has been released under the Reciprocal Public License. 
'   You should have received a copy of this license with the source code.
'   If not, you can find a copy of it at http://www.programiscellaneous.com/programming-projects/screenwatcher/source-code-licensing/.
'   Additionally, the license template can be found at http://www.opensource.org/licenses/rpl1.5.txt.
'ScreenWatcher is protected under U.S. law from infringement of any rights not granted to the end user
'   of the program by this license, including the acquisition of profit from redistribution
'   of the original source or any variant of it except as permitted by the license itself.

Public Class EnlargeForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal NewPic As Image)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Image = NewPic

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
    Friend WithEvents picMain As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.picMain = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'picMain
        '
        Me.picMain.BackColor = System.Drawing.SystemColors.ControlDark
        Me.picMain.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picMain.Location = New System.Drawing.Point(0, 0)
        Me.picMain.Name = "picMain"
        Me.picMain.Size = New System.Drawing.Size(300, 300)
        Me.picMain.TabIndex = 0
        Me.picMain.TabStop = False
        '
        'EnlargeForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(300, 300)
        Me.Controls.Add(Me.picMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "EnlargeForm"
        Me.Text = "EnlargeForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Image As Image

    Private Sub EnlargeForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With Me
            .Top = 0
            .Left = 0
            .Height = Screen.PrimaryScreen.Bounds.Height
            .Width = Screen.PrimaryScreen.Bounds.Width
            .Cursor = System.Windows.Forms.Cursors.Hand
            .TopMost = True
        End With

        With picMain
            .Image = Image
            .SizeMode = PictureBoxSizeMode.CenterImage
            .Left = 0
            .Top = 0
            .Height = Screen.PrimaryScreen.Bounds.Height
            .Width = Screen.PrimaryScreen.Bounds.Width
            .Cursor = System.Windows.Forms.Cursors.Hand
        End With

        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Me.picMain, "Click to Zoom Out")

    End Sub

    Private Sub picMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picMain.Click
        Me.TopMost = False
        Me.Close()
    End Sub

    Private Sub EnlargeForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        Me.TopMost = False
        Me.Close()
    End Sub
End Class
