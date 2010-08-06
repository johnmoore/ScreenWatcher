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

Public Class MessageOptions
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef NewStudent As Student)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Student = NewStudent
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
    Friend WithEvents grpMessageType As System.Windows.Forms.GroupBox
    Friend WithEvents grpMessageOptions As System.Windows.Forms.GroupBox
    Friend WithEvents radFullScreen As System.Windows.Forms.RadioButton
    Friend WithEvents radMsgBox As System.Windows.Forms.RadioButton
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents txtTime As System.Windows.Forms.TextBox
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblTime As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpMessageType = New System.Windows.Forms.GroupBox
        Me.radMsgBox = New System.Windows.Forms.RadioButton
        Me.radFullScreen = New System.Windows.Forms.RadioButton
        Me.grpMessageOptions = New System.Windows.Forms.GroupBox
        Me.txtTime = New System.Windows.Forms.TextBox
        Me.lblTime = New System.Windows.Forms.Label
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.lblMessage = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me.txtTitle = New System.Windows.Forms.TextBox
        Me.btnSend = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.grpMessageType.SuspendLayout()
        Me.grpMessageOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMessageType
        '
        Me.grpMessageType.Controls.Add(Me.radMsgBox)
        Me.grpMessageType.Controls.Add(Me.radFullScreen)
        Me.grpMessageType.Location = New System.Drawing.Point(8, 8)
        Me.grpMessageType.Name = "grpMessageType"
        Me.grpMessageType.Size = New System.Drawing.Size(184, 136)
        Me.grpMessageType.TabIndex = 0
        Me.grpMessageType.TabStop = False
        Me.grpMessageType.Text = "MessageType"
        '
        'radMsgBox
        '
        Me.radMsgBox.Location = New System.Drawing.Point(24, 80)
        Me.radMsgBox.Name = "radMsgBox"
        Me.radMsgBox.TabIndex = 1
        Me.radMsgBox.Text = "Message Box"
        '
        'radFullScreen
        '
        Me.radFullScreen.Checked = True
        Me.radFullScreen.Location = New System.Drawing.Point(24, 40)
        Me.radFullScreen.Name = "radFullScreen"
        Me.radFullScreen.TabIndex = 0
        Me.radFullScreen.TabStop = True
        Me.radFullScreen.Text = "FullScreen"
        '
        'grpMessageOptions
        '
        Me.grpMessageOptions.Controls.Add(Me.txtTime)
        Me.grpMessageOptions.Controls.Add(Me.lblTime)
        Me.grpMessageOptions.Controls.Add(Me.txtMessage)
        Me.grpMessageOptions.Controls.Add(Me.lblMessage)
        Me.grpMessageOptions.Controls.Add(Me.lblTitle)
        Me.grpMessageOptions.Controls.Add(Me.txtTitle)
        Me.grpMessageOptions.Location = New System.Drawing.Point(200, 8)
        Me.grpMessageOptions.Name = "grpMessageOptions"
        Me.grpMessageOptions.Size = New System.Drawing.Size(200, 232)
        Me.grpMessageOptions.TabIndex = 1
        Me.grpMessageOptions.TabStop = False
        Me.grpMessageOptions.Text = "Message Options"
        '
        'txtTime
        '
        Me.txtTime.Location = New System.Drawing.Point(32, 176)
        Me.txtTime.Name = "txtTime"
        Me.txtTime.Size = New System.Drawing.Size(152, 22)
        Me.txtTime.TabIndex = 5
        Me.txtTime.Text = "10"
        '
        'lblTime
        '
        Me.lblTime.Location = New System.Drawing.Point(16, 136)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(176, 40)
        Me.lblTime.TabIndex = 4
        Me.lblTime.Text = "Seconds to Display Message (0 - indefinite):"
        Me.lblTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(32, 104)
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(152, 22)
        Me.txtMessage.TabIndex = 3
        Me.txtMessage.Text = "See me after class!"
        '
        'lblMessage
        '
        Me.lblMessage.Location = New System.Drawing.Point(8, 80)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.TabIndex = 2
        Me.lblMessage.Text = "Message:"
        Me.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(8, 24)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "Title:"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(32, 48)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(152, 22)
        Me.txtTitle.TabIndex = 0
        Me.txtTitle.Text = "Get Back to Work!"
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(48, 160)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(112, 23)
        Me.btnSend.TabIndex = 2
        Me.btnSend.Text = "SendMessage"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(48, 200)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(112, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'MessageOptions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(424, 256)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.grpMessageOptions)
        Me.Controls.Add(Me.grpMessageType)
        Me.Name = "MessageOptions"
        Me.Text = "MessageOptions"
        Me.grpMessageType.ResumeLayout(False)
        Me.grpMessageOptions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Message As Message
    Private Student As Student
    Private SWServer As SWServer

    Private fldTimeSpan As TimeSpan
    Private fldMessage As String
    Private fldTitle As String
    Private fldMessageType As Message.Messagetype
    Private fldClient As Student

    Public ReadOnly Property KeyName() As String
        Get
            Return Student.KeyName
        End Get
    End Property

    Private Sub MessageOptions_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SWServer = Me.Owner
    End Sub

    Private Sub radButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radFullScreen.CheckedChanged, radMsgBox.CheckedChanged
        If Me.radFullScreen.Checked = True And Me.radMsgBox.Checked = False Then
            fldMessageType = Message.Messagetype.FullScreen
            Me.lblTime.Enabled = True
            Me.txtTime.Enabled = True

        ElseIf Me.radMsgBox.Checked = True And Me.radFullScreen.Checked = False Then
            fldMessageType = Message.Messagetype.Messagebox
            Me.lblTime.Enabled = False
            Me.txtTime.Enabled = False
        End If

    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click
        fldMessage = Me.txtMessage.Text
        fldTitle = Me.txtTitle.Text

        If fldMessageType = Message.Messagetype.FullScreen And Me.txtTime.Text <> "0" Then
            Dim Time As String = "0:0:" & Me.txtTime.Text
            fldTimeSpan = TimeSpan.Parse(Time)
        End If

        Dim NewMessage As New Message(Student, fldMessageType, fldMessage, fldTitle, fldTimeSpan)

        Dim DisplayMsg As String
        DisplayMsg = "Are you sure you want to send:" & vbNewLine & IIf(fldMessageType = Message.Messagetype.FullScreen, ("- a fullscreen message" & vbNewLine), ("- a messagebox" & vbNewLine)) _
            & "- titled '" & fldTitle & "'" & vbNewLine & "- displaying '" & fldMessage & "'" & _
            IIf(fldMessageType = Message.Messagetype.FullScreen, IIf(Me.txtTime.Text = "0", (vbNewLine & " - for an indefinite amount of time?"), (vbNewLine & "- for " & fldTimeSpan.Seconds.ToString & " seconds" & vbNewLine)), "?")

        Dim Continue As DialogResult = MessageBox.Show(DisplayMsg, "Are You Sure?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)

        If Continue = DialogResult.No Or Continue = DialogResult.Cancel Then
            Exit Sub
        End If

        SWServer.SetMessage(Student.KeyName, NewMessage)

        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim NewMessage As New Message
        SWServer.SetMessage(Student.KeyName, NewMessage)

        Me.Close()
    End Sub

End Class
