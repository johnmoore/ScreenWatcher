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

Public Class Student

#Region "Constants"

    Public Const cPICWIDTH As Integer = 170
    Public Const cPICHEIGHT As Integer = 128
    Public Const cPICSPACING As Integer = 2
    Const CPICTOP As Integer = 1
    Const CPICLEFT As Integer = 1

    Public Const cLBLWIDTH As Integer = cPICWIDTH
    Public Const cLBLHEIGHT As Integer = 23
    Public Const cLBLSPACING As Integer = 5
    Const cLBLTOP As Integer = CPICTOP + cPICHEIGHT + cLBLSPACING
    Const cLBLLEFT As Integer = CPICLEFT

    Const cBTNWIDTH As Integer = 23
    Const cBTNHEIGHT As Integer = 23
    Const cBTNSPACING As Integer = cLBLSPACING
    Const cBTNTOP As Integer = CPICTOP + cPICHEIGHT + cBTNSPACING

    Const cBTNTHBCAPLEFT As Integer = (cLBLLEFT + cLBLWIDTH) - cBTNWIDTH
    Const cBTNFULLLEFT As Integer = (cLBLLEFT + cLBLWIDTH) - (cBTNWIDTH * 2)
    Const cBTNMSGLEFT As Integer = (cLBLLEFT + cLBLWIDTH) - (cBTNWIDTH * 3)

    Const cPNLWIDTH As Integer = cPICWIDTH + cPICSPACING
    Const cPNLHEIGHT As Integer = cPICHEIGHT + cLBLSPACING + cLBLHEIGHT + cPICSPACING
    Const cPNLSPACING As Integer = 20

    'Const cGRPSTARTWIDTH As Integer = cPNLSPACING + cPNLWIDTH + cPNLSPACING + cPNLWIDTH + cPNLSPACING
    'Const cGRPSTARTHEIGHT As Integer = cPNLSPACING + cPNLHEIGHT + cPNLSPACING

    'Const cFRMSPACINGWIDTH As Integer = 50
    'Const cFRMSPACINGHEIGHT As Integer = 110

#End Region             'Constants

#Region "Type Declares"

    Structure structPosition
        Dim Column As Integer
        Dim Row As Integer
        Public Sub New(ByVal intRow As Integer, ByVal intCol As Integer)
            Column = intCol
            Row = intRow
        End Sub
    End Structure

    Enum ClientModeType
        Thumbnail
        FullScreen
    End Enum

#End Region 'Type Declares

#Region "Fields"
    Private fldFullName As String
    Private fldUserName As String
    Private fldKeyName As String
    Private fldMachineName As String
    Private fldNetworkName As String

    Private fldNetworkStream As System.Net.Sockets.NetworkStream
    Private fldTcpClient As System.Net.Sockets.TcpClient

    Private fldPanel As Panel
    Private fldThumbPicturebox As Picturebox
    Private fldLabel As Label
    Private btnCaptureThumb As Button
    Private btnInitFullScreen As Button
    Private btnSendMsg As Button

    Private fldFullTabPage As TabPage
    Private fldFullPicturebox As Picturebox
    Private fldFullbtnClose As Button
    Private fldFullbtnMsg As Button
    Private fldFullbtnCapture As Button
    Private fldFullbtnPause As Button

    Private fldFullScreenPause As Boolean = False

    Private fldbtnCaptureThumbToolTip As System.Windows.Forms.ToolTip
    Private fldbtnInitFullScreenToolTip As System.Windows.Forms.ToolTip
    Private fldbtnSendMsgToolTip As System.Windows.Forms.ToolTip

    Private WithEvents fldMessage As Message
    Private fldMessageThread As Threading.Thread
    Private fldMessaging As Boolean = False

    Private fldImage As Image
    Private fldThumbTimestamp As DateTime
    Private fldPosition As structPosition
    Private fldImageMode As ClientModeType
    Private fldAutoCapture As Boolean

    Public Event MessageEnded(ByRef Client As Student)

#End Region     'Fields

#Region "Properties"

    Public Property Message() As Message
        Get
            Return fldMessage
        End Get
        Set(ByVal Value As Message)
            fldMessage = Value
        End Set
    End Property

    Public Property Messaging() As Boolean
        Get
            Return fldMessaging
        End Get
        Set(ByVal Value As Boolean)
            fldMessaging = Value
        End Set
    End Property

    Public ReadOnly Property FullScreenTabPage() As TabPage
        Get
            Return fldFullTabPage
        End Get
    End Property

    Public Property Picturebox() As Picturebox
        Get
            If fldImageMode = ClientModeType.Thumbnail Then
                Return fldThumbPicturebox
            ElseIf fldImageMode = ClientModeType.FullScreen Then
                Return fldFullPicturebox
            End If

        End Get
        Set(ByVal Value As Picturebox)
            If fldImageMode = ClientModeType.Thumbnail Then
                fldThumbPicturebox = Value
            ElseIf fldImageMode = ClientModeType.FullScreen Then
                fldFullPicturebox = Value
            End If
        End Set
    End Property

    Public Property Panel() As Panel
        Get
            Return fldPanel
        End Get
        Set(ByVal Value As Panel)
            fldPanel = Value
        End Set
    End Property

    Public ReadOnly Property FullScreenPauseButton() As Button
        Get
            Return fldFullbtnPause
        End Get
    End Property

    Public ReadOnly Property FullScreenCloseButton() As Button
        Get
            Return fldFullbtnClose
        End Get
    End Property

    Public ReadOnly Property FullScreenSendMessageButton() As Button
        Get
            Return fldFullbtnMsg
        End Get
    End Property

    Public ReadOnly Property FullScreenCaptureImageButton() As Button
        Get
            Return fldFullbtnCapture
        End Get
    End Property

    Public ReadOnly Property FullScreenButton() As Button
        Get
            Return Me.btnInitFullScreen
        End Get
    End Property

    Public ReadOnly Property CaptureThumbButton() As Button
        Get
            Return Me.btnCaptureThumb
        End Get
    End Property

    Public ReadOnly Property SendMessageButton() As Button
        Get
            Return Me.btnSendMsg
        End Get
    End Property

    Public Property FullScreenPause() As Boolean
        Get
            Return fldFullScreenPause
        End Get
        Set(ByVal Value As Boolean)
            fldFullScreenPause = Value
        End Set
    End Property

    Public ReadOnly Property FullName() As String
        Get
            Return fldFullName
        End Get
    End Property

    Public ReadOnly Property UserName() As String
        Get
            Return fldUserName
        End Get
    End Property

    Public Property KeyName() As String
        Get
            Return fldKeyName
        End Get
        Set(ByVal Value As String)
            fldKeyName = Value
        End Set
    End Property

    Public ReadOnly Property MachineName() As String
        Get
            Return fldMachineName
        End Get
    End Property

    Public ReadOnly Property NetworkName() As String
        Get
            Return fldNetworkName
        End Get
    End Property

    Public ReadOnly Property TcpClient() As System.Net.Sockets.TcpClient
        Get
            Return fldTcpClient
        End Get
    End Property

    Public ReadOnly Property NetworkStream() As System.Net.Sockets.NetworkStream
        Get
            Return fldNetworkStream
        End Get
    End Property

    Public Property ThumbTimestamp() As System.DateTime
        Get
            Return fldThumbTimestamp
        End Get

        Set(ByVal value As System.DateTime)
            fldThumbTimestamp = value
        End Set
    End Property

    Public Property Picture() As Image
        Get
            Return fldImage
        End Get
        Set(ByVal Value As Image)
            fldImage = Value
            If fldImageMode = ClientModeType.Thumbnail Then
                fldThumbPicturebox.Image = Value
            ElseIf fldImageMode = ClientModeType.FullScreen Then
                fldFullPicturebox.Image = Value
            End If
        End Set
    End Property

    Public Property Position() As structPosition
        Get
            Return fldPosition
        End Get
        Set(ByVal Value As structPosition)
            fldPosition = Value
        End Set
    End Property

    Public Property ImageMode() As ClientModeType
        Get
            Return fldImageMode
        End Get
        Set(ByVal Value As ClientModeType)
            fldImageMode = Value
        End Set
    End Property

    Public Property AutoCapture() As Boolean
        Get
            Return fldAutoCapture
        End Get
        Set(ByVal Value As Boolean)
            fldAutoCapture = Value
        End Set
    End Property

#End Region 'Properties

#Region "Methods"

    Public Sub New(ByVal FullName As String, ByVal TcpClient As System.Net.Sockets.TcpClient, ByVal Networkstream As System.Net.Sockets.NetworkStream)
        fldFullName = FullName
        Dim NameArray() As String = fldFullName.Split("\")
        fldNetworkName = NameArray(0)
        fldUserName = NameArray(1)
        fldMachineName = NameArray(NameArray.Length - 1)

        fldTcpClient = TcpClient
        fldNetworkStream = Networkstream

        fldMessage = New Message
    End Sub

    Public Sub InitThumbControls(ByVal NewPanel As Panel, ByVal NewKeyName As String, ByVal btnThumbImage As Image, ByVal btnFullImage As Image, ByVal btnSendMsgImage As Image)
        fldKeyName = NewKeyName
        fldImageMode = ClientModeType.Thumbnail

        fldPanel = NewPanel

        fldThumbPicturebox = New PictureBox
        fldLabel = New Label
        btnCaptureThumb = New Button
        btnInitFullScreen = New Button
        btnSendMsg = New Button

        With fldPanel
            .BringToFront()
            .Width = cPNLWIDTH
            .Height = cPNLHEIGHT
            .Tag = fldKeyName
            .Enabled = True
            .Visible = True
        End With

        fldPanel.Controls.Add(fldThumbPicturebox)
        fldPanel.Controls.Add(fldLabel)
        fldPanel.Controls.Add(btnCaptureThumb)
        fldPanel.Controls.Add(btnInitFullScreen)
        fldPanel.Controls.Add(btnSendMsg)

        With fldThumbPicturebox
            .BringToFront()
            .Name = "picThumb"
            .Width = cPICWIDTH
            .Height = cPICHEIGHT
            .Left = CPICLEFT
            .Top = CPICTOP
            .Image = fldImage
            .BackColor = System.Drawing.Color.DarkGray
            .SizeMode = PictureBoxSizeMode.StretchImage
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        With fldLabel
            .BringToFront()
            .Name = "lblName"
            .Text = fldUserName
            .TextAlign = ContentAlignment.MiddleLeft
            .Width = cLBLWIDTH
            .Height = cLBLHEIGHT
            .Left = cLBLLEFT
            .Top = cLBLTOP
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        With btnCaptureThumb
            .Name = "btnCaptureThumb"
            .BringToFront()
            .BackgroundImage = btnThumbImage
            '.Image = btnThumbImage
            '.ImageAlign = ContentAlignment.MiddleCenter
            .Width = cBTNWIDTH
            .Height = cBTNHEIGHT
            .FlatStyle = FlatStyle.Flat
            .Left = cBTNTHBCAPLEFT
            .Top = cBTNTOP
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        Me.fldbtnCaptureThumbToolTip = New ToolTip
        Me.fldbtnCaptureThumbToolTip.SetToolTip(Me.btnCaptureThumb, "Capture Thumbnail")
        Me.fldbtnCaptureThumbToolTip.Active = True

        With btnInitFullScreen
            .Name = "btnInitFullScreen"
            .BringToFront()
            .BackgroundImage = btnFullImage
            '.Image = btnFullImage
            '.ImageAlign = ContentAlignment.MiddleCenter
            .Width = cBTNWIDTH
            .Height = cBTNHEIGHT
            .FlatStyle = FlatStyle.Flat
            .Left = cBTNFULLLEFT
            .Top = cBTNTOP
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        Me.fldbtnInitFullScreenToolTip = New ToolTip
        Me.fldbtnInitFullScreenToolTip.SetToolTip(Me.btnInitFullScreen, "View FullScreen")
        Me.fldbtnInitFullScreenToolTip.Active = True

        With btnSendMsg
            .Name = "btnSendMsg"
            .BringToFront()
            .BackgroundImage = btnSendMsgImage
            '.Image = btnSendMsgImage
            '.ImageAlign = ContentAlignment.MiddleCenter
            .Width = cBTNWIDTH
            .Height = cBTNHEIGHT
            .FlatStyle = FlatStyle.Flat
            .Left = cBTNMSGLEFT
            .Top = cBTNTOP
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        Me.fldbtnSendMsgToolTip = New ToolTip
        Me.fldbtnSendMsgToolTip.SetToolTip(Me.btnSendMsg, "Send Message")
        Me.fldbtnSendMsgToolTip.Active = True

    End Sub

    Public Sub InitFullTab(ByVal NewTabPage As TabPage)
        Const cBTNFULLHEIGHT As Integer = 23
        Const cBTNFULLWIDTH As Integer = 75
        Const cBTNFULLLONGWIDTH As Integer = 110
        Const cBTNFULLSPACING As Integer = 8

        fldFullTabPage = NewTabPage
        With fldFullTabPage
            .Text = fldUserName & " Full"
            .Name = fldKeyName & "Full"
            .Enabled = True
            .Tag = fldKeyName
        End With

        fldFullPicturebox = New PictureBox
        With fldFullPicturebox
            .BringToFront()
            .Name = "picFull"
            .BackColor = System.Drawing.Color.Gray
            .Left = 8
            .Top = 8
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Top
            .Width = fldFullTabPage.Width
            .Height = fldFullTabPage.Height - 2 * cBTNFULLSPACING - cBTNHEIGHT
            .SizeMode = PictureBoxSizeMode.StretchImage
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        fldFullbtnClose = New Button
        With fldFullbtnClose
            .BringToFront()
            .Name = "btnFullClose"
            .Text = "Close"
            .Width = cBTNFULLWIDTH
            .Height = cBTNFULLHEIGHT
            .Left = fldFullPicturebox.Left
            .Top = fldFullPicturebox.Top + fldFullPicturebox.Height + cBTNFULLSPACING
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Left
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        fldFullbtnMsg = New Button
        With fldFullbtnMsg
            .BringToFront()
            .Name = "btnFullMsg"
            .Text = "Send Message"
            .Width = cBTNFULLLONGWIDTH
            .Height = cBTNFULLHEIGHT
            .Left = fldFullPicturebox.Left + (cBTNFULLWIDTH * 1) + cBTNFULLSPACING
            .Top = fldFullPicturebox.Top + fldFullPicturebox.Height + cBTNFULLSPACING
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Left
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        fldFullbtnCapture = New Button
        With fldFullbtnCapture
            .BringToFront()
            .Name = "btnFullCapture"
            .Text = "Capture Image"
            .Width = cBTNFULLLONGWIDTH
            .Height = cBTNFULLHEIGHT
            .Left = fldFullPicturebox.Left + cBTNFULLWIDTH + cBTNFULLLONGWIDTH + (cBTNFULLSPACING * 2)
            .Top = fldFullPicturebox.Top + fldFullPicturebox.Height + cBTNFULLSPACING
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Left
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        fldFullbtnPause = New Button
        With fldFullbtnPause
            .BringToFront()
            .Name = "FullbtnPause"
            .Text = "Pause"
            .Width = cBTNFULLWIDTH
            .Height = cBTNFULLHEIGHT
            .Left = fldFullPicturebox.Left + cBTNFULLWIDTH + (cBTNFULLLONGWIDTH * 2) + (cBTNFULLSPACING * 3)
            .Top = fldFullPicturebox.Top + fldFullPicturebox.Height + cBTNFULLSPACING
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Left
            .Enabled = True
            .Tag = fldKeyName
            .Visible = True
        End With

        fldFullTabPage.Controls.Add(fldFullPicturebox)
        fldFullTabPage.Controls.Add(fldFullbtnClose)
        fldFullTabPage.Controls.Add(fldFullbtnMsg)
        fldFullTabPage.Controls.Add(fldFullbtnCapture)
        fldFullTabPage.Controls.Add(fldFullbtnPause)

        fldImageMode = ClientModeType.FullScreen

    End Sub

    Public Sub SendMessage()
        fldMessageThread = New Threading.Thread(AddressOf Me.fldMessage.SendMessage)
        fldMessageThread.Start()
    End Sub

    Public Sub EndMessage()
        fldMessage.EndMessage()
        fldMessageThread.Abort()
    End Sub

    Private Sub Message_Sent() Handles fldMessage.MessageSent
        If fldMessage.CurrentMessageType = Message.Messagetype.Messagebox Then
            fldMessageThread.Abort()
            Exit Sub
        End If

        If fldMessage.TimeSpan.Equals(TimeSpan.Parse("0")) Then
            Exit Sub
        Else
            fldMessageThread.Sleep(fldMessage.TimeSpan)
        End If

    End Sub

    Private Sub Message_Ended() Handles fldMessage.MessageEnded
        fldMessaging = False
        RaiseEvent MessageEnded(Me)
    End Sub

#End Region 'Methods

End Class

