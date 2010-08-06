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

Public Class SWServer
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
    Friend WithEvents tmrMain As System.Windows.Forms.Timer
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents imgList As System.Windows.Forms.ImageList
    Friend WithEvents grpStudents As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents listOutput As System.Windows.Forms.ListBox
    Friend WithEvents tabMain As System.Windows.Forms.TabPage
    Friend WithEvents tabControl As System.Windows.Forms.TabControl
    Friend WithEvents tmrAutoCapture As System.Windows.Forms.Timer
    Friend WithEvents tabAutoCapture As System.Windows.Forms.TabPage
    Friend WithEvents grpACControls As System.Windows.Forms.GroupBox
    Friend WithEvents trkACInterval As System.Windows.Forms.TrackBar
    Friend WithEvents lblACInterval As System.Windows.Forms.Label
    Friend WithEvents btnACToggle As System.Windows.Forms.Button
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTabs As System.Windows.Forms.MenuItem
    Friend WithEvents mnnuMainTabView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAutoCaptureTabView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSavedImageViewer As System.Windows.Forms.MenuItem
    Friend WithEvents tabViewer As System.Windows.Forms.TabPage
    Friend WithEvents mnuViewer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreferences As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents tabNotepad As System.Windows.Forms.TabPage
    Friend WithEvents txtNotepad As System.Windows.Forms.RichTextBox
    Friend WithEvents btnNotepadSave As System.Windows.Forms.Button
    Friend WithEvents btnNotepadReload As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SWServer))
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuTabs = New System.Windows.Forms.MenuItem
        Me.mnnuMainTabView = New System.Windows.Forms.MenuItem
        Me.mnuViewer = New System.Windows.Forms.MenuItem
        Me.mnuAutoCaptureTabView = New System.Windows.Forms.MenuItem
        Me.mnuSavedImageViewer = New System.Windows.Forms.MenuItem
        Me.mnuOptions = New System.Windows.Forms.MenuItem
        Me.mnuPreferences = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuAbout = New System.Windows.Forms.MenuItem
        Me.imgList = New System.Windows.Forms.ImageList(Me.components)
        Me.tabControl = New System.Windows.Forms.TabControl
        Me.tabMain = New System.Windows.Forms.TabPage
        Me.grpStudents = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.listOutput = New System.Windows.Forms.ListBox
        Me.tabViewer = New System.Windows.Forms.TabPage
        Me.tabAutoCapture = New System.Windows.Forms.TabPage
        Me.grpACControls = New System.Windows.Forms.GroupBox
        Me.btnACToggle = New System.Windows.Forms.Button
        Me.lblACInterval = New System.Windows.Forms.Label
        Me.trkACInterval = New System.Windows.Forms.TrackBar
        Me.tmrAutoCapture = New System.Windows.Forms.Timer(Me.components)
        Me.tabNotepad = New System.Windows.Forms.TabPage
        Me.txtNotepad = New System.Windows.Forms.RichTextBox
        Me.btnNotepadSave = New System.Windows.Forms.Button
        Me.btnNotepadReload = New System.Windows.Forms.Button
        Me.tabControl.SuspendLayout()
        Me.tabMain.SuspendLayout()
        Me.grpStudents.SuspendLayout()
        Me.tabAutoCapture.SuspendLayout()
        Me.grpACControls.SuspendLayout()
        CType(Me.trkACInterval, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabNotepad.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrMain
        '
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuView, Me.mnuOptions, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
        Me.mnuFile.Text = "File"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 0
        Me.mnuExit.Text = "Exit"
        '
        'mnuView
        '
        Me.mnuView.Index = 1
        Me.mnuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTabs, Me.mnuSavedImageViewer})
        Me.mnuView.Text = "View"
        '
        'mnuTabs
        '
        Me.mnuTabs.Index = 0
        Me.mnuTabs.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnnuMainTabView, Me.mnuViewer, Me.mnuAutoCaptureTabView})
        Me.mnuTabs.Text = "Tabs"
        '
        'mnnuMainTabView
        '
        Me.mnnuMainTabView.Index = 0
        Me.mnnuMainTabView.Text = "Main"
        '
        'mnuViewer
        '
        Me.mnuViewer.Index = 1
        Me.mnuViewer.Text = "Viewer"
        '
        'mnuAutoCaptureTabView
        '
        Me.mnuAutoCaptureTabView.Index = 2
        Me.mnuAutoCaptureTabView.Text = "AutoCapture"
        '
        'mnuSavedImageViewer
        '
        Me.mnuSavedImageViewer.Index = 1
        Me.mnuSavedImageViewer.Text = "Saved Image Viewer"
        '
        'mnuOptions
        '
        Me.mnuOptions.Index = 2
        Me.mnuOptions.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPreferences})
        Me.mnuOptions.Text = "Options"
        '
        'mnuPreferences
        '
        Me.mnuPreferences.Index = 0
        Me.mnuPreferences.Text = "Preferences"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 3
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAbout})
        Me.mnuHelp.Text = "Help"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 0
        Me.mnuAbout.Text = "About"
        '
        'imgList
        '
        Me.imgList.ImageSize = New System.Drawing.Size(23, 23)
        Me.imgList.ImageStream = CType(resources.GetObject("imgList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgList.TransparentColor = System.Drawing.Color.Transparent
        '
        'tabControl
        '
        Me.tabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabControl.Controls.Add(Me.tabMain)
        Me.tabControl.Controls.Add(Me.tabViewer)
        Me.tabControl.Controls.Add(Me.tabNotepad)
        Me.tabControl.Controls.Add(Me.tabAutoCapture)
        Me.tabControl.Location = New System.Drawing.Point(0, 8)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(387, 370)
        Me.tabControl.TabIndex = 3
        '
        'tabMain
        '
        Me.tabMain.AutoScroll = True
        Me.tabMain.Controls.Add(Me.grpStudents)
        Me.tabMain.Controls.Add(Me.listOutput)
        Me.tabMain.Location = New System.Drawing.Point(4, 22)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.Size = New System.Drawing.Size(379, 344)
        Me.tabMain.TabIndex = 0
        Me.tabMain.Text = "Main"
        '
        'grpStudents
        '
        Me.grpStudents.Controls.Add(Me.Label2)
        Me.grpStudents.Controls.Add(Me.Label1)
        Me.grpStudents.Controls.Add(Me.PictureBox2)
        Me.grpStudents.Controls.Add(Me.PictureBox1)
        Me.grpStudents.Location = New System.Drawing.Point(8, 152)
        Me.grpStudents.Name = "grpStudents"
        Me.grpStudents.Size = New System.Drawing.Size(333, 171)
        Me.grpStudents.TabIndex = 4
        Me.grpStudents.TabStop = False
        Me.grpStudents.Text = "Students:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(175, 133)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(142, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label2.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(17, 133)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(141, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label1.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PictureBox2.Location = New System.Drawing.Point(175, 17)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(142, 111)
        Me.PictureBox2.TabIndex = 1
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PictureBox1.Location = New System.Drawing.Point(17, 17)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(141, 111)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'listOutput
        '
        Me.listOutput.HorizontalScrollbar = True
        Me.listOutput.Location = New System.Drawing.Point(8, 8)
        Me.listOutput.Name = "listOutput"
        Me.listOutput.Size = New System.Drawing.Size(333, 121)
        Me.listOutput.TabIndex = 3
        '
        'tabViewer
        '
        Me.tabViewer.Location = New System.Drawing.Point(4, 22)
        Me.tabViewer.Name = "tabViewer"
        Me.tabViewer.Size = New System.Drawing.Size(379, 344)
        Me.tabViewer.TabIndex = 2
        Me.tabViewer.Text = "Viewer"
        '
        'tabAutoCapture
        '
        Me.tabAutoCapture.AutoScroll = True
        Me.tabAutoCapture.Controls.Add(Me.grpACControls)
        Me.tabAutoCapture.Location = New System.Drawing.Point(4, 22)
        Me.tabAutoCapture.Name = "tabAutoCapture"
        Me.tabAutoCapture.Size = New System.Drawing.Size(379, 344)
        Me.tabAutoCapture.TabIndex = 1
        Me.tabAutoCapture.Text = "Auto Capture"
        '
        'grpACControls
        '
        Me.grpACControls.Controls.Add(Me.btnACToggle)
        Me.grpACControls.Controls.Add(Me.lblACInterval)
        Me.grpACControls.Controls.Add(Me.trkACInterval)
        Me.grpACControls.Location = New System.Drawing.Point(7, 7)
        Me.grpACControls.Name = "grpACControls"
        Me.grpACControls.Size = New System.Drawing.Size(373, 78)
        Me.grpACControls.TabIndex = 0
        Me.grpACControls.TabStop = False
        '
        'btnACToggle
        '
        Me.btnACToggle.Location = New System.Drawing.Point(293, 49)
        Me.btnACToggle.Name = "btnACToggle"
        Me.btnACToggle.Size = New System.Drawing.Size(63, 19)
        Me.btnACToggle.TabIndex = 5
        Me.btnACToggle.Text = "Start"
        '
        'lblACInterval
        '
        Me.lblACInterval.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblACInterval.Location = New System.Drawing.Point(13, 49)
        Me.lblACInterval.Name = "lblACInterval"
        Me.lblACInterval.Size = New System.Drawing.Size(274, 19)
        Me.lblACInterval.TabIndex = 4
        '
        'trkACInterval
        '
        Me.trkACInterval.LargeChange = 60
        Me.trkACInterval.Location = New System.Drawing.Point(7, 7)
        Me.trkACInterval.Maximum = 300
        Me.trkACInterval.Minimum = 5
        Me.trkACInterval.Name = "trkACInterval"
        Me.trkACInterval.Size = New System.Drawing.Size(360, 45)
        Me.trkACInterval.TabIndex = 3
        Me.trkACInterval.TickFrequency = 60
        Me.trkACInterval.Value = 5
        '
        'tmrAutoCapture
        '
        Me.tmrAutoCapture.Interval = 1000
        '
        'tabNotepad
        '
        Me.tabNotepad.Controls.Add(Me.btnNotepadReload)
        Me.tabNotepad.Controls.Add(Me.btnNotepadSave)
        Me.tabNotepad.Controls.Add(Me.txtNotepad)
        Me.tabNotepad.Location = New System.Drawing.Point(4, 22)
        Me.tabNotepad.Name = "tabNotepad"
        Me.tabNotepad.Size = New System.Drawing.Size(379, 344)
        Me.tabNotepad.TabIndex = 3
        Me.tabNotepad.Text = "Notepad"
        '
        'txtNotepad
        '
        Me.txtNotepad.AcceptsTab = True
        Me.txtNotepad.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNotepad.Location = New System.Drawing.Point(8, 8)
        Me.txtNotepad.Name = "txtNotepad"
        Me.txtNotepad.Size = New System.Drawing.Size(368, 296)
        Me.txtNotepad.TabIndex = 0
        Me.txtNotepad.Text = ""
        '
        'btnNotepadSave
        '
        Me.btnNotepadSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNotepadSave.Location = New System.Drawing.Point(320, 312)
        Me.btnNotepadSave.Name = "btnNotepadSave"
        Me.btnNotepadSave.Size = New System.Drawing.Size(48, 24)
        Me.btnNotepadSave.TabIndex = 1
        Me.btnNotepadSave.Text = "Save"
        '
        'btnNotepadReload
        '
        Me.btnNotepadReload.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNotepadReload.Location = New System.Drawing.Point(264, 312)
        Me.btnNotepadReload.Name = "btnNotepadReload"
        Me.btnNotepadReload.Size = New System.Drawing.Size(48, 24)
        Me.btnNotepadReload.TabIndex = 2
        Me.btnNotepadReload.Text = "Reload"
        '
        'SWServer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(388, 377)
        Me.Controls.Add(Me.tabControl)
        Me.Menu = Me.mnuMain
        Me.Name = "SWServer"
        Me.Text = "SWServer"
        Me.tabControl.ResumeLayout(False)
        Me.tabMain.ResumeLayout(False)
        Me.grpStudents.ResumeLayout(False)
        Me.tabAutoCapture.ResumeLayout(False)
        Me.grpACControls.ResumeLayout(False)
        CType(Me.trkACInterval, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabNotepad.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constants"

    Const cPNLWIDTH As Integer = Student.cPICWIDTH + Student.cPICSPACING
    Const cPNLHEIGHT As Integer = Student.cPICHEIGHT + Student.cLBLSPACING + Student.cLBLHEIGHT + Student.cPICSPACING
    Const cPNLSPACING As Integer = 20

    Const cGRPSTARTWIDTH As Integer = cPNLSPACING + cPNLWIDTH + cPNLSPACING + cPNLWIDTH + cPNLSPACING
    Const cGRPSTARTHEIGHT As Integer = cPNLSPACING + cPNLHEIGHT + cPNLSPACING

    Const cFRMSPACINGWIDTH As Integer = 50
    Const cFRMSPACINGHEIGHT As Integer = 110

    Const cACSTARTTOP As Integer = 100

    Const cACGRPMAINWIDTH As Integer = 190
    Const cACGRPMAINHEIGHT As Integer = 30

    Const cACLBLNAMEWIDTH As Integer = cACGRPMAINWIDTH - 50
    Const cACLBLNAMEHEIGHT As Integer = 18
    Const cACLBLNAMETOP As Integer = 10
    Const cACLBLNAMELEFT As Integer = 10

    Const cACCHKAUTOCAPTUREWIDTH As Integer = 20
    Const cACCHKAUTOCAPTUREHEIGHT As Integer = 18
    Const cACCHKAUTOCAPTURETOP As Integer = 10
    Const cACCHKAUTOCAPTURELEFT As Integer = cACLBLNAMEWIDTH + 25

    Const cACROWSPACING As Integer = 10
    Const cACCOLUMNSPACING As Integer = 15

    Const cSAVEDNOTESFILEPATH As String = "\Saved Notes.txt"    'Leading slash required

    'Stay On Top

    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40

    Enum TabPageType
        FullScreen
    End Enum

    Enum ButtonImages
        SaveImage = 0
        FullScreen = 1
        SendMessage = 2
        ExitFullScreen = 3
        EndMessage = 4
    End Enum

#End Region     'Constants

#Region "Declares"

    Structure ACvars
        Dim Interval As Integer
        Dim Enabled As Boolean
        Dim TickCount As Integer

        Public Sub New(ByVal intInterval As Integer, ByVal bolEnabled As Boolean)
            Interval = intInterval
            Enabled = bolEnabled
        End Sub
    End Structure

    Private StudentCollection As New Collection 'Collection of Structure Student
    Private Dimensions As Student.structPosition
    Private DimensionsPlaceHolder(,) As Boolean
    Private AutoCapture As New ACvars(5000, False)

    Private WithEvents MessageOptions As MessageOptions

    Private cnmnuPicThumb As New ContextMenu
    Private cnmnuTabControl As New ContextMenu
    Private cnmnuTabPage As New ContextMenu

    'Forms
    Dim frmPreferences As Form = Nothing
    Dim frmSavedImageViewer As Form = Nothing
    Dim frmEnlargeForm As Form = Nothing

#End Region         'Declares

    Private Sub SWServer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dimensions = New Student.structPosition(5, 6)

        If Dimensions.Column = 0 Then
            Dimensions.Column = 1
        End If
        If Dimensions.Row = 0 Then
            Dimensions.Row = 1
        End If
        ReDim Preserve DimensionsPlaceHolder(Dimensions.Row - 1, Dimensions.Column - 1)
        DimensionsPlaceHolder.Initialize()

        Me.listOutput.Top = 8
        Me.listOutput.Left = cPNLSPACING
        Me.listOutput.Width = cGRPSTARTWIDTH
        Me.listOutput.Height = 132

        Me.grpStudents.Top = Me.listOutput.Top + Me.listOutput.Height + cPNLSPACING
        Me.grpStudents.Left = cPNLSPACING
        Me.grpStudents.Width = cGRPSTARTWIDTH
        Me.grpStudents.Height = cGRPSTARTHEIGHT

        Me.Width = Me.grpStudents.Left + cGRPSTARTWIDTH + cFRMSPACINGWIDTH
        Me.Height = Me.grpStudents.Top + cGRPSTARTHEIGHT + cFRMSPACINGHEIGHT

        Me.MinimumSize = New Size(New Point(Me.Width, Me.Height))

        'following are random init constants derived from designer
        With Me.grpACControls
            .Top = 8
            .Left = 8
            .Width = 448
            .Height = 90
        End With
        With Me.trkACInterval
            .Top = 8
            .Left = 8
            .Width = 432
            .Height = 56
        End With
        With Me.lblACInterval
            .Width = 328
            .Height = 21
            .Top = 57
            .Left = 16
            .Text = ""
        End With
        With Me.btnACToggle
            .Width = 75
            .Height = 21
            .Top = 51
            .Left = Me.trkACInterval.Left + Me.trkACInterval.Width - Me.btnACToggle.Width
        End With

        Me.lblACInterval.BringToFront()
        Me.btnACToggle.BringToFront()

        Application.DoEvents()

        Me.Show()
        Initialize()
        InitContextMenus()

        Me.tmrMain.Enabled = True
    End Sub

    Private Sub tmrMain_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMain.Tick
        tmrMain.Enabled = False
        Dim NewStudents() As Student
        If PendingClients() = True Then
            NewStudents = ConnectWaitingClients()

            For Each Client As Student In NewStudents
                AddStudent(Client)
            Next
        End If

        For Each Client As Student In StudentCollection
            Try
                If Client.NetworkStream.DataAvailable = True Then
                    GetData(Client)
                Else
                    If Now.Subtract(Client.ThumbTimestamp).Seconds > 5 Then
                        If SendTypeByte(Client, "P") Then
                            Client.ThumbTimestamp = Now
                        Else
                            Throw New Exception
                        End If
                    End If
                End If

            Catch ex As Exception
                Disconnect(client)
            End Try
        Next

        tmrMain.Enabled = True
    End Sub

#Region "Network Events"

    Private Sub CapturePicture(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Dim KeyName As String = GetKeyName(sender)
        Dim client As Student

        If KeyName <> "" Then
            client = StudentCollection(KeyName)
        End If

        SaveImage(client)

        Display("Screenshot of " & client.UserName & " Captured")
    End Sub

    Private Sub InitFullScreen(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Dim Client As Student

        Dim KeyName As String = GetKeyName(sender)

        If KeyName = "" Then
            Exit Sub
        Else
            Client = StudentCollection(KeyName)
        End If

        If AddTab(Client, TabPageType.FullScreen) Then
            SendTypeByte(Client, "F")
            SendTypeByte(Client, "R")

            Client.FullScreenButton.BackgroundImage = imgList.Images(3)


            Dim RemoveEventHandler As [Delegate] = New EventHandler(AddressOf InitFullScreen)
            RemoveHandler Client.FullScreenButton.Click, RemoveEventHandler

            Dim AddEventHandler As [Delegate] = New EventHandler(AddressOf EndFullScreenEventCall)
            AddHandler Client.FullScreenButton.Click, AddEventHandler

            AddHandler Client.FullScreenPauseButton.Click, New EventHandler(AddressOf PauseFullScreen)

            Client.ImageMode = Student.ClientModeType.FullScreen
        End If

    End Sub

    Private Sub EndFullScreenEventCall(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim KeyName As String = GetKeyName(sender)
        Dim client As Student

        If KeyName <> "" Then
            client = StudentCollection(KeyName)
            EndFullScreen(client)
        End If
    End Sub

    Private Sub PauseFullScreen(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim KeyName As String = GetKeyName(sender)
        Dim Client As Student

        If KeyName = "" Then
            Exit Sub
        Else
            Client = StudentCollection(KeyName)
        End If

        If Client.FullScreenPause = False Then
            Client.FullScreenPause = True
            Client.FullScreenPauseButton.Text = "Resume"
            For Each mnuItem As MenuItem In Client.Picturebox.ContextMenu.MenuItems
                If mnuItem.Text = "Pause" Then
                    mnuItem.Text = "Resume"
                    Exit For
                End If
            Next
        ElseIf Client.FullScreenPause = True Then
            Client.FullScreenPause = False
            SendTypeByte(Client, "R")
            Client.FullScreenPauseButton.Text = "Pause"
            For Each mnuItem As MenuItem In Client.Picturebox.ContextMenu.MenuItems
                If mnuItem.Text = "Resume" Then
                    mnuItem.Text = "Pause"
                End If
            Next
        End If

    End Sub

    Private Sub DisplayMode(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim KeyName As String = GetKeyName(sender)
        Dim client As Student

        If KeyName <> "" Then
            client = StudentCollection(KeyName)
            Display(client.UserName & " Mode: " & client.ImageMode.ToString)
        End If
    End Sub

    Private Sub OpenMessage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim KeyName As String = GetKeyName(sender)
        Dim client As Student

        If KeyName = "" Then
            Exit Sub
        Else
            client = StudentCollection(KeyName)
        End If

        If client.Messaging = False Then
            ShowMessageOptions(client)
        ElseIf client.Messaging = True Then
            client.EndMessage()
            client.Messaging = False
            client.SendMessageButton.BackgroundImage = Me.imgList.Images(ButtonImages.SendMessage)
            client.FullScreenSendMessageButton.Text = "Send Message"
        End If

    End Sub

    Private Sub MessageOptions_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MessageOptions.Closing
        Dim frmMessageOptions As MessageOptions
        If sender.GetType.Equals(GetType(MessageOptions)) Then
            frmMessageOptions = sender
        Else
            Exit Sub
        End If

        Dim Client As Student = StudentCollection(frmMessageOptions.KeyName)

        If Client.Message.Initialized = False Then
            Display("Message Not Sent")
            Exit Sub
        End If

        If Client.Message.CurrentMessageType = Message.Messagetype.FullScreen Then
            Client.Messaging = True
            Client.SendMessageButton.BackgroundImage = Me.imgList.Images(ButtonImages.EndMessage)
            Client.FullScreenSendMessageButton.Text = "End Message"
        End If

        Client.SendMessage()

    End Sub

    Private Sub ClientMessage_Ended(ByRef Client As Student)
        Display("Message Sent")
        Client.SendMessageButton.BackgroundImage = Me.imgList.Images(ButtonImages.SendMessage)
        Client.FullScreenSendMessageButton.Text = "Send Message"
    End Sub

    Private Sub SelectTabPage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If sender.GetType.Equals(GetType(MenuItem)) Then
            Dim MenuItem As MenuItem = sender
            SelectTab(MenuItem.Text)
        End If
    End Sub

    Private Sub EnlargePicturebox(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Exit Sub
        End If

        Dim picBox As PictureBox = sender
        Dim KeyName As String = picBox.Tag
        Dim Student As Student = StudentCollection(KeyName)

        If Student.ImageMode = Student.ClientModeType.FullScreen Then
            PauseFullScreen(picBox, e.Empty)
        End If

        Me.ShowEnlargeForm(picBox.Image)

    End Sub

#End Region     'Network Events

#Region "Network Utiliites"

    Public Sub SetMessage(ByVal Keyname As String, ByVal NewMessage As Message)
        Dim Client As Student = StudentCollection(Keyname)
        Client.Message = NewMessage
    End Sub

    Private Sub GetData(ByVal client As Student)
        Dim TypeChar As Char = GetDataType(client.NetworkStream)

        Select Case TypeChar
            Case "T"
                If client.ImageMode = Student.ClientModeType.FullScreen Then
                    client.ImageMode = Student.ClientModeType.Thumbnail
                End If
                client.Picture = GetStreamImage(client)
            Case "F"
                If client.FullScreenPause = True Then
                    Exit Select
                End If

                If client.ImageMode = Student.ClientModeType.Thumbnail Then
                    client.ImageMode = Student.ClientModeType.FullScreen
                End If
                client.Picture = GetStreamImage(client)

                SendTypeByte(client, "R")

            Case "P"

            Case Else
                FlushNetworkStream(client)
        End Select

    End Sub

    Private Sub Disconnect(ByRef client As Student)
        If client.ImageMode = Student.ClientModeType.FullScreen Then
            EndFullScreen(client)
        End If

        client.Picturebox.Image = Nothing
        client.FullScreenButton.Enabled = False
        client.SendMessageButton.Enabled = False
        client.CaptureThumbButton.Enabled = False

        StudentCollection.Remove(client.KeyName)
        Display(client.UserName & " disconnected.")
        DimensionsPlaceHolder(client.Position.Row - 1, client.Position.Column - 1) = False
    End Sub

    Private Sub EndFullScreen(ByRef client As Student)
        On Error Resume Next

        client.ImageMode = Student.ClientModeType.Thumbnail

        SendTypeByte(client, "T")

        Dim FullTabPage As TabPage
        For Each Tab As TabPage In tabControl.TabPages
            If Tab.Tag = client.KeyName Then
                FullTabPage = Tab
            End If
        Next

        client.FullScreenButton.BackgroundImage = imgList.Images(1)

        Dim RemoveEventHandler As [Delegate] = New EventHandler(AddressOf EndFullScreenEventCall)
        RemoveHandler client.FullScreenButton.Click, RemoveEventHandler

        Dim AddEventHandler As [Delegate] = New EventHandler(AddressOf InitFullScreen)
        AddHandler client.FullScreenButton.Click, AddEventHandler

        Dim View As MenuItem = GetMenuItem(mnuMain, "View")
        Dim Tabs As MenuItem = GetMenuItem(View, "Tabs")

        For Each itmMnu As MenuItem In Tabs.MenuItems
            If itmMnu.Text = FullTabPage.Text Then
                Tabs.MenuItems.Remove(itmMnu)
            End If
        Next

        tabControl.TabPages.RemoveAt(tabControl.TabPages.IndexOf(FullTabPage))
        tabControl.SelectedTab = Me.tabMain
    End Sub

    Private Sub FlushNetworkStream(ByVal client As Student)
        Dim buffer(client.TcpClient.ReceiveBufferSize - 1) As Byte
        Do Until client.NetworkStream.DataAvailable = False
            client.NetworkStream.Read(buffer, 0, client.TcpClient.ReceiveBufferSize)
        Loop
    End Sub

#End Region 'Network Utiliites

#Region "Interface Methods"

    Private Sub AddStudent(ByVal NewStudent As Student)
        Dim NewPosition As Student.structPosition = FirstOpenPosition()
        Static NewColumns As Integer = 0

        If NewPosition.Row = 0 And NewPosition.Column = 0 Then
            Dim DiaADDResult As DialogResult = MessageBox.Show(NewStudent.UserName & " is trying to connect but" & vbNewLine & "you have exceeded your entered dimensions." & vbNewLine & vbNewLine & "Would you like to force add the new client?", "Exceeded Dimensions", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            If DiaADDResult = DialogResult.No Then
                Exit Sub
            Else
                ReDim Preserve Me.DimensionsPlaceHolder(Dimensions.Row - 1, Dimensions.Column + NewColumns)
                NewColumns += 1
                Dimensions.Column += 1
            End If
        End If

        Dim NewPanel As New Panel

        Me.grpStudents.Controls.Add(NewPanel)

        'Set Panel Position
        NewStudent.Position = FirstOpenPosition()
        If NewStudent.Position.Column = 1 And NewStudent.Position.Row = 1 Then
            With NewPanel
                .BringToFront()
                .Left = cPNLSPACING
                .Top = cPNLSPACING
            End With
        Else
            Dim ReferenceStudent As Student = Me.GetReferenceStudent(NewStudent.Position)

            If NewStudent.Position.Column = 1 Then 'new row
                With NewPanel
                    .BringToFront()
                    .Left = ReferenceStudent.Panel.Left
                    .Top = ReferenceStudent.Panel.Top + cPNLHEIGHT + cPNLSPACING
                End With

                Dim NewTotalHeight As Integer = (NewStudent.Position.Row * cPNLHEIGHT) + (NewStudent.Position.Row + 1) * (cPNLSPACING)
                If NewTotalHeight > Me.grpStudents.Height Then
                    Me.grpStudents.Height += cPNLHEIGHT + cPNLSPACING
                    If Me.grpStudents.Top + Me.grpStudents.Height > Me.Height And Me.grpStudents.Top + Me.grpStudents.Height + cFRMSPACINGHEIGHT < Screen.PrimaryScreen.Bounds.Height Then
                        Me.Height = Me.grpStudents.Top + Me.grpStudents.Height + cFRMSPACINGHEIGHT
                    End If
                End If

            Else 'new column
                With NewPanel
                    .BringToFront()
                    .Left = ReferenceStudent.Panel.Left + cPNLWIDTH + cPNLSPACING
                    .Top = ReferenceStudent.Panel.Top
                End With

                Dim NewTotalWidth As Integer = (NewStudent.Position.Column * cPNLWIDTH) + (NewStudent.Position.Column + 1) * (cPNLSPACING)
                If NewTotalWidth > Me.grpStudents.Width Then
                    Me.grpStudents.Width += cPNLWIDTH + cPNLSPACING
                    If Me.grpStudents.Left + Me.grpStudents.Width > Me.Width And Me.grpStudents.Left + Me.grpStudents.Width + cFRMSPACINGWIDTH < Screen.PrimaryScreen.Bounds.Width Then
                        Me.Width = Me.grpStudents.Left + Me.grpStudents.Width + cFRMSPACINGWIDTH
                    End If
                End If
            End If
        End If

        NewStudent.InitThumbControls(NewPanel, GetKeyName(NewStudent.FullName), Me.imgList.Images(0), Me.imgList.Images(1), Me.imgList.Images(2))

        NewStudent.Picturebox.ContextMenu = Me.cnmnuPicThumb
        AddHandler NewStudent.FullScreenButton.Click, New EventHandler(AddressOf Me.InitFullScreen)
        AddHandler NewStudent.CaptureThumbButton.Click, New EventHandler(AddressOf Me.CapturePicture)
        AddHandler NewStudent.SendMessageButton.Click, New EventHandler(AddressOf OpenMessage)
        AddHandler NewStudent.MessageEnded, New Student.MessageEndedEventHandler(AddressOf ClientMessage_Ended)
        DimensionsPlaceHolder(NewStudent.Position.Row - 1, NewStudent.Position.Column - 1) = True

        Display(NewStudent.UserName & " has connected.")

        Try
            StudentCollection.Add(NewStudent, NewStudent.KeyName)
        Catch ex As Exception
            Display("Could not add " & NewStudent.UserName & " to Collection at " & System.DateTime.Now.TimeOfDay.ToString & ". (AddFirstStudent())")
        End Try

    End Sub

    Private Function AddTab(ByRef Client As Student, ByVal Type As TabPageType) As Boolean
        Dim NewTabPage As New TabPage

        Select Case Type
            Case TabPageType.FullScreen
                If Client.ImageMode = Student.ClientModeType.FullScreen Then
                    Display("Cannot open two FullScreen tabs of same student")
                    Return False
                End If

                Client.Picturebox.Image = Nothing

                Client.ImageMode = Student.ClientModeType.FullScreen
                Client.InitFullTab(NewTabPage)

                AddHandler Client.FullScreenCloseButton.Click, New EventHandler(AddressOf EndFullScreenEventCall)
                AddHandler Client.FullScreenSendMessageButton.Click, New EventHandler(AddressOf OpenMessage)
                AddHandler Client.FullScreenCaptureImageButton.Click, New EventHandler(AddressOf CapturePicture)
                AddHandler Client.Picturebox.MouseUp, New MouseEventHandler(AddressOf EnlargePicturebox)

                Dim CMnuPicFull As ContextMenu = InitFullPicContextMenu()
                Client.Picturebox.ContextMenu = CMnuPicFull

                Client.Picturebox.Cursor = System.Windows.Forms.Cursors.Hand
        End Select

        NewTabPage.ContextMenu = cnmnuTabPage
        tabControl.TabPages.Add(NewTabPage)
        tabControl.SelectedTab = NewTabPage

        Dim NewMenuItem As New MenuItem(NewTabPage.Text, New EventHandler(AddressOf SelectTabPage))

        Dim FirstLevelMenu As MenuItem = GetMenuItem(mnuMain, "View")
        Dim Tabs As MenuItem = GetMenuItem(FirstLevelMenu, "Tabs")
        Tabs.MenuItems.Add(NewMenuItem)

        Return True

    End Function

    Public Sub Display(ByVal strOutput As String)
        Dim SplitChars() As Char = {vbCrLf, vbNewLine}
        Dim splitoutput() As String = strOutput.Split(SplitChars)

        For i As Integer = 0 To splitoutput.Length - 1
            splitoutput(i).Trim(SplitChars)
            Me.listOutput.Items.Add(splitoutput(i))
        Next
        listOutput.SelectedIndex = listOutput.Items.Count - 1

    End Sub

    Private Function InitFullPicContextMenu() As ContextMenu
        Dim CMnuPicFull As New ContextMenu

        CMnuPicFull.MenuItems.Add("Pause", New EventHandler(AddressOf PauseFullScreen))
        CMnuPicFull.MenuItems.Add("-")
        CMnuPicFull.MenuItems.Add("Capture Image", New EventHandler(AddressOf CapturePicture))
        CMnuPicFull.MenuItems.Add("Send Message", New EventHandler(AddressOf OpenMessage))
        CMnuPicFull.MenuItems.Add("End FullScreen", New EventHandler(AddressOf EndFullScreenEventCall))
        CMnuPicFull.MenuItems.Add("-")
        CMnuPicFull.MenuItems.Add("Close", New EventHandler(AddressOf EndFullScreenEventCall))

        Return CMnuPicFull
    End Function

    Private Sub InitContextMenus()
        Me.cnmnuPicThumb.MenuItems.Add("Capture Thumbnail", New EventHandler(AddressOf CapturePicture))
        Me.cnmnuPicThumb.MenuItems.Add("View FullScreen", New EventHandler(AddressOf InitFullScreen))
        Me.cnmnuPicThumb.MenuItems.Add("Send Message", New EventHandler(AddressOf OpenMessage))
        Me.cnmnuPicThumb.MenuItems.Add("-")
        Me.cnmnuPicThumb.MenuItems.Add("Current Mode", New EventHandler(AddressOf DisplayMode))

        Me.cnmnuTabControl.MenuItems.Add("Close", New EventHandler(AddressOf EndFullScreenEventCall))
        tabControl.ContextMenu = Me.cnmnuTabControl

        Me.tabMain.ContextMenu = cnmnuTabPage
    End Sub

    Private Sub ShowImageViewer()
        Dim SavedImageViewer As Form = New SavedImageViewer
        If Not frmSavedImageViewer Is Nothing Then
            frmSavedImageViewer.Dispose()
        End If
        frmSavedImageViewer = SavedImageViewer
        frmSavedImageViewer.Owner = Me
        frmSavedImageViewer.Show()
    End Sub

    Private Sub ShowMessageOptions(ByRef client As Student)
        MessageOptions = New MessageOptions(client)
        MessageOptions.Owner = Me
        MessageOptions.Show()
    End Sub

    Private Sub ShowPreferences()
        Dim Preferences As Form = New Preferences
        If Not frmPreferences Is Nothing Then
            frmPreferences.Dispose()
        End If
        frmPreferences = Preferences
        frmPreferences.Owner = Me
        frmPreferences.Show()
    End Sub

    Private Sub ShowEnlargeForm(ByVal Image As Image)
        Dim EnlargeForm As New EnlargeForm(Image)
        If Not frmEnlargeForm Is Nothing Then
            frmEnlargeForm.Dispose()
        End If
        frmEnlargeForm = EnlargeForm
        frmEnlargeForm.Owner = Me
        frmEnlargeForm.Show()
    End Sub

#End Region         'Interface Methods

#Region "Interface Events"

    Private Sub tabControl_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControl.SelectedIndexChanged
        If tabControl.SelectedTab.Name = Me.tabAutoCapture.Name Then
            LoadAutoCapture()
            UpdateInterfaceInterval(sender, e)
        ElseIf tabControl.SelectedTab.Name = Me.tabViewer.Name Then
            tabControl.SelectedTab = Me.tabMain
            ShowImageViewer()
        End If
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        For Each Client As Student In StudentCollection
            If Client.ImageMode = Student.ClientModeType.FullScreen Then
                EndFullScreen(Client)
            End If
        Next
        End
    End Sub

    Private Sub mnuPreferences_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreferences.Click
        ShowPreferences()
    End Sub

    Private Sub mnnuMainTabView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnnuMainTabView.Click
        SelectTab("Main")
    End Sub

    Private Sub mnuAutoCaptureTabView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAutoCaptureTabView.Click
        SelectTab("Auto Capture")
    End Sub

    Private Sub mnuViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewer.Click
        SelectTab("Viewer")
    End Sub

    Private Sub mnuSavedImageViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSavedImageViewer.Click
        Me.ShowImageViewer()
    End Sub

    Private Sub mnuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        MessageBox.Show("Homepage: http://www.screenwatcher.net" & vbNewLine & vbNewLine & "ScreenWatcher, composed of the client application, server application, and Class Designer, is © (copyright) of and was originally written by John Moore and André Wiggins." & vbNewLine & vbNewLine & "ScreenWatcher has been released under the Reciprocal Public License. You should have received a copy of this license with the source code. If not, you can find a copy of it at http://www.programiscellaneous.com/programming-projects/screenwatcher/licensing-copyright/." & vbNewLine & vbNewLine & "Additionally, the license template can be found at http://www.opensource.org/licenses/rpl1.5.txt." & vbNewLine & vbNewLine & "ScreenWatcher is protected under U.S. law from infringement of any rights not granted to the end user of the program by this license, including the acquisition of profit from redistribution of the original source or any variant of it except as permitted by the license itself.")
    End Sub

    Private Sub SWServer_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        For Each Client As Student In StudentCollection
            If Client.ImageMode = Student.ClientModeType.FullScreen Then
                EndFullScreen(Client)
            End If
        Next
        SaveNotepad()
    End Sub

#End Region 'Interface Events

#Region "Interface Utilities"

    Private Function GetMenuItem(ByVal Menu As Menu, ByVal Name As String) As MenuItem
        For Each itmMenu As MenuItem In Menu.MenuItems
            If itmMenu.Text = Name Then
                Return itmMenu
            End If
        Next
    End Function

    Public Sub SaveImage(ByVal client As Student)
        Dim strDirectory As String = Application.StartupPath & "\Saved Images\" & client.UserName & "\" & client.ImageMode.ToString & "\" & Now.Date.Day & "-" & Now.Date.Month & "-" & Now.Date.Year & "\" & IIf(client.AutoCapture = True, "AutoCapture", "Manual")
        Dim strFileName As String = Now.TimeOfDay.Hours & "-" & IIf(Now.TimeOfDay.Minutes.ToString.Length = 1, "0" & Now.TimeOfDay.Minutes, Now.TimeOfDay.Minutes) & "-" & IIf(Now.TimeOfDay.Seconds.ToString.Length = 1, "0" & Now.TimeOfDay.Seconds, Now.TimeOfDay.Seconds) & "-" & Now.TimeOfDay.Milliseconds
        If System.IO.Directory.Exists(strDirectory) = False Then
            MkDir(strDirectory)
        End If
        client.Picture.Save(strDirectory & "\" & strFileName & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
    End Sub

    Private Function GetReferenceStudent(ByVal CurrentPosition As Student.structPosition) As Student
        Dim RefStudentPosition As Student.structPosition
        If CurrentPosition.Column = 1 Then
            RefStudentPosition.Column = CurrentPosition.Column
            RefStudentPosition.Row = CurrentPosition.Row - 1
        Else
            RefStudentPosition.Column = CurrentPosition.Column - 1
            RefStudentPosition.Row = CurrentPosition.Row
        End If

        For Each newStudent As Student In StudentCollection
            If newStudent.Position.Column = RefStudentPosition.Column And newStudent.Position.Row = RefStudentPosition.Row Then
                Return newStudent
            End If
        Next

    End Function

    Private Function GetKeyName(ByVal FullName As String) As String
        Dim Instances As Integer = NumOfInstances(StudentCollection, FullName)
        If Instances = 0 Then
            Return FullName
        Else
            Return FullName & Instances.ToString
        End If

    End Function

    Private Function GetKeyName(ByVal sender As System.Object)
        Dim strReturn As String

        If sender.GetType.Equals(GetType(Button)) Then
            Dim senderButton As Button = sender
            strReturn = senderButton.Tag

        ElseIf sender.GetType.Equals(GetType(PictureBox)) Then
            Dim senderPicturebox As PictureBox = sender
            strReturn = senderPicturebox.Tag

        ElseIf sender.GetType.Equals(GetType(MenuItem)) Then
            Dim mnuItem As MenuItem = sender

            Dim ParentMenu As ContextMenu = mnuItem.GetContextMenu
            Dim ParentControl As Control = ParentMenu.SourceControl

            If ParentControl.GetType.Equals(GetType(PictureBox)) Then
                Dim Picturebox As PictureBox = ParentControl
                strReturn = Picturebox.Tag
            ElseIf ParentControl.GetType.Equals(GetType(TabControl)) Then
                Dim TabControl As TabControl = ParentControl
                Dim SelectedTab As TabPage = TabControl.SelectedTab

                If SelectedTab.Tag = "MAIN" Then
                    strReturn = ""
                Else
                    strReturn = SelectedTab.Tag
                End If
            End If
        Else
            strReturn = ""
        End If

        Return strReturn
    End Function

    Private Sub SelectTab(ByVal Name As String)
        For Each Tab As TabPage In tabControl.TabPages
            If Tab.Text = Name Then
                tabControl.SelectedTab = Tab
                Exit Sub
            End If
        Next
    End Sub

    Private Function FirstOpenPosition() As Student.structPosition
        For i As Integer = 0 To Dimensions.Row - 1
            For j As Integer = 0 To Dimensions.Column - 1
                If DimensionsPlaceHolder(i, j) = False Then
                    Return New Student.structPosition(i + 1, j + 1)
                End If
            Next
        Next

        Return New Student.structPosition(0, 0)
    End Function

    Private Function NumOfInstances(ByVal Collection As Collection, ByVal FullName As String) As Integer
        Dim count As Integer = 0

        For Each TestStudent As Student In Collection
            If FullName = TestStudent.FullName Then
                count += 1
            End If
        Next

        Return count
    End Function

    Private Function PrintArray(ByVal array As Byte()) As String
        Dim strString As String
        Dim Percent As Integer
        Dim Fraction As Double
        Dim LastPercent As Integer

        For i As Integer = 0 To array.Length - 1
            strString = strString & ChrW(array(i))

            Fraction = i / (array.Length - 1)
            Percent = Math.Floor(Fraction * 100)
            If Percent > 0 And Percent Mod 10 = 0 And Percent <> LastPercent Then
                LastPercent = Percent
            End If
            Application.DoEvents()
        Next

        Return strString
    End Function

#End Region 'Interface Utilities

#Region "AutoCapture"

    Private Sub LoadAutoCapture()
        Dim intStartLeft As Integer = Int((tabControl.Width - (((cACGRPMAINWIDTH + cACCOLUMNSPACING) * (IIf(StudentCollection.Count >= Dimensions.Column, Dimensions.Column - 1, IIf((StudentCollection.Count - 1) <= 2, 1, StudentCollection.Count - 1)))) + cACGRPMAINWIDTH)) / 2) 'finds the spacing of the first column from the side of the form by subtracting the maximum achievable width of the content from either the tabMainControl's current width, or, if that has not been reached yet (not enough clients connected), from the width it is at now and then divides it by two; in short, this finds the correct spacing for the first column from the side of the form
        Dim intLeft As Integer = intStartLeft
        Dim intTop As Integer = cACSTARTTOP
        Dim i As Integer
RemoveControls:
        For Each control As Control In Me.tabAutoCapture.Controls
            Try
                If control.GetType.Equals(GetType(GroupBox)) Then
                    If control.Tag = 1 Then
                        Me.tabAutoCapture.Controls.Remove(control)
                    End If
                End If
            Catch ex As Exception
                GoTo RemoveControls
            End Try
        Next
        For Each student As Student In StudentCollection
            Dim grpMain As New GroupBox
            Dim lblName As New Label
            Dim chkAutoCapture As New CheckBox
            With grpMain
                .Width = cACGRPMAINWIDTH
                .Height = cACGRPMAINHEIGHT
                .Left = intLeft
                .Top = intTop
                .Tag = 1
                .Show()
            End With
            With lblName
                .Width = cACLBLNAMEWIDTH
                .Height = cACLBLNAMEHEIGHT
                .Top = cACLBLNAMETOP
                .Left = cACLBLNAMELEFT
                .TextAlign = ContentAlignment.MiddleLeft
                .Text = student.UserName
                .Show()
            End With
            With chkAutoCapture
                .Width = cACCHKAUTOCAPTUREWIDTH
                .Height = cACCHKAUTOCAPTUREHEIGHT
                .Top = cACCHKAUTOCAPTURETOP
                .Left = cACCHKAUTOCAPTURELEFT
                .Text = ""
                .Tag = student.KeyName
                .Checked = student.AutoCapture
                .Show()
            End With
            AddHandler chkAutoCapture.CheckedChanged, New EventHandler(AddressOf ToggleStudentAC)
            grpMain.Controls.Add(lblName)
            grpMain.Controls.Add(chkAutoCapture)
            Me.tabAutoCapture.Controls.Add(grpMain)
            i += 1
            Select Case i Mod Dimensions.Column
                Case 0
                    intLeft = intStartLeft
                    intTop += cACGRPMAINHEIGHT + cACROWSPACING
                Case Else
                    intLeft += cACGRPMAINWIDTH + cACCOLUMNSPACING
            End Select
        Next
    End Sub

    Private Sub ToggleStudentAC(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CheckBox As CheckBox = sender
        Dim Student As Student = StudentCollection(CheckBox.Tag)
        Student.AutoCapture = CheckBox.Checked
    End Sub

    Private Sub UpdateInterfaceInterval(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkACInterval.ValueChanged
        lblACInterval.Text = IIf(trkACInterval.Value >= 60, "00:" & Int(trkACInterval.Value / 60).ToString.PadLeft(2, "0") & ":" & Int(trkACInterval.Value Mod 60).ToString.PadLeft(2, "0"), "00:00:" & trkACInterval.Value.ToString.PadLeft(2, "0"))
    End Sub

    Private Sub ToggleAutoCapture(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnACToggle.Click
        If AutoCapture.Enabled = True Then
            btnACToggle.Text = "Start"
            AutoCapture.Enabled = False
            trkACInterval.Enabled = True
            tmrAutoCapture.Enabled = False
            Dim DialogResult As DialogResult = MessageBox.Show("Would you like to open the Saved Image Viewer to view the captured thumbnails?", "View Thumbnails", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If DialogResult = Windows.Forms.DialogResult.Yes Then
                ShowImageViewer()
            End If
        Else
            btnACToggle.Text = "Stop"
            AutoCapture.Enabled = True
            trkACInterval.Enabled = False
            AutoCapture.Interval = trkACInterval.Value
            tmrAutoCapture.Enabled = True
        End If
    End Sub

    Private Sub tmrAutoCapture_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrAutoCapture.Tick
        AutoCapture.TickCount += 1
        If AutoCapture.TickCount = AutoCapture.Interval Then
            AutoCapture.TickCount = 0
            ProcessAutoCapture()
        End If
    End Sub

    Private Sub ProcessAutoCapture()
        For Each Student As Student In StudentCollection
            If Student.AutoCapture = True Then
                SaveImage(Student)
            End If
        Next
    End Sub

#End Region     'AutoCapture

#Region "Notepad"

    Private Sub LoadNotepad()
        Try
            Dim StreamReader As New System.IO.StreamReader(Application.StartupPath & cSAVEDNOTESFILEPATH)
            txtNotepad.Text = StreamReader.ReadToEnd
            StreamReader.Close()
        Catch ex As Exception
            Try
                'create blank notes file
                System.IO.File.Create(Application.StartupPath & cSAVEDNOTESFILEPATH)
            Catch subex As Exception
                'do nothing
            End Try
            txtNotepad.Text = ""
            Exit Sub
        End Try
    End Sub

    Private Sub SaveNotepad()
        Try
            Dim StreamWriter As New System.IO.StreamWriter(Application.StartupPath & cSAVEDNOTESFILEPATH)
            StreamWriter.Write(txtNotepad.Text)
            StreamWriter.Close()
        Catch ex As Exception
            MessageBox.Show("There was an error saving the notes.", "Error")
            Exit Sub
        End Try
    End Sub

    'The following subs are best left uncombined since they have varying parameters

    Private Sub tabNotepad_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles tabNotepad.Paint
        LoadNotepad()
    End Sub

    Private Sub btnNotepadReload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNotepadReload.Click
        LoadNotepad()
    End Sub

    Private Sub btnNotepadSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNotepadSave.Click
        SaveNotepad()
    End Sub

    Private Sub tabNotepad_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabNotepad.Leave
        SaveNotepad()
    End Sub

    'Also note SaveNotepad() on SWServer_Closing()
#End Region

End Class