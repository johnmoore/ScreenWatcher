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

Public Class ClassDesigner
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
    Friend WithEvents tclMain As System.Windows.Forms.TabControl
    Friend WithEvents tabOverview As System.Windows.Forms.TabPage
    Friend WithEvents tclBank As System.Windows.Forms.TabControl
    Friend WithEvents tabNodeStudents As System.Windows.Forms.TabPage
    Friend WithEvents tabNodeComputers As System.Windows.Forms.TabPage
    Friend WithEvents tclOptions As System.Windows.Forms.TabControl
    Friend WithEvents tabAddNode As System.Windows.Forms.TabPage
    Friend WithEvents pnlNodeStudent As System.Windows.Forms.Panel
    Friend WithEvents btnNodeAddStudent As System.Windows.Forms.Button
    Friend WithEvents txtNodeIDNumber As System.Windows.Forms.TextBox
    Friend WithEvents lblNodeSIDPrompt As System.Windows.Forms.Label
    Friend WithEvents lblNodeSLNPrompt As System.Windows.Forms.Label
    Friend WithEvents txtNodeLastName As System.Windows.Forms.TextBox
    Friend WithEvents pnlNodeComputer As System.Windows.Forms.Panel
    Friend WithEvents btnNodeAddComputer As System.Windows.Forms.Button
    Friend WithEvents txtNodeComputerName As System.Windows.Forms.TextBox
    Friend WithEvents lblNodeCompNamePrompt As System.Windows.Forms.Label
    Friend WithEvents cboNodeType As System.Windows.Forms.ComboBox
    Friend WithEvents tabActions As System.Windows.Forms.TabPage
    Friend WithEvents btnAddSeat As System.Windows.Forms.Button
    Friend WithEvents btnRemoveSeat As System.Windows.Forms.Button
    Friend WithEvents btnAddRow As System.Windows.Forms.Button
    Friend WithEvents btnRemoveRow As System.Windows.Forms.Button
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuClose As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNew As System.Windows.Forms.MenuItem
    Friend WithEvents dlgSaveFile As System.Windows.Forms.SaveFileDialog
    Friend WithEvents dlgOpenFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnEditPeriodInfo As System.Windows.Forms.Button
    Friend WithEvents tabClearBank As System.Windows.Forms.TabPage
    Friend WithEvents mnuSeperator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSeperator2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tclMain = New System.Windows.Forms.TabControl
        Me.tabOverview = New System.Windows.Forms.TabPage
        Me.tclBank = New System.Windows.Forms.TabControl
        Me.tabNodeStudents = New System.Windows.Forms.TabPage
        Me.tabNodeComputers = New System.Windows.Forms.TabPage
        Me.tabClearBank = New System.Windows.Forms.TabPage
        Me.tclOptions = New System.Windows.Forms.TabControl
        Me.tabAddNode = New System.Windows.Forms.TabPage
        Me.pnlNodeStudent = New System.Windows.Forms.Panel
        Me.btnNodeAddStudent = New System.Windows.Forms.Button
        Me.txtNodeIDNumber = New System.Windows.Forms.TextBox
        Me.lblNodeSIDPrompt = New System.Windows.Forms.Label
        Me.lblNodeSLNPrompt = New System.Windows.Forms.Label
        Me.txtNodeLastName = New System.Windows.Forms.TextBox
        Me.pnlNodeComputer = New System.Windows.Forms.Panel
        Me.btnNodeAddComputer = New System.Windows.Forms.Button
        Me.txtNodeComputerName = New System.Windows.Forms.TextBox
        Me.lblNodeCompNamePrompt = New System.Windows.Forms.Label
        Me.cboNodeType = New System.Windows.Forms.ComboBox
        Me.tabActions = New System.Windows.Forms.TabPage
        Me.btnEditPeriodInfo = New System.Windows.Forms.Button
        Me.btnRemoveRow = New System.Windows.Forms.Button
        Me.btnAddRow = New System.Windows.Forms.Button
        Me.btnRemoveSeat = New System.Windows.Forms.Button
        Me.btnAddSeat = New System.Windows.Forms.Button
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuNew = New System.Windows.Forms.MenuItem
        Me.mnuOpen = New System.Windows.Forms.MenuItem
        Me.mnuSeperator1 = New System.Windows.Forms.MenuItem
        Me.mnuSave = New System.Windows.Forms.MenuItem
        Me.mnuSaveAs = New System.Windows.Forms.MenuItem
        Me.mnuSeperator2 = New System.Windows.Forms.MenuItem
        Me.mnuClose = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.dlgSaveFile = New System.Windows.Forms.SaveFileDialog
        Me.dlgOpenFile = New System.Windows.Forms.OpenFileDialog
        Me.tclMain.SuspendLayout()
        Me.tclBank.SuspendLayout()
        Me.tclOptions.SuspendLayout()
        Me.tabAddNode.SuspendLayout()
        Me.pnlNodeStudent.SuspendLayout()
        Me.pnlNodeComputer.SuspendLayout()
        Me.tabActions.SuspendLayout()
        Me.SuspendLayout()
        '
        'tclMain
        '
        Me.tclMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tclMain.Controls.Add(Me.tabOverview)
        Me.tclMain.Location = New System.Drawing.Point(205, 12)
        Me.tclMain.Name = "tclMain"
        Me.tclMain.SelectedIndex = 0
        Me.tclMain.Size = New System.Drawing.Size(484, 469)
        Me.tclMain.TabIndex = 0
        Me.tclMain.Tag = "1"
        '
        'tabOverview
        '
        Me.tabOverview.AutoScroll = True
        Me.tabOverview.Location = New System.Drawing.Point(4, 22)
        Me.tabOverview.Name = "tabOverview"
        Me.tabOverview.Size = New System.Drawing.Size(476, 443)
        Me.tabOverview.TabIndex = 0
        Me.tabOverview.Tag = "0"
        Me.tabOverview.Text = "Overview"
        '
        'tclBank
        '
        Me.tclBank.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tclBank.Controls.Add(Me.tabNodeStudents)
        Me.tclBank.Controls.Add(Me.tabNodeComputers)
        Me.tclBank.Controls.Add(Me.tabClearBank)
        Me.tclBank.Location = New System.Drawing.Point(12, 142)
        Me.tclBank.Name = "tclBank"
        Me.tclBank.SelectedIndex = 0
        Me.tclBank.Size = New System.Drawing.Size(187, 339)
        Me.tclBank.TabIndex = 1
        Me.tclBank.Tag = "0"
        '
        'tabNodeStudents
        '
        Me.tabNodeStudents.AutoScroll = True
        Me.tabNodeStudents.Location = New System.Drawing.Point(4, 22)
        Me.tabNodeStudents.Name = "tabNodeStudents"
        Me.tabNodeStudents.Size = New System.Drawing.Size(179, 313)
        Me.tabNodeStudents.TabIndex = 0
        Me.tabNodeStudents.Tag = "0"
        Me.tabNodeStudents.Text = "Students"
        '
        'tabNodeComputers
        '
        Me.tabNodeComputers.AutoScroll = True
        Me.tabNodeComputers.Location = New System.Drawing.Point(4, 22)
        Me.tabNodeComputers.Name = "tabNodeComputers"
        Me.tabNodeComputers.Size = New System.Drawing.Size(179, 313)
        Me.tabNodeComputers.TabIndex = 1
        Me.tabNodeComputers.Tag = "1"
        Me.tabNodeComputers.Text = "Computers"
        '
        'tabClearBank
        '
        Me.tabClearBank.Location = New System.Drawing.Point(4, 22)
        Me.tabClearBank.Name = "tabClearBank"
        Me.tabClearBank.Size = New System.Drawing.Size(179, 313)
        Me.tabClearBank.TabIndex = 2
        Me.tabClearBank.Tag = "2"
        Me.tabClearBank.Text = "Clear..."
        '
        'tclOptions
        '
        Me.tclOptions.Controls.Add(Me.tabAddNode)
        Me.tclOptions.Controls.Add(Me.tabActions)
        Me.tclOptions.Location = New System.Drawing.Point(8, 9)
        Me.tclOptions.Name = "tclOptions"
        Me.tclOptions.SelectedIndex = 0
        Me.tclOptions.Size = New System.Drawing.Size(191, 127)
        Me.tclOptions.TabIndex = 2
        '
        'tabAddNode
        '
        Me.tabAddNode.Controls.Add(Me.pnlNodeStudent)
        Me.tabAddNode.Controls.Add(Me.pnlNodeComputer)
        Me.tabAddNode.Controls.Add(Me.cboNodeType)
        Me.tabAddNode.Location = New System.Drawing.Point(4, 22)
        Me.tabAddNode.Name = "tabAddNode"
        Me.tabAddNode.Size = New System.Drawing.Size(183, 101)
        Me.tabAddNode.TabIndex = 0
        Me.tabAddNode.Text = "Add Node"
        '
        'pnlNodeStudent
        '
        Me.pnlNodeStudent.BackColor = System.Drawing.Color.Transparent
        Me.pnlNodeStudent.Controls.Add(Me.btnNodeAddStudent)
        Me.pnlNodeStudent.Controls.Add(Me.txtNodeIDNumber)
        Me.pnlNodeStudent.Controls.Add(Me.lblNodeSIDPrompt)
        Me.pnlNodeStudent.Controls.Add(Me.lblNodeSLNPrompt)
        Me.pnlNodeStudent.Controls.Add(Me.txtNodeLastName)
        Me.pnlNodeStudent.Location = New System.Drawing.Point(4, 35)
        Me.pnlNodeStudent.Name = "pnlNodeStudent"
        Me.pnlNodeStudent.Size = New System.Drawing.Size(175, 54)
        Me.pnlNodeStudent.TabIndex = 14
        '
        'btnNodeAddStudent
        '
        Me.btnNodeAddStudent.Location = New System.Drawing.Point(133, 31)
        Me.btnNodeAddStudent.Name = "btnNodeAddStudent"
        Me.btnNodeAddStudent.Size = New System.Drawing.Size(36, 20)
        Me.btnNodeAddStudent.TabIndex = 9
        Me.btnNodeAddStudent.Tag = "0"
        Me.btnNodeAddStudent.Text = "Add"
        '
        'txtNodeIDNumber
        '
        Me.txtNodeIDNumber.Location = New System.Drawing.Point(72, 31)
        Me.txtNodeIDNumber.Name = "txtNodeIDNumber"
        Me.txtNodeIDNumber.Size = New System.Drawing.Size(56, 20)
        Me.txtNodeIDNumber.TabIndex = 8
        Me.txtNodeIDNumber.Text = ""
        '
        'lblNodeSIDPrompt
        '
        Me.lblNodeSIDPrompt.AutoSize = True
        Me.lblNodeSIDPrompt.Location = New System.Drawing.Point(7, 31)
        Me.lblNodeSIDPrompt.Name = "lblNodeSIDPrompt"
        Me.lblNodeSIDPrompt.Size = New System.Drawing.Size(62, 16)
        Me.lblNodeSIDPrompt.TabIndex = 7
        Me.lblNodeSIDPrompt.Text = "ID Number:"
        '
        'lblNodeSLNPrompt
        '
        Me.lblNodeSLNPrompt.AutoSize = True
        Me.lblNodeSLNPrompt.Location = New System.Drawing.Point(7, 8)
        Me.lblNodeSLNPrompt.Name = "lblNodeSLNPrompt"
        Me.lblNodeSLNPrompt.Size = New System.Drawing.Size(62, 16)
        Me.lblNodeSLNPrompt.TabIndex = 6
        Me.lblNodeSLNPrompt.Text = "Last Name:"
        '
        'txtNodeLastName
        '
        Me.txtNodeLastName.Location = New System.Drawing.Point(72, 5)
        Me.txtNodeLastName.Name = "txtNodeLastName"
        Me.txtNodeLastName.Size = New System.Drawing.Size(97, 20)
        Me.txtNodeLastName.TabIndex = 5
        Me.txtNodeLastName.Text = ""
        '
        'pnlNodeComputer
        '
        Me.pnlNodeComputer.BackColor = System.Drawing.Color.Transparent
        Me.pnlNodeComputer.Controls.Add(Me.btnNodeAddComputer)
        Me.pnlNodeComputer.Controls.Add(Me.txtNodeComputerName)
        Me.pnlNodeComputer.Controls.Add(Me.lblNodeCompNamePrompt)
        Me.pnlNodeComputer.Location = New System.Drawing.Point(4, 35)
        Me.pnlNodeComputer.Name = "pnlNodeComputer"
        Me.pnlNodeComputer.Size = New System.Drawing.Size(175, 54)
        Me.pnlNodeComputer.TabIndex = 13
        Me.pnlNodeComputer.Visible = False
        '
        'btnNodeAddComputer
        '
        Me.btnNodeAddComputer.Location = New System.Drawing.Point(133, 17)
        Me.btnNodeAddComputer.Name = "btnNodeAddComputer"
        Me.btnNodeAddComputer.Size = New System.Drawing.Size(36, 21)
        Me.btnNodeAddComputer.TabIndex = 10
        Me.btnNodeAddComputer.Tag = "1"
        Me.btnNodeAddComputer.Text = "Add"
        '
        'txtNodeComputerName
        '
        Me.txtNodeComputerName.Location = New System.Drawing.Point(10, 18)
        Me.txtNodeComputerName.Name = "txtNodeComputerName"
        Me.txtNodeComputerName.Size = New System.Drawing.Size(118, 20)
        Me.txtNodeComputerName.TabIndex = 8
        Me.txtNodeComputerName.Text = ""
        '
        'lblNodeCompNamePrompt
        '
        Me.lblNodeCompNamePrompt.AutoSize = True
        Me.lblNodeCompNamePrompt.Location = New System.Drawing.Point(7, 2)
        Me.lblNodeCompNamePrompt.Name = "lblNodeCompNamePrompt"
        Me.lblNodeCompNamePrompt.Size = New System.Drawing.Size(90, 16)
        Me.lblNodeCompNamePrompt.TabIndex = 7
        Me.lblNodeCompNamePrompt.Text = "Computer Name:"
        '
        'cboNodeType
        '
        Me.cboNodeType.DisplayMember = "0, 1"
        Me.cboNodeType.Items.AddRange(New Object() {"Student", "Computer"})
        Me.cboNodeType.Location = New System.Drawing.Point(14, 11)
        Me.cboNodeType.Name = "cboNodeType"
        Me.cboNodeType.Size = New System.Drawing.Size(159, 21)
        Me.cboNodeType.TabIndex = 12
        Me.cboNodeType.Text = "Select Type"
        '
        'tabActions
        '
        Me.tabActions.Controls.Add(Me.btnEditPeriodInfo)
        Me.tabActions.Controls.Add(Me.btnRemoveRow)
        Me.tabActions.Controls.Add(Me.btnAddRow)
        Me.tabActions.Controls.Add(Me.btnRemoveSeat)
        Me.tabActions.Controls.Add(Me.btnAddSeat)
        Me.tabActions.Location = New System.Drawing.Point(4, 22)
        Me.tabActions.Name = "tabActions"
        Me.tabActions.Size = New System.Drawing.Size(183, 101)
        Me.tabActions.TabIndex = 1
        Me.tabActions.Text = "Actions"
        '
        'btnEditPeriodInfo
        '
        Me.btnEditPeriodInfo.Location = New System.Drawing.Point(6, 64)
        Me.btnEditPeriodInfo.Name = "btnEditPeriodInfo"
        Me.btnEditPeriodInfo.Size = New System.Drawing.Size(171, 23)
        Me.btnEditPeriodInfo.TabIndex = 6
        Me.btnEditPeriodInfo.Text = "Edit Period Info..."
        '
        'btnRemoveRow
        '
        Me.btnRemoveRow.Location = New System.Drawing.Point(92, 6)
        Me.btnRemoveRow.Name = "btnRemoveRow"
        Me.btnRemoveRow.Size = New System.Drawing.Size(85, 23)
        Me.btnRemoveRow.TabIndex = 5
        Me.btnRemoveRow.Tag = "X"
        Me.btnRemoveRow.Text = "Remove Row"
        '
        'btnAddRow
        '
        Me.btnAddRow.Location = New System.Drawing.Point(6, 6)
        Me.btnAddRow.Name = "btnAddRow"
        Me.btnAddRow.Size = New System.Drawing.Size(86, 23)
        Me.btnAddRow.TabIndex = 4
        Me.btnAddRow.Tag = "X"
        Me.btnAddRow.Text = "Add Row"
        '
        'btnRemoveSeat
        '
        Me.btnRemoveSeat.Location = New System.Drawing.Point(92, 35)
        Me.btnRemoveSeat.Name = "btnRemoveSeat"
        Me.btnRemoveSeat.Size = New System.Drawing.Size(85, 23)
        Me.btnRemoveSeat.TabIndex = 3
        Me.btnRemoveSeat.Tag = "Y"
        Me.btnRemoveSeat.Text = "Remove Seat"
        '
        'btnAddSeat
        '
        Me.btnAddSeat.Location = New System.Drawing.Point(6, 35)
        Me.btnAddSeat.Name = "btnAddSeat"
        Me.btnAddSeat.Size = New System.Drawing.Size(86, 23)
        Me.btnAddSeat.TabIndex = 2
        Me.btnAddSeat.Tag = "Y"
        Me.btnAddSeat.Text = "Add Seat"
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNew, Me.mnuOpen, Me.mnuSeperator1, Me.mnuSave, Me.mnuSaveAs, Me.mnuSeperator2, Me.mnuClose, Me.mnuExit})
        Me.mnuFile.Text = "File"
        '
        'mnuNew
        '
        Me.mnuNew.Index = 0
        Me.mnuNew.Text = "New..."
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 1
        Me.mnuOpen.Text = "Open..."
        '
        'mnuSeperator1
        '
        Me.mnuSeperator1.Index = 2
        Me.mnuSeperator1.Text = "-"
        '
        'mnuSave
        '
        Me.mnuSave.Index = 3
        Me.mnuSave.Text = "Save"
        '
        'mnuSaveAs
        '
        Me.mnuSaveAs.Index = 4
        Me.mnuSaveAs.Text = "Save As..."
        '
        'mnuSeperator2
        '
        Me.mnuSeperator2.Index = 5
        Me.mnuSeperator2.Text = "-"
        '
        'mnuClose
        '
        Me.mnuClose.Index = 6
        Me.mnuClose.Text = "Close"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 7
        Me.mnuExit.Text = "Exit"
        '
        'dlgOpenFile
        '
        Me.dlgOpenFile.FileName = "OpenFileDialog1"
        '
        'ClassDesigner
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(701, 488)
        Me.Controls.Add(Me.tclOptions)
        Me.Controls.Add(Me.tclBank)
        Me.Controls.Add(Me.tclMain)
        Me.Menu = Me.mnuMain
        Me.Name = "ClassDesigner"
        Me.Text = "Class Designer"
        Me.tclMain.ResumeLayout(False)
        Me.tclBank.ResumeLayout(False)
        Me.tclOptions.ResumeLayout(False)
        Me.tabAddNode.ResumeLayout(False)
        Me.pnlNodeStudent.ResumeLayout(False)
        Me.pnlNodeComputer.ResumeLayout(False)
        Me.tabActions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constants"
    Private Const cPSGRPMAINWIDTH As Integer = 150
    Private Const cPSGRPMAINHEIGHT As Integer = 30

    Private Const cPSLBLNAMEWIDTH As Integer = 120
    Private Const cPSLBLNAMEHEIGHT As Integer = 13
    Private Const cPSLBLNAMELEFT As Integer = 5
    Private Const cPSLBLNAMETOP As Integer = 15

    Private Const cPSBTNREMOVEWIDTH As Integer = 10
    Private Const cPSBTNREMOVEHEIGHT As Integer = 10
    Private Const cPSBTNREMOVELEFT As Integer = cPSLBLNAMELEFT + cPSLBLNAMEWIDTH + 10
    Private Const cPSBTNREMOVETOP As Integer = 12


    Private Const cBANKPSLEFTSPACING As Integer = 10
    Private Const cBANKPSTOPSPACING As Integer = 10


    Private Const cMAINPSTOPSPACING As Integer = 10


    Private Const cPSCAPTIONSWAP As String = "[swap]"
    Private Const cPSCAPTIONNONE As String = ""
    Private Const cPSCAPTIONTOBANK As String = "[back to bank]"
    Private Const cPSCAPTIONSELECT As String = "[select]"
    Private Const cPSCAPTIONDESELECT As String = "[deselect]"
    Private Const cPSCAPTIONDEFAULT As String = cPSCAPTIONSELECT


    Private Const cOVERVIEWGRPMAINWIDTH As Integer = 100
    Private Const cOVERVIEWGRPMAINHEIGHT As Integer = 100

    Private Const cOVERVIEWLBLNAMEWIDTH As Integer = 85
    Private Const cOVERVIEWLBLNAMEHEIGHT As Integer = 85
    Private Const cOVERVIEWLBLMAINLEFT As Integer = 7
    Private Const cOVERVIEWLBLMAINTOP As Integer = 7

    Private Const cOVERVIEWGRPMAINROWSPACING As Integer = 10
    Private Const cOVERVIEWGRPMAINCOLUMNSPACING As Integer = 10

#End Region

#Region "Enums"
    Enum ErrorType
        InvalidNodeType = 0
        InvalidNodeInfo = 1
        DuplicateNodeName = 2
        NodeRemoveFailed = 3
        NodeSelectFailed = 4
        NodeSwapFailed = 5
        AddDimensionFailed = 6
        RemoveDimensionFailed = 7
        SavePeriodFailed = 8
        FileOpenFailed = 9
        LoadPeriodFailed = 10   'invalid period file, etc.
        Unknown = 11
    End Enum
    Enum StudentNodeType
        Student = 0
        Computer = 1
        Placeholder = 2
    End Enum
    Enum StudentNodeLocation
        Bank = 0
        Main = 1
    End Enum
#End Region

#Region "Structures"
    Structure PeriodStudentData
        Dim NodeName As String
        Dim NodeType As StudentNodeType
        Dim Position As Point
    End Structure
    Structure PeriodStudentControls
        Dim grpMain As GroupBox
        Dim lblName As Label
        Dim btnRemove As Button
    End Structure
    Structure PeriodStudent
        Dim Data As PeriodStudentData
        Dim Controls As PeriodStudentControls

        Public Sub New(ByVal Type As StudentNodeType, ByVal Name As String, ByVal Position As Point, ByVal RemoveNode As EventHandler, ByVal NodeAction As EventHandler)
            Data.NodeType = Type
            Data.NodeName = Name
            Data.Position = Position
            'create controls
            Controls.grpMain = New GroupBox
            Controls.lblName = New Label
            Controls.btnRemove = New Button
            With Controls.grpMain
                .Width = cPSGRPMAINWIDTH
                .Height = cPSGRPMAINHEIGHT
                '.left set later
                '.top set later
                .Tag = Data.NodeName
                .Text = cPSCAPTIONDEFAULT
                .Show()
            End With
            With Controls.lblName
                .Width = cPSLBLNAMEWIDTH
                .Height = cPSLBLNAMEHEIGHT
                .Left = cPSLBLNAMELEFT
                .Top = cPSLBLNAMETOP
                .TextAlign = ContentAlignment.MiddleLeft
                .Text = IIf(Data.NodeType = StudentNodeType.Placeholder, "", Data.NodeName)
                .Show()
            End With
            With Controls.btnRemove
                .Width = cPSBTNREMOVEWIDTH
                .Height = cPSBTNREMOVEHEIGHT
                .Left = cPSBTNREMOVELEFT
                .Top = cPSBTNREMOVETOP
                .Text = "X"
                .Tag = Data.NodeName
                .FlatStyle = FlatStyle.Flat
                .Show()
            End With
            'handlers
            AddHandler Controls.btnRemove.Click, RemoveNode
            AddHandler Controls.grpMain.Click, NodeAction

            Controls.grpMain.Controls.Add(Controls.lblName)
            If Data.NodeType <> StudentNodeType.Placeholder Then
                Controls.grpMain.Controls.Add(Controls.btnRemove)
            End If
        End Sub
    End Structure

    Structure SelectedStudentData
        Dim bolSelected As Boolean
        Dim PeriodStudent As PeriodStudent
    End Structure

    Structure ClassPeriodInfo
        Dim PeriodNumber As Integer
        Dim Dimensions As Point
        Dim Title As String
    End Structure

    Structure PeriodDialogData
        Dim Form As PeriodDialog
        Dim bolOpen As Boolean
    End Structure

#End Region

#Region "New Stuff"
    Public ReadOnly Property pPeriodInfo() As ClassPeriodInfo
        Get
            Return PeriodInfo
        End Get
    End Property

    Public ReadOnly Property pPeriodDialog() As PeriodDialogData
        Get
            Return stcPeriodDialog
        End Get
    End Property

    Public Sub EditPeriodInfo(ByVal PeriodNumber As Integer, ByVal Title As String)
        Me.PeriodInfo.PeriodNumber = PeriodNumber
        Me.PeriodInfo.Title = Title
    End Sub

    Public Sub NewPeriodInfo(ByVal Dimensions As Point, ByVal PeriodNumber As Integer, ByVal Title As String)
        PeriodInfo.Dimensions = Dimensions
        PeriodInfo.PeriodNumber = PeriodNumber
        PeriodInfo.Title = Title
    End Sub

    Public Sub PeriodDialogBolOpen(ByVal TorF As Boolean)
        stcPeriodDialog.bolOpen = TorF
    End Sub

    Private Function CollectionContains(ByVal Collection As Collection, ByVal Key As String) As Boolean
        Try
            Dim test As Object = Collection(Key)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub RemoveTabPage(ByRef Collection As TabControl.TabPageCollection, ByVal Key As String)
        For Each TabPage As TabPage In Collection
            If TabPage.Name = Key Then
                Collection.Remove(TabPage)
                Exit Sub
            End If
        Next
    End Sub

    Private evtNodeAction As New EventHandler(AddressOf NodeAction)

    Private evtRemoveNode As New EventHandler(AddressOf RemoveNode)

#End Region   'Must be before globals

#Region "Globals"
    Dim BankPeriodStudents As New Collection
    Dim MainPeriodStudents As New Collection

    Dim StaticBankNode As PeriodStudent = New PeriodStudent(StudentNodeType.Placeholder, "", New Point, evtRemoveNode, evtNodeAction)

    'Dim PeriodStudentCollections() As Collection = {BankPeriodStudents, MainPeriodStudents}

    Dim Dimensions As Point
    Dim SelectedStudent As SelectedStudentData

    Public FileOpen As Boolean = False
    Dim LastSavePath As String = ""

    Public PeriodInfo As ClassPeriodInfo

    Public stcPeriodDialog As PeriodDialogData

#End Region

#Region "Methods"

#Region "Event Handlers"

    Private Sub NodeAction(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim grpMain As GroupBox = sender
        Dim StudentType As StudentNodeLocation = Me.GetStudentType(grpMain.Tag)
        Dim PeriodStudentCollection As Collection
        Dim PeriodStudent As PeriodStudent
        Try
            Select Case StudentType
                Case -1
                    If Me.SelectedStudent.bolSelected = False Then
                        grpMain.BackColor = Color.LightBlue
                        Me.SelectedStudent.bolSelected = True
                        Me.SelectedStudent.PeriodStudent = Me.StaticBankNode 'New PeriodStudent(StudentNodeType.Placeholder, "", New Point())
                    Else
                        If Me.SelectedStudent.PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                            Me.SwapNodes(New PeriodStudent(StudentNodeType.Placeholder, "", New Point, evtRemoveNode, evtNodeAction), Me.SelectedStudent.PeriodStudent)
                        Else
                            Me.ClearSelections()
                        End If
                    End If
                Case StudentNodeLocation.Bank
                    PeriodStudentCollection = GetStudentCollection(StudentType)
                    PeriodStudent = PeriodStudentCollection.Item(grpMain.Tag)
                    If Me.SelectedStudent.bolSelected = False Then
                        PeriodStudent.Controls.grpMain.BackColor = Color.LightBlue
                        Me.SelectedStudent.bolSelected = True
                        Me.SelectedStudent.PeriodStudent = PeriodStudent
                    Else
                        If Me.SelectedStudent.PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Or PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                            Me.SwapNodes(Me.SelectedStudent.PeriodStudent, PeriodStudent)
                        Else
                            Me.ClearSelections()
                        End If
                    End If
                Case StudentNodeLocation.Main
                    PeriodStudentCollection = GetStudentCollection(StudentType)
                    PeriodStudent = PeriodStudentCollection.Item(grpMain.Tag)
                    If Me.SelectedStudent.bolSelected = False Then
                        PeriodStudent.Controls.grpMain.BackColor = Color.LightBlue
                        Me.SelectedStudent.bolSelected = True
                        Me.SelectedStudent.PeriodStudent = PeriodStudent
                    Else
                        Me.SwapNodes(Me.SelectedStudent.PeriodStudent, PeriodStudent)
                    End If
            End Select
            Me.UpdateNodeCaptions()
        Catch ex As Exception
            Me.RaiseError(ErrorType.NodeSelectFailed)
            Application.DoEvents()
            Exit Sub
        End Try
        Application.DoEvents()
    End Sub

    Private Sub SwapNodes(ByVal psOriginal As PeriodStudent, ByVal psNew As PeriodStudent)
        ClearSelections()
        Try
            If psOriginal.Data.Position.Equals(psNew.Data.Position) Then
                Exit Sub
            End If
            Dim OriginalTag As String = psOriginal.Controls.grpMain.Tag
            Dim OriginalStudentType As StudentNodeLocation = GetStudentType(OriginalTag)
            Dim OriginalPeriodStudent As PeriodStudent
            Dim OriginalPeriodStudentCollection As Collection
            Dim NewTag As String = psNew.Controls.grpMain.Tag
            Dim NewStudentType As StudentNodeLocation = GetStudentType(NewTag)
            Dim NewPeriodStudent As PeriodStudent
            Dim NewPeriodStudentCollection As Collection
            NewPeriodStudentCollection = GetStudentCollection(NewStudentType)
            NewPeriodStudent = NewPeriodStudentCollection.Item(NewTag)
            If OriginalStudentType <> -1 Then
                OriginalPeriodStudentCollection = GetStudentCollection(OriginalStudentType)
                OriginalPeriodStudent = OriginalPeriodStudentCollection.Item(OriginalTag)
            Else
                OriginalPeriodStudentCollection = New Collection
                OriginalPeriodStudent = New PeriodStudent(StudentNodeType.Placeholder, "", New Point, evtRemoveNode, evtNodeAction)
            End If

            If IsStaticBankNode(OriginalPeriodStudent) Or IsStaticBankNode(NewPeriodStudent) Then
                If OriginalPeriodStudent.Data.NodeType = NewPeriodStudent.Data.NodeType Then
                    Exit Sub
                ElseIf OriginalPeriodStudent.Data.Position.Equals(NewPeriodStudent.Data.Position) Then
                    Exit Sub
                Else
                    'static with node
                    If OriginalPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                        'original is static
                        Dim x As Integer = NewPeriodStudent.Data.Position.X
                        Dim y As Integer = NewPeriodStudent.Data.Position.Y
                        NewPeriodStudentCollection.Remove(NewTag)
                        NewPeriodStudentCollection.Add(New PeriodStudent(StudentNodeType.Placeholder, x & "," & y, NewPeriodStudent.Data.Position, evtRemoveNode, evtNodeAction), x & "," & y)
                        MoveNodeToBank(NewPeriodStudent)
                    Else
                        'new is static
                        OriginalPeriodStudentCollection.Remove(OriginalTag)
                        OriginalPeriodStudentCollection.Add(New PeriodStudent(StudentNodeType.Placeholder, "", OriginalPeriodStudent.Data.Position, evtRemoveNode, evtNodeAction), OriginalPeriodStudent.Data.Position.X & "," & OriginalPeriodStudent.Data.Position.Y)
                        MoveNodeToBank(OriginalPeriodStudent)
                    End If
                End If
            ElseIf OriginalPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Or NewPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                If OriginalPeriodStudent.Data.NodeType = NewPeriodStudent.Data.NodeType Then
                    Exit Sub
                Else
                    'blank with node
                    NewPeriodStudentCollection.Remove(NewTag)
                    OriginalPeriodStudentCollection.Remove(OriginalTag)
                    Dim OriginalPoint As Point = OriginalPeriodStudent.Data.Position
                    Dim NewPoint As Point = NewPeriodStudent.Data.Position
                    OriginalPeriodStudent.Data.Position = NewPoint
                    NewPeriodStudent.Data.Position = OriginalPoint
                    If OriginalPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                        OriginalPeriodStudent.Data.NodeName = OriginalPeriodStudent.Data.Position.X & "," & OriginalPeriodStudent.Data.Position.Y
                    ElseIf NewPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                        NewPeriodStudent.Data.NodeName = NewPeriodStudent.Data.Position.X & "," & NewPeriodStudent.Data.Position.Y
                    End If
                    NewPeriodStudent.Controls.grpMain.Tag = NewPeriodStudent.Data.NodeName
                    OriginalPeriodStudent.Controls.grpMain.Tag = OriginalPeriodStudent.Data.NodeName

                    NewPeriodStudentCollection.Add(OriginalPeriodStudent, OriginalPeriodStudent.Data.NodeName)
                    OriginalPeriodStudentCollection.Add(NewPeriodStudent, NewPeriodStudent.Data.NodeName)
                    PurgeBank()
                End If
            Else
                'node with node
                NewPeriodStudentCollection.Remove(NewTag)
                OriginalPeriodStudentCollection.Remove(OriginalTag)
                Dim OriginalPoint As Point = OriginalPeriodStudent.Data.Position
                Dim NewPoint As Point = NewPeriodStudent.Data.Position
                OriginalPeriodStudent.Data.Position = NewPoint
                NewPeriodStudent.Data.Position = OriginalPoint
                If OriginalPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                    OriginalPeriodStudent.Data.NodeName = OriginalPeriodStudent.Data.Position.X & "," & OriginalPeriodStudent.Data.Position.Y
                ElseIf NewPeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                    NewPeriodStudent.Data.NodeName = NewPeriodStudent.Data.Position.X & "," & NewPeriodStudent.Data.Position.Y
                End If
                NewPeriodStudent.Controls.grpMain.Tag = NewPeriodStudent.Data.NodeName
                OriginalPeriodStudent.Controls.grpMain.Tag = OriginalPeriodStudent.Data.NodeName

                NewPeriodStudentCollection.Add(OriginalPeriodStudent, OriginalPeriodStudent.Data.NodeName)
                OriginalPeriodStudentCollection.Add(NewPeriodStudent, NewPeriodStudent.Data.NodeName)
            End If
        Catch ex As Exception
            RaiseError(ErrorType.NodeSwapFailed)
        End Try
        ReloadVisibleInterface()
    End Sub

    Private Sub AddNode(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNodeAddStudent.Click, btnNodeAddComputer.Click
        Dim btnAddNode As Button = sender
        Select Case Val(btnAddNode.Tag)
            Case StudentNodeType.Student
                If ValidateNodeInput(StudentNodeType.Student) = True Then
                    Dim PeriodStudent As New PeriodStudent(StudentNodeType.Student, txtNodeIDNumber.Text & txtNodeLastName.Text, New Point, evtRemoveNode, evtNodeAction)
                    Try
                        If CollectionPeriodStudentExists(MainPeriodStudents, PeriodStudent.Data.NodeName) = False And CollectionPeriodStudentExists(BankPeriodStudents, PeriodStudent.Data.NodeName) = False Then
                            BankPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.NodeName)
                        Else
                            Throw New Exception 'Duplicate
                        End If
                    Catch ex As Exception
                        RaiseError(ErrorType.DuplicateNodeName)
                        Exit Sub
                    End Try
                    ReloadBank(PeriodStudent.Data.NodeType)
                Else
                    RaiseError(ErrorType.InvalidNodeInfo)
                End If
            Case StudentNodeType.Computer
                If ValidateNodeInput(StudentNodeType.Computer) = True Then
                    Dim PeriodStudent As New PeriodStudent(StudentNodeType.Computer, txtNodeComputerName.Text, New Point, evtRemoveNode, evtNodeAction)
                    Try
                        If CollectionPeriodStudentExists(MainPeriodStudents, PeriodStudent.Data.NodeName) = False And CollectionPeriodStudentExists(BankPeriodStudents, PeriodStudent.Data.NodeName) = False Then
                            BankPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.NodeName)
                        Else
                            Throw New Exception
                        End If
                    Catch ex As Exception
                        RaiseError(ErrorType.DuplicateNodeName)
                        Exit Sub
                    End Try
                    ReloadBank(tclBank.SelectedIndex)
                Else
                    RaiseError(ErrorType.InvalidNodeInfo)
                End If
        End Select
        ClearNodeInput()
    End Sub

    Private Sub RemoveNode(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim btnRemove As Button = sender
        Dim StudentType As StudentNodeLocation = Me.GetStudentType(btnRemove.Tag)
        Dim PeriodStudentCollection As Collection
        Dim PeriodStudent As PeriodStudent
        Try
            Select Case StudentType
                Case StudentNodeLocation.Bank
                    PeriodStudentCollection = GetStudentCollection(StudentType)
                    PeriodStudent = PeriodStudentCollection.Item(btnRemove.Tag)
                    PeriodStudentCollection.Remove(btnRemove.Tag)
                    Me.ReloadBank(PeriodStudent.Data.NodeType)
                Case StudentNodeLocation.Main
                    PeriodStudentCollection = GetStudentCollection(StudentType)
                    PeriodStudent = PeriodStudentCollection.Item(btnRemove.Tag)
                    PeriodStudentCollection.Remove(btnRemove.Tag)
                    Dim Placeholder As New PeriodStudent(StudentNodeType.Placeholder, PeriodStudent.Data.Position.X & "," & PeriodStudent.Data.Position.Y, New Point(PeriodStudent.Data.Position.X, PeriodStudent.Data.Position.Y), evtRemoveNode, evtNodeAction)
                    Me.MainPeriodStudents.Add(Placeholder, Placeholder.Data.Position.X & "," & Placeholder.Data.Position.Y)
                    Me.ReloadRow(PeriodStudent.Data.Position.X)
            End Select
        Catch ex As Exception
            Me.RaiseError(ErrorType.NodeRemoveFailed)
            Exit Sub
        End Try
    End Sub

#Region "Static Form Events"

    Private Sub AddStaticHandlers()
        AddHandler tclBank.SelectedIndexChanged, New EventHandler(AddressOf TabChange)
        AddHandler tclMain.SelectedIndexChanged, New EventHandler(AddressOf TabChange)
        AddHandler tclMain.SizeChanged, New EventHandler(AddressOf tclMain_SizeChanged)
        AddHandler cboNodeType.SelectedIndexChanged, New EventHandler(AddressOf cboNodeType_SelectedIndexChanged)
    End Sub

    Private Sub mnuFile_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFile.Popup
        UpdateMenuControls()
    End Sub

    Private Sub tclMain_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles tclMain.SizeChanged
        If tclMain.SelectedIndex = 0 Then
            ReloadOverview()
        Else
            ReloadRow(Val(tclMain.SelectedTab.Tag))
        End If
    End Sub

    Private Sub cboNodeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles cboNodeType.SelectedIndexChanged
        Select Case cboNodeType.SelectedIndex
            Case -1
                RaiseError(ErrorType.InvalidNodeType)
                Exit Sub
            Case StudentNodeType.Student
                pnlNodeComputer.Visible = False
                pnlNodeStudent.Visible = True
            Case StudentNodeType.Computer
                pnlNodeStudent.Visible = False
                pnlNodeComputer.Visible = True
        End Select
    End Sub

    Private Sub TabChange(ByVal sender As Object, ByVal e As System.EventArgs) 'handles tclBank.SelectedIndexChanged, tclMain.SelectedIndexChanged 
        Dim TabControl As TabControl = sender
        If Val(TabControl.Tag) = 0 Then 'bank
            If Val(TabControl.SelectedTab.Tag) = 2 Then
                tclBank.SelectedIndex = 0
                Dim DialogResult As DialogResult = MessageBox.Show("Are you sure you want to clear all students and computers from the bank?", "Clear Bank", MessageBoxButtons.YesNo)
                Select Case DialogResult
                    Case Windows.Forms.DialogResult.Yes
                        BankPeriodStudents = New Collection
                        ReloadBank(Val(tclBank.SelectedTab.Tag))
                        Exit Sub
                    Case Windows.Forms.DialogResult.No
                        Exit Sub
                End Select
            Else
                ReloadBank(Val(TabControl.SelectedTab.Tag))
            End If
        ElseIf Val(TabControl.Tag) = 1 Then 'main
            If Val(TabControl.SelectedTab.Tag) > 0 Then
                ReloadRow(Val(TabControl.SelectedTab.Tag))
            Else
                ReloadOverview()
            End If
        End If
    End Sub

    Private Sub ClassDesigner_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboNodeType.SelectedIndex = 0
        tclBank.Anchor = AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Bottom
        tclMain.Anchor = AnchorStyles.Left + AnchorStyles.Bottom + AnchorStyles.Right + AnchorStyles.Top
        tclMain.TabPages(0).AutoScroll = True
        For Each BankTab As TabPage In tclBank.TabPages
            BankTab.AutoScroll = True
        Next
        Me.MinimumSize = New System.Drawing.Size(New Point(Me.Width, Me.Height))
        Application.DoEvents()
        AddStaticHandlers()
        ReloadVisibleInterface()
    End Sub

    Private Sub AddDimension(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddRow.Click, btnAddSeat.Click
        Try
            If FileOpen = False Then
                MessageBox.Show("Please create a new period or open an exiting one first.", "Error")
                Exit Sub
            End If
            Dim btnAdd As Button = sender
            Dim ButtonTag As String = btnAdd.Tag
            Select Case ButtonTag
                Case "X"    'add row
                    Dimensions = New Point(Dimensions.X + 1, Dimensions.Y)
                    PeriodInfo.Dimensions = Dimensions
                    AddRow(Dimensions.X)
                    For y As Integer = 1 To Dimensions.Y
                        Dim PeriodStudent As New PeriodStudent(StudentNodeType.Placeholder, Dimensions.X & "," & y, New Point(Dimensions.X, y), evtRemoveNode, evtNodeAction)
                        MainPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.Position.X & "," & PeriodStudent.Data.Position.Y)
                    Next
                    ReloadVisibleInterface()
                Case "Y"    'add column
                    Dimensions = New Point(Dimensions.X, Dimensions.Y + 1)
                    PeriodInfo.Dimensions = Dimensions
                    For x As Integer = 1 To Dimensions.X
                        Dim PeriodStudent As New PeriodStudent(StudentNodeType.Placeholder, x & "," & Dimensions.Y, New Point(x, Dimensions.Y), evtRemoveNode, evtNodeAction)
                        MainPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.Position.X & "," & PeriodStudent.Data.Position.Y)
                    Next
                    ReloadVisibleInterface()
            End Select
        Catch ex As Exception
            RaiseError(ErrorType.AddDimensionFailed)
        End Try
    End Sub

    Private Sub RemoveDimension(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveRow.Click, btnRemoveSeat.Click
        Try
            If FileOpen = False Then
                MessageBox.Show("Please create a new period or open an exiting one first.", "Error")
                Exit Sub
            End If
            Dim btnRemove As Button = sender
            Dim ButtonTag As String = btnRemove.Tag
            Select Case ButtonTag
                Case "X"    'remove row
                    Dim DialogResult As DialogResult = MessageBox.Show("This will remove the outermost row. Do you wish to continue?", "Remove Row", MessageBoxButtons.YesNo)
                    If DialogResult = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                    Dimensions = New Point(Dimensions.X - 1, Dimensions.Y)
                    PeriodInfo.Dimensions = Dimensions
                    'tclMain.TabPages.RemoveByKey(Dimensions.X + 1)
                    RemoveTabPage(tclMain.TabPages, Dimensions.X + 1)
                    For y As Integer = 1 To Dimensions.Y
                        MainPeriodStudents.Remove(Dimensions.X + 1 & "," & y)
                    Next
                    ReloadVisibleInterface()
                Case "Y"    'remove column
                    Dim DialogResult As DialogResult = MessageBox.Show("This will remove the outermost seat in each row. Do you wish to continue?", "Remove Seat", MessageBoxButtons.YesNo)
                    If DialogResult = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                    Dimensions = New Point(Dimensions.X, Dimensions.Y - 1)
                    PeriodInfo.Dimensions = Dimensions
                    For x As Integer = 1 To Dimensions.X
                        MainPeriodStudents.Remove(x & "," & Dimensions.Y + 1)
                    Next
                    ReloadVisibleInterface()
            End Select
        Catch ex As Exception
            RaiseError(ErrorType.RemoveDimensionFailed)
        End Try
    End Sub

#Region "Menu Events"

    Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveAs.Click
        SaveAsPeriod()
    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        SavePeriod()
    End Sub

    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        If FileOpen = True Then
            Dim DialogResult As DialogResult = MessageBox.Show("A class period is currently open. Would you like to save changes before closing it?", "Close", MessageBoxButtons.YesNoCancel)
            Select Case DialogResult
                Case Windows.Forms.DialogResult.Yes
                    SavePeriod()
                Case Windows.Forms.DialogResult.No
                    ClosePeriod()
                    'Continue Below
                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub
            End Select
        End If
        If stcPeriodDialog.bolOpen = True Then
            stcPeriodDialog.Form.Close()
            stcPeriodDialog.bolOpen = False
        End If
        stcPeriodDialog.Form = New PeriodDialog
        With stcPeriodDialog.Form
            .Owner = Me
            .Mode = PeriodDialog.PeriodMode.NewPeriod
            .Show()
            stcPeriodDialog.bolOpen = True
        End With
    End Sub

    Private Sub mnuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClose.Click
        If FileOpen = True Then
            Dim DialogResult As DialogResult = MessageBox.Show("A class period is currently open. Would you like to save changes before closing it?", "Close", MessageBoxButtons.YesNoCancel)
            Select Case DialogResult
                Case Windows.Forms.DialogResult.Yes
                    SavePeriod()
                Case Windows.Forms.DialogResult.No
                    'ClosePeriod() done below
                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub
            End Select
        End If
        ClosePeriod()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        If FileOpen = True Then
            Dim DialogResult As DialogResult = MessageBox.Show("A class period is currently open. Would you like to save changes before closing it?", "Close", MessageBoxButtons.YesNoCancel)
            Select Case DialogResult
                Case Windows.Forms.DialogResult.Yes
                    SavePeriod()
                Case Windows.Forms.DialogResult.No
                    ClosePeriod()
                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub
            End Select
        Else
            ClosePeriod()
        End If
        Me.Close()
    End Sub

    Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        If FileOpen = True Then
            Dim DialogResult As DialogResult = MessageBox.Show("A class period is currently open. Would you like to save changes before closing it?", "Close", MessageBoxButtons.YesNoCancel)
            Select Case DialogResult
                Case Windows.Forms.DialogResult.Yes
                    SavePeriod()
                    ClosePeriod()
                Case Windows.Forms.DialogResult.No
                    ClosePeriod()
                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub
            End Select
        Else
            ClosePeriod()
        End If
        OpenPeriod()
    End Sub

    Private Sub btnEditPeriodInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditPeriodInfo.Click
        If stcPeriodDialog.bolOpen = True Then
            stcPeriodDialog.Form.Close()
            stcPeriodDialog.bolOpen = False
        End If
        stcPeriodDialog.Form = New PeriodDialog
        With stcPeriodDialog.Form
            .Owner = Me
            .Mode = PeriodDialog.PeriodMode.EditPeriod
            .Show()
            stcPeriodDialog.bolOpen = True
        End With
    End Sub

#End Region
#End Region
#End Region

#Region "Interface Methods"

    Private Sub UpdateMenuControls()
        If FileOpen = True Then
            mnuSave.Enabled = True
            mnuSaveAs.Enabled = True
            mnuClose.Enabled = True
        Else
            mnuSave.Enabled = False
            mnuSaveAs.Enabled = False
            mnuClose.Enabled = False
        End If
    End Sub

    Private Sub UpdateNodeCaptions()
        If SelectedStudent.bolSelected = True Then
            If SelectedStudent.PeriodStudent.Data.Position.X = 0 And SelectedStudent.PeriodStudent.Data.Position.Y = 0 And SelectedStudent.PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                For Each PeriodStudent As PeriodStudent In MainPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        PeriodStudent.Controls.grpMain.Text = cPSCAPTIONSWAP
                    End If
                Next
                For Each PeriodStudent As PeriodStudent In BankPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        PeriodStudent.Controls.grpMain.Text = cPSCAPTIONNONE
                    End If
                Next
                StaticBankNode.Controls.grpMain.Text = cPSCAPTIONNONE
            ElseIf SelectedStudent.PeriodStudent.Data.Position.X = 0 And SelectedStudent.PeriodStudent.Data.Position.Y = 0 Then
                'static bank tab selected
                For Each PeriodStudent As PeriodStudent In MainPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        If PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                            PeriodStudent.Controls.grpMain.Text = cPSCAPTIONTOBANK
                        Else
                            PeriodStudent.Controls.grpMain.Text = cPSCAPTIONNONE
                        End If
                    End If
                Next
                For Each PeriodStudent As PeriodStudent In BankPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        PeriodStudent.Controls.grpMain.Text = cPSCAPTIONNONE
                    End If
                Next
                StaticBankNode.Controls.grpMain.Text = cPSCAPTIONDESELECT
            Else
                For Each PeriodStudent As PeriodStudent In BankPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        If PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                            PeriodStudent.Controls.grpMain.Text = cPSCAPTIONSWAP
                        Else
                            PeriodStudent.Controls.grpMain.Text = cPSCAPTIONNONE
                        End If
                    End If
                Next
                For Each PeriodStudent As PeriodStudent In MainPeriodStudents
                    If PeriodStudent.Data.NodeName <> SelectedStudent.PeriodStudent.Data.NodeName Then
                        PeriodStudent.Controls.grpMain.Text = cPSCAPTIONSWAP
                    End If
                Next
                If SelectedStudent.PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder Then
                    StaticBankNode.Controls.grpMain.Text = cPSCAPTIONTOBANK
                Else
                    StaticBankNode.Controls.grpMain.Text = cPSCAPTIONNONE
                End If
            End If
            SelectedStudent.PeriodStudent.Controls.grpMain.Text = cPSCAPTIONDESELECT
        Else
            'For Each PeriodStudentCollection As Collection In PeriodStudentCollections
            For Each PeriodStudent As PeriodStudent In BankPeriodStudents
                PeriodStudent.Controls.grpMain.Text = cPSCAPTIONSELECT
            Next
            For Each PeriodStudent As PeriodStudent In MainPeriodStudents
                PeriodStudent.Controls.grpMain.Text = cPSCAPTIONSELECT
            Next
            'Next
            StaticBankNode.Controls.grpMain.Text = cPSCAPTIONSELECT
        End If
    End Sub
    Private Sub ClearSelections()
        SelectedStudent.bolSelected = False
        SelectedStudent.PeriodStudent = Nothing
        'For Each PeriodStudentCollection As Collection In PeriodStudentCollections
        For Each PeriodStudent As PeriodStudent In BankPeriodStudents
            PeriodStudent.Controls.grpMain.BackColor = SystemColors.Control
        Next
        For Each PeriodStudent As PeriodStudent In MainPeriodStudents
            PeriodStudent.Controls.grpMain.BackColor = SystemColors.Control
        Next
        'Next
        StaticBankNode.Controls.grpMain.BackColor = SystemColors.Control
        ReloadVisibleInterface()
    End Sub
    Private Sub ClearNodeInput()
        txtNodeLastName.Text = ""
        txtNodeIDNumber.Text = ""
        txtNodeComputerName.Text = ""
    End Sub
    Private Sub SetupClass()
        For i As Integer = 1 To Dimensions.X
            AddRow(i)
        Next
        'Fill with blanks
        For x As Integer = 1 To Dimensions.X
            For y As Integer = 1 To Dimensions.Y
                Dim PeriodStudent As New PeriodStudent(StudentNodeType.Placeholder, x & "," & y, New Point(x, y), evtRemoveNode, evtNodeAction)
                MainPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.Position.X & "," & PeriodStudent.Data.Position.Y)
            Next
        Next
        ReloadBank(StudentNodeType.Student)
        ReloadOverview()
    End Sub

#Region "Add/Remove Methods"
    Private Sub AddRow(ByVal RowNumber As Integer)
        Dim NewTabPage As New TabPage("Row " & RowNumber)
        With NewTabPage
            .Name = RowNumber.ToString
            .Tag = RowNumber
            .AutoScroll = True
        End With
        tclMain.TabPages.Add(NewTabPage)
    End Sub
    Private Sub AddBankNode(ByVal Type As StudentNodeType, ByRef PeriodStudent As PeriodStudent, ByRef index As Integer)
        Dim BankTab As TabPage = tclBank.TabPages(Type)
        PeriodStudent.Controls.grpMain.Left = cBANKPSLEFTSPACING
        PeriodStudent.Controls.grpMain.Top = ((cBANKPSTOPSPACING + cPSGRPMAINHEIGHT) * index) + cBANKPSTOPSPACING
        BankTab.Controls.Add(PeriodStudent.Controls.grpMain)
        index += 1
    End Sub
    Private Sub AddRowNode(ByRef PeriodStudent As PeriodStudent)
        Dim RowNumber As Integer = PeriodStudent.Data.Position.X
        Dim MainTab As TabPage = tclMain.TabPages(RowNumber)
        PeriodStudent.Controls.grpMain.Left = Int((MainTab.Width - cPSGRPMAINWIDTH) / 2)
        If (MainTab.Height - ((cPSGRPMAINHEIGHT + cMAINPSTOPSPACING) * Dimensions.Y)) < cMAINPSTOPSPACING Then
            PeriodStudent.Controls.grpMain.Top = Int(((cPSGRPMAINHEIGHT + cMAINPSTOPSPACING) * (PeriodStudent.Data.Position.Y - 1)))
        Else
            PeriodStudent.Controls.grpMain.Top = Int((((MainTab.Height - ((cPSGRPMAINHEIGHT + cMAINPSTOPSPACING) * Dimensions.Y))) / 2) + ((cPSGRPMAINHEIGHT + cMAINPSTOPSPACING) * (PeriodStudent.Data.Position.Y - 1)))
        End If
        MainTab.Controls.Add(PeriodStudent.Controls.grpMain)
    End Sub
#End Region

#Region "Reload Methods"
    Private Sub ReloadBank(ByVal Type As StudentNodeType)
        Dim index As Integer = 0
        tclBank.TabPages(Type).Controls.Clear()
        AddBankNode(Type, StaticBankNode, index)
        For Each PeriodStudent As PeriodStudent In BankPeriodStudents
            If PeriodStudent.Data.NodeType = Type Then
                AddBankNode(Type, PeriodStudent, index)
            End If
        Next
        UpdateNodeCaptions()
    End Sub
    Private Sub ReloadRow(ByVal RowNumber As Integer)
        tclMain.TabPages(RowNumber).Controls.Clear()
        For Each PeriodStudent As PeriodStudent In MainPeriodStudents
            If PeriodStudent.Data.Position.X = RowNumber Then
                AddRowNode(PeriodStudent)
            End If
        Next
        UpdateNodeCaptions()
    End Sub
    Private Sub ReloadVisibleInterface()
        If Val(tclMain.SelectedTab.Tag) > 0 Then
            ReloadRow(Val(tclMain.SelectedTab.Tag))
        Else
            ReloadOverview()
        End If
        ReloadBank(Val(tclBank.SelectedTab.Tag))
    End Sub
    Private Sub ReloadOverview()
        tclMain.TabPages(0).Controls.Clear()
        For Each PeriodStudent As PeriodStudent In MainPeriodStudents
            Dim grpMain As New GroupBox
            Dim lblName As New Label
            Dim DisplayName As String = System.Text.RegularExpressions.Regex.Replace(PeriodStudent.Data.NodeName, "[^a-z,A-Z]", "")
            If DisplayName.Length > 8 Then
                DisplayName = DisplayName.Substring(0, 8) & "..."
            End If
            If PeriodStudent.Data.NodeType = StudentNodeType.Student Then
                DisplayName = DisplayName & vbCrLf & System.Text.RegularExpressions.Regex.Replace(PeriodStudent.Data.NodeName, "[^0-9]", "")
            End If
            With grpMain
                .Width = cOVERVIEWGRPMAINWIDTH
                .Height = cOVERVIEWGRPMAINHEIGHT
                .Left = ((cOVERVIEWGRPMAINROWSPACING + 100) * (PeriodStudent.Data.Position.X - 1)) + cOVERVIEWGRPMAINROWSPACING
                .Top = ((cOVERVIEWGRPMAINCOLUMNSPACING + 100) * (PeriodStudent.Data.Position.Y - 1)) + cOVERVIEWGRPMAINCOLUMNSPACING
            End With
            With lblName
                .Width = cOVERVIEWLBLNAMEWIDTH
                .Height = cOVERVIEWLBLNAMEHEIGHT
                .Left = cOVERVIEWLBLMAINLEFT
                .Top = cOVERVIEWLBLMAINTOP
                .TextAlign = ContentAlignment.MiddleCenter
                .Text = IIf(PeriodStudent.Data.NodeType <> StudentNodeType.Placeholder, DisplayName, "")
            End With
            grpMain.Controls.Add(lblName)
            tclMain.TabPages(0).Controls.Add(grpMain)
        Next
    End Sub
#End Region
#End Region

#Region "Tool Methods"
    Private Function IsStaticBankNode(ByVal PeriodStudent As PeriodStudent) As Boolean
        If PeriodStudent.Data.Position.X = 0 And PeriodStudent.Data.Position.Y = 0 Then
            If PeriodStudent.Data.NodeType = StudentNodeType.Placeholder Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
    Private Function GetStudentType(ByVal Name As String) As StudentNodeLocation
        If CollectionContains(BankPeriodStudents, Name) = True Then
            Return 0
        ElseIf CollectionContains(MainPeriodStudents, Name) = True Then
            Return 1
        Else
            Return -1
        End If
    End Function
    Private Sub PurgeBank()
        For Each BankItem As PeriodStudent In BankPeriodStudents
            If BankItem.Data.NodeType = StudentNodeType.Placeholder Then
                BankPeriodStudents.Remove(BankItem.Data.NodeName)
            End If
        Next
    End Sub
    Private Sub MoveNodeToBank(ByVal PeriodStudent As PeriodStudent)
        Try
            Dim AddPeriodStudent As New PeriodStudent(PeriodStudent.Data.NodeType, PeriodStudent.Data.NodeName, New Point, evtRemoveNode, evtNodeAction)
            BankPeriodStudents.Add(AddPeriodStudent, PeriodStudent.Data.NodeName)
            ReloadBank(Val(tclBank.SelectedTab.Tag))
        Catch ex As Exception
            RaiseError(ErrorType.NodeSwapFailed)
        End Try
    End Sub

    Private Function CollectionPeriodStudentExists(ByVal Collection As Collection, ByVal ItemKey As String) As Boolean
        Try
            Dim PeriodStudent As PeriodStudent
            PeriodStudent = Collection.Item(ItemKey)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetStudentCollection(ByVal intType As Integer) As Collection
        Select Case intType
            Case 0
                Return BankPeriodStudents
            Case 1
                Return MainPeriodStudents
            Case Else
                RaiseError(ErrorType.Unknown)
                Return New Collection
                Exit Function
        End Select
    End Function

#End Region

#Region "Validation Methods"
    Private Function ValidateNodeInput(ByVal NodeType As StudentNodeType) As Boolean
        Select Case NodeType
            Case StudentNodeType.Student
                If txtNodeLastName.Text <> "" And txtNodeIDNumber.Text <> "" Then
                    Return True
                Else
                    Return False
                End If
            Case StudentNodeType.Computer
                If txtNodeComputerName.Text <> "" Then
                    Return True
                Else
                    Return False
                End If
        End Select
    End Function
    Private Function RaiseError(ByVal Type As ErrorType) As DialogResult
        Dim Response As DialogResult
        Select Case Type
            Case ErrorType.InvalidNodeType
                MessageBox.Show("Invalid node type selected.", "Error")
            Case ErrorType.InvalidNodeInfo
                MessageBox.Show("Invalid node information entered.", "Error")
            Case ErrorType.DuplicateNodeName
                MessageBox.Show("You cannot add two identical nodes.", "Error")
            Case ErrorType.NodeRemoveFailed
                MessageBox.Show("Failed to remove the node.", "Error")
            Case ErrorType.NodeSelectFailed
                MessageBox.Show("Failed to select the node.", "Error")
            Case ErrorType.NodeSwapFailed
                MessageBox.Show("Failed to swap nodes.", "Error")
            Case ErrorType.AddDimensionFailed
                MessageBox.Show("Failed to add the row or seat.", "Error")
            Case ErrorType.RemoveDimensionFailed
                MessageBox.Show("Failed to remove the row or seat.", "Error")
            Case ErrorType.SavePeriodFailed
                MessageBox.Show("Failed to save the class period file.", "Error")
            Case ErrorType.FileOpenFailed
                MessageBox.Show("Failed to open the class period file.", "Error")
            Case ErrorType.LoadPeriodFailed
                MessageBox.Show("Failed to load the class period file. Usually this indicates a corrupt or invalid class period file.", "Error")
            Case ErrorType.Unknown
                MessageBox.Show("An unknown error has occurred.", "Error")
        End Select
        Return Response
    End Function
#End Region

#Region "File Methods"
    Private Function SavePeriod(ByVal PeriodStudentCollection As Collection, ByVal PeriodInfo As ClassPeriodInfo, ByVal FilePath As String) As Boolean
        Dim FileString As String
        Try
            FileString = "[" & PeriodInfo.Dimensions.X & "," & PeriodInfo.Dimensions.Y & "][" & PeriodInfo.PeriodNumber & "][""" & PeriodInfo.Title & """]"
            PeriodStudentCollection = SortPeriodStudentCollection(PeriodStudentCollection)
            For Each PeriodStudent As PeriodStudent In PeriodStudentCollection
                FileString = FileString & vbNewLine & "[" & PeriodStudent.Data.Position.X & "," & PeriodStudent.Data.Position.Y & "][" & PeriodStudent.Data.NodeType & "][""" & PeriodStudent.Data.NodeName & """]"
            Next
            Dim FileStream As New System.IO.StreamWriter(FilePath)
            FileStream.Write(FileString)
            FileStream.Close()
            LastSavePath = FilePath
        Catch ex As Exception
            RaiseError(ErrorType.SavePeriodFailed)
        End Try
    End Function

    Private Function SortPeriodStudentCollection(ByVal PeriodStudentCollection As Collection) As Collection
        Dim x As Integer = 1
        Dim y As Integer = 1
        Dim SortedCollection As New Collection
        Dim bolFound As Boolean
        Dim count As Integer = 0
SearchXY:
        For Each PeriodStudent As PeriodStudent In PeriodStudentCollection
            If PeriodStudent.Data.Position.X = x Then
                If PeriodStudent.Data.Position.Y = y Then
                    SortedCollection.Add(PeriodStudent, x & "," & y)
                    bolFound = True
                End If
            End If
        Next
        If bolFound = True Then
            count += 1
            y += 1
            bolFound = False
            GoTo SearchXY
        Else
            If count > 0 Then
                count = 0
                x += 1
                y = 1
                GoTo SearchXY
            Else
                GoTo Sorted
            End If
        End If
Sorted:
        Return SortedCollection
    End Function

    Private Sub SavePeriod()
        If LastSavePath <> "" Then
            SavePeriod(MainPeriodStudents, PeriodInfo, LastSavePath)
        Else
            SaveAsPeriod()
        End If
    End Sub

    Private Sub SaveAsPeriod()
        Dim SavePath As String
        With dlgSaveFile
            .AddExtension = True
            .FileName = "Period " & PeriodInfo.PeriodNumber '""
            .DefaultExt = ".swp"
            .Filter = "ScreenWatcher Period File (.swp)|*.swp"
            .FilterIndex = 0
            .ShowDialog()
        End With
        SavePath = dlgSaveFile.FileName
        If SavePath <> "" Then
            SavePeriod(MainPeriodStudents, PeriodInfo, SavePath)
        End If
        dlgSaveFile.FileName = ""
    End Sub

    Private Sub ClosePeriod()
        Dimensions = Nothing
        PeriodInfo = Nothing
        MainPeriodStudents = New Collection
        For Each RowTab As TabPage In tclMain.TabPages
            If Val(RowTab.Tag) > 0 Then
                RemoveTabPage(tclMain.TabPages, RowTab.Tag)
            End If
        Next
        FileOpen = False
        ReloadVisibleInterface()
    End Sub

    Public Sub NewPeriod()
        Dimensions = PeriodInfo.Dimensions
        FileOpen = True
        SetupClass()
    End Sub

    Private Sub OpenPeriod()
        Dim OpenPath As String
        Dim strFileContents As String
        Try
            With dlgOpenFile
                .AddExtension = True
                .FileName = ""
                .DefaultExt = ".swp"
                .Filter = "ScreenWatcher Period File (.swp)|*.swp"
                .FilterIndex = 0
                .ShowDialog()
            End With
            OpenPath = dlgOpenFile.FileName
            If OpenPath <> "" Then
                LastSavePath = OpenPath
                Dim FileStream As New System.IO.StreamReader(OpenPath)
                strFileContents = FileStream.ReadToEnd
                FileStream.Close()
                LoadPeriod(strFileContents)
                ReloadVisibleInterface()
            End If
            dlgOpenFile.FileName = ""
        Catch ex As Exception
            RaiseError(ErrorType.FileOpenFailed)
        End Try
    End Sub

    Private Sub LoadPeriod(ByVal FileContents As String)
        Try
            Dim strLines() As String = FileContents.Split(vbCrLf)
            Dim intLine As Integer = 1
            For Each Line As String In strLines
                Dim strParts() As String = Line.Split("[")
                Dim intPart As Integer = 0
                Dim PeriodStudentData As PeriodStudentData
                With PeriodStudentData
                    .NodeName = ""
                    .NodeType = StudentNodeType.Placeholder
                    .Position = New Point
                End With
                For Each Part As String In strParts
                    Dim strTemp() As String = Part.Split("]")
                    Dim strPart As String = strTemp(0)
                    Select Case intPart
                        Case 1  'dimensions/position
                            If intLine = 1 Then
                                Dim SubParts() As String = strPart.Split(",")
                                Dimensions = New Point(SubParts(0), SubParts(1))
                                PeriodInfo.Dimensions = Dimensions
                                'load rows
                                For intCount As Integer = 1 To Dimensions.X
                                    AddRow(intCount)
                                Next
                            Else
                                Dim SubParts() As String = strPart.Split(",")
                                PeriodStudentData.Position = New Point(SubParts(0), SubParts(1))
                            End If
                        Case 2  'period/type
                            If intLine = 1 Then
                                PeriodInfo.PeriodNumber = Val(strPart)
                            Else
                                PeriodStudentData.NodeType = Val(strPart)
                            End If
                        Case 3  'title/name
                            If intLine = 1 Then
                                PeriodInfo.Title = strPart.TrimStart("""").TrimEnd("""")
                            Else
                                PeriodStudentData.NodeName = strPart.TrimStart("""").TrimEnd("""")
                            End If
                    End Select
                    intPart += 1
                Next
                If intLine > 1 Then
                    Dim PeriodStudent As New PeriodStudent(PeriodStudentData.NodeType, PeriodStudentData.NodeName, PeriodStudentData.Position, evtRemoveNode, evtNodeAction)
                    MainPeriodStudents.Add(PeriodStudent, PeriodStudent.Data.NodeName)
                End If
                intLine += 1
            Next
            FileOpen = True
        Catch ex As Exception
            RaiseError(ErrorType.LoadPeriodFailed)
        End Try
    End Sub
#End Region

#End Region

End Class
