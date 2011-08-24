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

Public Class SavedImageViewer
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
    Friend WithEvents trvMain As System.Windows.Forms.TreeView
    Friend WithEvents tabMain As System.Windows.Forms.TabControl
    Friend WithEvents tabFileView As System.Windows.Forms.TabPage
    Friend WithEvents tabImageView As System.Windows.Forms.TabPage
    Friend WithEvents picLarge As System.Windows.Forms.PictureBox
    Friend WithEvents imgList As System.Windows.Forms.ImageList
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SavedImageViewer))
        Me.trvMain = New System.Windows.Forms.TreeView
        Me.tabMain = New System.Windows.Forms.TabControl
        Me.tabFileView = New System.Windows.Forms.TabPage
        Me.tabImageView = New System.Windows.Forms.TabPage
        Me.picLarge = New System.Windows.Forms.PictureBox
        Me.imgList = New System.Windows.Forms.ImageList(Me.components)
        Me.tabMain.SuspendLayout()
        Me.tabImageView.SuspendLayout()
        Me.SuspendLayout()
        '
        'trvMain
        '
        Me.trvMain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.trvMain.ImageIndex = -1
        Me.trvMain.Indent = 18
        Me.trvMain.Location = New System.Drawing.Point(15, 15)
        Me.trvMain.Name = "trvMain"
        Me.trvMain.SelectedImageIndex = -1
        Me.trvMain.Size = New System.Drawing.Size(175, 400)
        Me.trvMain.TabIndex = 0
        '
        'tabMain
        '
        Me.tabMain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tabMain.Controls.Add(Me.tabFileView)
        Me.tabMain.Controls.Add(Me.tabImageView)
        Me.tabMain.Location = New System.Drawing.Point(205, 15)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SelectedIndex = 0
        Me.tabMain.Size = New System.Drawing.Size(430, 400)
        Me.tabMain.TabIndex = 1
        '
        'tabFileView
        '
        Me.tabFileView.AutoScroll = True
        Me.tabFileView.Location = New System.Drawing.Point(4, 25)
        Me.tabFileView.Name = "tabFileView"
        Me.tabFileView.Size = New System.Drawing.Size(422, 371)
        Me.tabFileView.TabIndex = 0
        Me.tabFileView.Text = "File View"
        '
        'tabImageView
        '
        Me.tabImageView.Controls.Add(Me.picLarge)
        Me.tabImageView.Location = New System.Drawing.Point(4, 25)
        Me.tabImageView.Name = "tabImageView"
        Me.tabImageView.Size = New System.Drawing.Size(422, 371)
        Me.tabImageView.TabIndex = 1
        Me.tabImageView.Text = "Image View"
        '
        'picLarge
        '
        Me.picLarge.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picLarge.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picLarge.Location = New System.Drawing.Point(8, 8)
        Me.picLarge.Name = "picLarge"
        Me.picLarge.Size = New System.Drawing.Size(406, 360)
        Me.picLarge.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLarge.TabIndex = 0
        Me.picLarge.TabStop = False
        '
        'imgList
        '
        Me.imgList.ImageSize = New System.Drawing.Size(23, 23)
        Me.imgList.ImageStream = CType(resources.GetObject("imgList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgList.TransparentColor = System.Drawing.Color.Transparent
        '
        'SavedImageViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(642, 430)
        Me.Controls.Add(Me.tabMain)
        Me.Controls.Add(Me.trvMain)
        Me.MinimumSize = New System.Drawing.Size(625, 470)
        Me.Name = "SavedImageViewer"
        Me.Text = "Saved Image Viewer"
        Me.tabMain.ResumeLayout(False)
        Me.tabImageView.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constants"
    Const cSPACING As Integer = 15

    Const cTRVMAINLEFT As Integer = cSPACING
    Const cTRVMAINTOP As Integer = cSPACING
    Const cTRVMAINWIDTH As Integer = 175
    Const cTRVMAINHEIGHT As Integer = 400

    Const cTABMAINLEFT As Integer = cSPACING + cTRVMAINWIDTH + cSPACING
    Const cTABMAINTOP As Integer = cSPACING
    Const CTABMAINWIDTH As Integer = 430
    Const cTABMAINHEIGHT As Integer = 400

    Const cFRMWIDTH As Integer = cTABMAINLEFT + CTABMAINWIDTH + cSPACING
    Const cFRMHEIGHT As Integer = cTABMAINTOP + cTABMAINHEIGHT + cSPACING



    Const cSTARTTOP As Integer = 20
    Const cSTARTLEFT As Integer = 10

    Const cROWSPACING As Integer = 190
    Const cCOLUMNSPACING As Integer = 200

    Const cBTNDELETEWIDTH As Integer = 23
    Const cBTNDELETEHEIGHT As Integer = 23
    Const cBTNDELETETOP As Integer = 150
    Const cBTNDELETELEFT As Integer = 10

    Const cPICTUREBOXWIDTH As Integer = 170
    Const cPICTUREBOXHEIGHT As Integer = 128
    Const cPICTUREBOXTOP As Integer = 15
    Const cPICTUREBOXLEFT As Integer = 10

    Const cGRPBOXWIDTH As Integer = 190
    Const cGRPBOXHEIGHT As Integer = 180

    Const cLBLTIMEWIDTH As Integer = 140
    Const cLBLTIMEHEIGHT As Integer = 23
    Const cLBLTIMETOP As Integer = 150
    Const cLBLTIMELEFT As Integer = 40

#End Region

    Private Sub ThumbnailViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.trvMain.Left = cTRVMAINLEFT
        Me.trvMain.Top = cTRVMAINTOP
        Me.trvMain.Width = cTRVMAINWIDTH
        Me.trvMain.Height = cTRVMAINHEIGHT

        Me.tabMain.Top = cTABMAINTOP
        Me.tabMain.Left = cTABMAINLEFT
        Me.tabMain.Width = CTABMAINWIDTH
        Me.tabMain.Height = cTABMAINHEIGHT

        Me.Width = cFRMWIDTH
        Me.Height = cFRMHEIGHT

        Me.trvMain.Anchor = AnchorStyles.Bottom + AnchorStyles.Left + AnchorStyles.Top
        Me.tabMain.Anchor = AnchorStyles.Bottom + AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Right

        Me.MinimumSize = New Size(New Point(cFRMWIDTH, cFRMHEIGHT))

        Dim picLargeToolTip As New ToolTip
        picLargeToolTip.SetToolTip(Me.picLarge, "Click to Enlarge")

        LoadFolderTree(Application.StartupPath & "\Saved Images")
    End Sub

    Private Function FileGetName(ByVal strFullPath As String) As String
        Dim parts() As String = strFullPath.Split("\")
        Return parts(UBound(parts))
    End Function

    Private Function FileGetPath(ByVal strFullPath As String) As String
        Return strFullPath.Substring(0, strFullPath.LastIndexOf("\"))
    End Function

    Private Function ImageFromFile(ByVal strPath As String) As Image
        Dim imgMain As Image
        Dim FileStream As New System.IO.FileStream(strPath, IO.FileMode.Open, IO.FileAccess.Read)
        imgMain = Image.FromStream(FileStream)
        FileStream.Close()
        Return imgMain
    End Function

    Public Sub LoadFolderTree(ByVal path As String)
        Dim basenode As System.Windows.Forms.TreeNode
        If IO.Directory.Exists(path) Then
            If path.Length <= 3 Then
                basenode = trvMain.Nodes.Add(path)
            Else
                basenode = trvMain.Nodes.Add(FileGetName(path))
            End If
            basenode.Tag = path
            LoadDir(path, basenode)
        Else
            Throw New System.IO.DirectoryNotFoundException
        End If
    End Sub

    Private Sub trvmain_AfterExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvMain.AfterExpand
        Dim n As System.Windows.Forms.TreeNode
        For Each n In e.Node.Nodes
            LoadDir(n.Tag, n)
        Next
    End Sub

    Public Sub LoadDir(ByVal DirPath As String, ByVal Node As Windows.Forms.TreeNode)
        On Error Resume Next
        Dim Index As Integer
        If Node.Nodes.Count = 0 Then
            Dim Root As New System.IO.DirectoryInfo(DirPath)
            For Each dioDir As System.IO.DirectoryInfo In Root.GetDirectories("*.*")
                Dim strDir As String = dioDir.FullName.ToString
                Index = strDir.LastIndexOf("\")
                Node.Nodes.Add(strDir.Substring(Index + 1, strDir.Length - Index - 1))
                Node.LastNode.Tag = strDir
                Node.LastNode.ImageIndex = 0
            Next
        End If
    End Sub

    Public Sub LoadImages(ByVal strPath As String)
        tabMain.TabPages(0).Controls.Clear()
        Dim i As Integer
        Dim intTop As Integer = cSTARTTOP
        Dim intLeft As Integer = cSTARTLEFT
        Dim Root As New System.IO.DirectoryInfo(strPath)

        For Each fioFile As System.IO.FileInfo In Root.GetFiles("*.jpg")
            Dim File As String = fioFile.FullName.ToString
            Dim GroupBox As New GroupBox
            Dim PictureBox As New PictureBox
            Dim btnDelete As New Button
            Dim lblTime As New Label
            Dim picToolTip As New ToolTip
            Dim btnToolTip As New ToolTip
            Dim FileName() As String = File.Split("\")

            With btnDelete
                .Width = cBTNDELETEWIDTH
                .Height = cBTNDELETEHEIGHT
                .Top = cBTNDELETETOP
                .Left = cBTNDELETELEFT
                .BringToFront()
                .BackgroundImage = imgList.Images(0)
                .FlatStyle = FlatStyle.Flat
                .Tag = File
                .Show()
            End With

            With lblTime
                .Width = cLBLTIMEWIDTH
                .Height = cLBLTIMEHEIGHT
                .Top = cLBLTIMETOP
                .Left = cLBLTIMELEFT
                .TextAlign = ContentAlignment.MiddleRight
                .Text = StringToTime(FileName(UBound(FileName)))
                .Show()
            End With

            With GroupBox
                .Width = cGRPBOXWIDTH
                .Height = cGRPBOXHEIGHT
                .Top = intTop
                .Left = intLeft
                .Text = ""
                .Show()
            End With

            With PictureBox
                .Width = cPICTUREBOXWIDTH
                .Height = cPICTUREBOXHEIGHT
                .Top = cPICTUREBOXTOP
                .Left = cPICTUREBOXLEFT
                .SizeMode = PictureBoxSizeMode.StretchImage
                .Image = ImageFromFile(File)
                .Tag = File
                .Cursor = Cursors.Hand
                .Show()
            End With

            With picToolTip
                .SetToolTip(PictureBox, "Click to Enlarge")
            End With

            With btnToolTip
                .SetToolTip(btnDelete, "Delete Image")
            End With

            AddHandler btnDelete.Click, New EventHandler(AddressOf DeleteImage)
            AddHandler PictureBox.Click, New EventHandler(AddressOf EnlargeImage)
            GroupBox.Controls.Add(PictureBox)
            GroupBox.Controls.Add(btnDelete)
            GroupBox.Controls.Add(lblTime)
            tabMain.TabPages(0).Controls.Add(GroupBox) 'PictureBox)
            i = i + 1
            If i Mod 2 = 0 Then
                intTop += cROWSPACING
                intLeft -= cCOLUMNSPACING
            Else
                intLeft += cCOLUMNSPACING
            End If
        Next
    End Sub

    Private Function StringToTime(ByVal strMain As String) As String
        On Error GoTo InvalidName
        Dim parts() As String = strMain.Split("-")
        Dim strTime As String
        If Val(parts(0)) > 12 Then
            strTime = Val(parts(0)) - 12 & ":" & parts(1) & ":" & parts(2) & " PM"
        Else
            strTime = Val(parts(0)) & ":" & parts(1) & ":" & parts(2) & " AM"
        End If
        Return strTime
        Exit Function
InvalidName:
        Return "Unknown"
    End Function

    Private Sub trvMain_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvMain.AfterSelect
        LoadImages(e.Node.Tag)
        tabMain.SelectedTab = tabMain.TabPages(0)
    End Sub

    Private Sub grpMain_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub EnlargeImage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim picEnlarge As PictureBox = sender
        Dim imgMain As Image = ImageFromFile(picEnlarge.Tag)
        picLarge.Image = imgMain
        If imgMain.Width > picLarge.Width Or imgMain.Height > picLarge.Height Then
            picLarge.SizeMode = PictureBoxSizeMode.StretchImage
        Else
            picLarge.SizeMode = PictureBoxSizeMode.CenterImage
        End If
        tabMain.SelectedTab = tabMain.TabPages(1)
    End Sub

    Private Sub DeleteImage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim btnDelete As Button = sender
        Dim DialogResult As DialogResult = MessageBox.Show("Are you sure you want to delete this image?", "Confirm File Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If DialogResult = Windows.Forms.DialogResult.Yes Then
            Try
                System.IO.File.Delete(btnDelete.Tag)
                LoadImages(FileGetPath(btnDelete.Tag))
            Catch ex As Exception
                MessageBox.Show("Unable to delete file.", "Error")
            End Try
        End If
    End Sub

    Private Sub picLarge_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles picLarge.Resize
        Dim imgMain As Image = picLarge.Image
        Try
            If imgMain.Width > picLarge.Width Or imgMain.Height > picLarge.Height Then
                picLarge.SizeMode = PictureBoxSizeMode.StretchImage
            Else
                picLarge.SizeMode = PictureBoxSizeMode.CenterImage
            End If
        Catch ex As Exception
            'No image set yet
        End Try
    End Sub

    Private Sub picLarge_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picLarge.MouseUp
        If e.Button = MouseButtons.Right Then
            Exit Sub
        End If

        Dim EnlargeForm As New EnlargeForm(Me.picLarge.Image)
        EnlargeForm.Owner = Me
        EnlargeForm.Show()

    End Sub

End Class
