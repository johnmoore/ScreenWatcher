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

Module Main
    Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

    Dim IP As String = "192.168.1.151"
    Dim strName As String
    Dim Student As StudentClass

    Dim MsgThread As Threading.Thread
    Dim WithEvents msgMain As Message

    Dim frmMsg As New Form
    Dim lblMsg As New Label

    Sub Main()
        On Error Resume Next

        InitMsgForm()

        strName = System.Security.Principal.WindowsIdentity.GetCurrent.Name & "\" & System.Environment.MachineName.ToString
        Student = New StudentClass(strName, IP)

        While True
            If Student.DataAvailable = True Then
                Dim buffer(Student.ReceiveBufferSize - 1) As Byte
                Dim TypeChar As Char = Student.GetDataType(buffer)

                Select Case TypeChar
                    Case "B"
                        BlockInput(True)
                    Case "F"
                        Student.Mode = Student.ClientMode.FullScreen
                    Case "M"
                        ReadMessage(Student, buffer)
                    Case "P"
                        'do nothing
                    Case "R"
                        Student.FullUpdate = True
                    Case "T"
                        Student.Mode = Student.ClientMode.Thumbnail
                    Case "U"
                        BlockInput(False)
                End Select
            End If

            Select Case Student.Mode
                Case StudentClass.ClientMode.Thumbnail
                    If Student.SendThumbnail() <> True Then
                        Student.ReconnectAttempt(IP)
                    End If
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(1000)

                Case StudentClass.ClientMode.FullScreen
                    If Student.FullUpdate = True Then
                        Student.FullUpdate = False
                        If Student.SendFullScreen() <> True Then
                            Student.ReconnectAttempt(IP)
                        End If
                        System.Threading.Thread.Sleep(2000)
                    Else
                        Dim TimeSpan As System.TimeSpan = Now.Subtract(Student.FullTimeStamp)
                        If TimeSpan.Seconds > 5 Then
                            Student.FullTimeStamp = Now
                            If Student.SendTypeByte("P") <> True Then
                                Student.ReconnectAttempt(IP)
                            End If
                        End If
                    End If
            End Select

        End While
    End Sub

    Private Sub ReadMessage(ByRef student As StudentClass, ByVal buffer As Byte())
        msgMain = New Message(student.GetMessage(buffer))

        Select Case msgMain.MessageType
            Case "M"
                ShowMessageBox(msgMain)
            Case "F"
                lblMsg.Text = msgMain.strTitle & IIf(msgMain.strTitle <> "", vbCrLf & vbCrLf, "") & msgMain.strMessage
                frmMsg.Show()
                frmMsg.BringToFront()
                Application.DoEvents()
            Case "X"
                lblMsg.Text = ""
                frmMsg.Hide()
        End Select
        If student.Mode = StudentClass.ClientMode.FullScreen Then
            student.FullUpdate = True
        End If
    End Sub

    Private Sub ShowMessageBox(ByVal Msg As Message)
        MsgThread = New Threading.Thread(AddressOf Msg.ShowMessage)
        MsgThread.Start()

    End Sub

    Private Sub MessageEnds() Handles msgMain.MessageComplete
        MsgThread.Abort()
    End Sub

    Public Class Message
        Private msgClient As StudentClass.ClientMessage

        Public ReadOnly Property strMessage() As String
            Get
                Return msgClient.strMessage
            End Get
        End Property

        Public ReadOnly Property strTitle() As String
            Get
                Return msgClient.strTitle
            End Get
        End Property

        Public ReadOnly Property MessageType() As Char
            Get
                Return msgClient.MessageType
            End Get
        End Property

        Public Sub New(ByVal MessageType As Char, ByVal Message As String, ByVal Title As String)
            msgClient.MessageType = MessageType
            msgClient.strMessage = Message
            msgClient.strTitle = Title

        End Sub

        Public Sub New(ByVal stcMessage As StudentClass.ClientMessage)
            msgClient.MessageType = stcMessage.MessageType
            msgClient.strMessage = stcMessage.strMessage
            msgClient.strTitle = stcMessage.strTitle
        End Sub

        Public Sub ShowMessage()
            MessageBox.Show(msgClient.strMessage, msgClient.strTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            RaiseEvent MessageComplete()
        End Sub

        Public Event MessageComplete()

    End Class

    Private Sub InitMsgForm()
        With frmMsg
            .Width = Screen.PrimaryScreen.Bounds.Width
            .Height = Screen.PrimaryScreen.Bounds.Height
            .Top = 0
            .Left = 0
            .TopMost = True
            .BackColor = Color.Black
            .FormBorderStyle = FormBorderStyle.None
        End With

        With lblMsg
            .Width = frmMsg.Width
            .Height = frmMsg.Height
            .Top = 0
            .Left = 0
            .TextAlign = ContentAlignment.MiddleCenter
            .Font = New Font(System.Drawing.FontFamily.GenericSerif, 20, FontStyle.Regular, GraphicsUnit.Point)
            .ForeColor = Color.White
            .Show()
        End With

        frmMsg.Controls.Add(lblMsg)
    End Sub
End Module
