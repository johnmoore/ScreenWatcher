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

Public Class Message

#Region "Fields"
    Enum Messagetype
        FullScreen
        Messagebox
    End Enum

    Private fldInitialized As Boolean
    Private fldMessage As String
    Private fldTitle As String
    Private fldMessageType As Messagetype
    Private fldClient As Student
    Private fldTimeSpan As TimeSpan

    Public Event MessageSent()
    Public Event MessageEnded()
#End Region

#Region "Properties"
    Public ReadOnly Property Student() As Student
        Get
            Return fldClient
        End Get
    End Property

    Public ReadOnly Property CurrentMessageType() As Messagetype
        Get
            Return fldMessageType
        End Get
    End Property

    Public ReadOnly Property TimeSpan() As TimeSpan
        Get
            Return fldTimeSpan
        End Get
    End Property

    Public ReadOnly Property Initialized() As Boolean
        Get
            Return fldInitialized
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        fldInitialized = False
    End Sub

    Public Sub New(ByRef Student As Student, ByVal MessageType As Messagetype, ByVal Message As String, ByVal Title As String, ByVal TimeSpan As TimeSpan)
        fldClient = Student
        fldMessageType = MessageType
        fldMessage = Message
        fldTitle = Title
        fldTimeSpan = TimeSpan
        fldInitialized = True
    End Sub

    Public Sub SendMessage()
        Select Case fldMessageType
            Case Messagetype.FullScreen
                SendTypeByte(fldClient, "B")
                TcpListenerModule.SendMessage(fldClient, "F", fldTitle, fldMessage)
                RaiseEvent MessageSent()

                If fldTimeSpan.Equals(TimeSpan.Parse("0")) Then
                    Exit Sub
                Else
                    EndMessage()
                End If

            Case Messagetype.Messagebox
                TcpListenerModule.SendMessage(fldClient, "M", fldTitle, fldMessage)
                RaiseEvent MessageSent()
        End Select

    End Sub

    Public Sub EndMessage()
        Select Case fldMessageType
            Case Messagetype.FullScreen
                TcpListenerModule.SendMessage(fldClient, "X", "", "")
                SendTypeByte(fldClient, "U")
            Case Messagetype.Messagebox
                'do nothing
        End Select

        RaiseEvent MessageEnded()
    End Sub

#End Region

End Class
