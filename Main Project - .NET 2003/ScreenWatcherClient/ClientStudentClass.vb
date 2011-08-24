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

Public Class StudentClass

#Region "Type Declares"

    Enum ClientMode
        Thumbnail
        FullScreen
    End Enum

    Public Structure ClientMessage
        Dim MessageType As Char
        Dim strTitle As String
        Dim strMessage As String
    End Structure

#End Region

#Region "Fields"
    Private fldName As String
    Private fldImage As Image
    Private fldFullTimeStamp As DateTime

    Private fldMode As ClientMode
    Private fldFullUpdate As Boolean = False

    Private fldNetworkStream As System.Net.Sockets.NetworkStream
    Private fldTcpClient As System.Net.Sockets.TcpClient
    Private fldClientFuncs As New ClientScreenFunctions

    Const intPort As Integer = 8000

#End Region 'Fields

#Region "Properties"

    Public ReadOnly Property ReceiveBufferSize() As Integer
        Get
            Return fldTcpClient.ReceiveBufferSize
        End Get
    End Property

    Public Property FullUpdate() As Boolean
        Get
            Return fldFullUpdate
        End Get
        Set(ByVal Value As Boolean)
            fldFullUpdate = Value
        End Set
    End Property

    Public Property Mode() As ClientMode
        Get
            Return fldMode
        End Get
        Set(ByVal Value As ClientMode)
            fldMode = Value
        End Set
    End Property

    Public ReadOnly Property DataAvailable() As Boolean
        Get
            Return fldNetworkStream.DataAvailable
        End Get
    End Property

    Public Property FullTimeStamp()
        Get
            Return fldFullTimeStamp
        End Get
        Set(ByVal Value)
            fldFullTimeStamp = Value
        End Set
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return fldName
        End Get
    End Property

    Public Property Picture() As Image
        Get
            Return fldImage
        End Get
        Set(ByVal Value As Image)
            fldImage = Value
        End Set
    End Property

#End Region 'Properties

#Region "Methods"

    Public Sub New(ByVal Name As String, ByVal IP As String)
        fldName = Name
        fldMode = ClientMode.Thumbnail
        StartTCPClient(IP, intPort)
    End Sub

    Public Function SendTypeByte(ByVal chrTypeChar As Char) As Boolean
        Dim TypeByte() As Byte = {AscW(chrTypeChar)}

        Try
            fldNetworkStream.Write(TypeByte, 0, TypeByte.Length)
        Catch ex As Exception
            Return False
            Exit Function
        End Try

        Return True
    End Function

    Public Function GetMessage(ByVal InputBuffer() As Byte) As ClientMessage
        Dim ClientMessage As ClientMessage

        Dim BufferStream As New System.IO.MemoryStream(InputBuffer)
        Dim MessageTypeByte As Byte = BufferStream.ReadByte
        ClientMessage.MessageType = ChrW(MessageTypeByte)

        Dim TitleLengthBytes() As Byte = {BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte}
        Dim TitleLength As Int64 = ByteArrayToInt64(TitleLengthBytes, False)

        Dim buffer(TitleLength - 1) As Byte
        BufferStream.Read(buffer, 0, TitleLength)
        ClientMessage.strTitle = System.Text.ASCIIEncoding.ASCII.GetString(buffer)

        Dim MessageLengthBytes() As Byte = {BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte, BufferStream.ReadByte}
        Dim MessageLength As Int64 = ByteArrayToInt64(MessageLengthBytes, False)
        ReDim buffer(MessageLength - 1)
        BufferStream.Read(buffer, 0, MessageLength)
        ClientMessage.strMessage = System.Text.ASCIIEncoding.ASCII.GetString(buffer)

        Return ClientMessage
    End Function

    Public Function GetDataType(ByRef buffer() As Byte) As Char
        Dim TypeByte As Byte = fldNetworkStream.ReadByte()

        Dim PayloadLengthBA() As Byte = {fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte, fldNetworkStream.ReadByte}
        Dim PayloadLength As Int64 = ByteArrayToInt64(PayloadLengthBA, False)
        fldTcpClient.ReceiveBufferSize = PayloadLength

        ReDim buffer(PayloadLength - 1)
        fldNetworkStream.Read(buffer, 0, fldTcpClient.ReceiveBufferSize)

        Return ChrW(TypeByte)
    End Function

    Public Function SendFullScreen() As Boolean
        Dim Image As Image = fldClientFuncs.PerformCapture

        Dim Imagebytes As Byte() = Me.ImageToByteArray(Image)
        Dim TypeByte() As Byte = {AscW("F")}
        Dim Size As Int64 = Imagebytes.Length
        Dim SizeBytes() As Byte = System.BitConverter.GetBytes(Size)

        Dim TempBytes As Byte() = Me.CombineByteArrays(TypeByte, SizeBytes)
        Dim sendbytes As Byte() = Me.CombineByteArrays(TempBytes, Imagebytes)

        Try
            fldNetworkStream.Write(sendbytes, 0, sendbytes.Length)
        Catch ex As Exception
            Return False
            Exit Function
        End Try

        fldFullTimeStamp = Now
        Return True

    End Function

    Public Function ReconnectAttempt(ByVal IP As String) As Boolean
        Try
            fldTcpClient = New System.Net.Sockets.TcpClient(IP, intPort)
            fldNetworkStream = fldTcpClient.GetStream

            Dim Size As Int64 = fldName.Length
            Dim SizeBytes() As Byte = System.BitConverter.GetBytes(Size)
            Dim NameBytes As [Byte]() = System.Text.Encoding.ASCII.GetBytes(fldName)
            Dim sendBytes As Byte() = Me.CombineByteArrays(SizeBytes, NameBytes)

            fldNetworkStream.Write(sendBytes, 0, sendBytes.Length)

            fldMode = ClientMode.Thumbnail

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function SendThumbnail() As Boolean
        Dim starttime As Integer = (System.DateTime.Now.Second * 1000) + System.DateTime.Now.Millisecond

        Dim Image As Image = fldClientFuncs.ThumbnailCapture(Screen.PrimaryScreen.Bounds.Width / 6, Screen.PrimaryScreen.Bounds.Height / 6)
        Dim Imagebytes As Byte() = Me.ImageToByteArray(Image)
        Dim TypeByte() As Byte = {AscW("T")}
        Dim Size As Int64 = Imagebytes.Length
        Dim SizeBytes() As Byte = System.BitConverter.GetBytes(Size)

        Dim TempBytes As Byte() = Me.CombineByteArrays(TypeByte, SizeBytes)
        Dim sendbytes As Byte() = Me.CombineByteArrays(TempBytes, Imagebytes)

        Try
            fldNetworkStream.Write(sendbytes, 0, sendbytes.Length)
        Catch ex As Exception
            Return False
            Exit Function
        End Try

        Dim endtime As Integer = (System.DateTime.Now.Second * 1000) + System.DateTime.Now.Millisecond

        fldImage = Image

        Return True

    End Function

#End Region     'Methods

#Region "Utilities"

    Private Sub StartTCPClient(ByVal IP As String, ByVal intPort As Integer)
        Dim blnConnected As Boolean = False
        While blnConnected = False
            Try
                fldTcpClient = New System.Net.Sockets.TcpClient(IP, intPort)
                fldNetworkStream = fldTcpClient.GetStream

                Dim Size As Int64 = fldName.Length
                Dim SizeBytes() As Byte = System.BitConverter.GetBytes(Size)
                Dim NameBytes As [Byte]() = System.Text.Encoding.ASCII.GetBytes(fldName)
                Dim sendBytes As Byte() = Me.CombineByteArrays(SizeBytes, NameBytes)

                fldNetworkStream.Write(sendBytes, 0, sendBytes.Length)
                blnConnected = True
            Catch ex As Exception
            End Try
            Application.DoEvents()
        End While
    End Sub

    Private Function ImageToByteArray(ByVal Image As Image) As Byte()
        Dim MemoryStream1 As New System.IO.MemoryStream
        Image.Save(MemoryStream1, System.Drawing.Imaging.ImageFormat.Jpeg)

        Dim MSArray As Byte()
        MSArray = MemoryStream1.ToArray()
        Return MSArray

    End Function

    Private Function CombineByteArrays(ByVal bytes1 As Byte(), ByVal bytes2 As Byte())
        Dim ReturnBytes(bytes1.Length + bytes2.Length - 1) As Byte
        Array.Copy(bytes1, ReturnBytes, bytes1.Length)
        Array.Copy(bytes2, 0, ReturnBytes, bytes1.Length, bytes2.Length)

        Return ReturnBytes
    End Function

    Private Function ByteArrayToInt64(ByVal BA() As Byte, ByVal SwitchEndian As Boolean) As Long
        Dim outVal As Long

        outVal = BitConverter.ToInt64(BA, 0)
        If SwitchEndian = True Then
            outVal = System.Net.IPAddress.NetworkToHostOrder(outVal)
        End If

        Return outVal
    End Function

#End Region     'Utilities

End Class
