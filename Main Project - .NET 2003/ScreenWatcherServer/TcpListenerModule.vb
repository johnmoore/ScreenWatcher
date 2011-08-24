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

Module TcpListenerModule

    Private tcplistener As New System.Net.Sockets.TcpListener(8000)

    Public Sub Initialize()
        tcplistener.Start()
    End Sub

    Public Function PendingClients() As Boolean
        Return tcplistener.Pending
    End Function

    Public Function ConnectWaitingClients() As Student()
        Dim StudentsColl As New Collection
        Dim ClientsWaiting As Boolean = tcplistener.Pending

        While ClientsWaiting = True
            Dim NewClient As System.Net.Sockets.TcpClient = tcplistener.AcceptTcpClient
            Dim NewNetworkStream As System.Net.Sockets.NetworkStream = NewClient.GetStream

            Dim PayloadLengthBA() As Byte = {NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte, NewNetworkStream.ReadByte}
            Dim PayloadLength As Int64 = ByteArrayToInt64(PayloadLengthBA, False)
            NewClient.ReceiveBufferSize = PayloadLength
            Dim bytes(NewClient.ReceiveBufferSize - 1) As Byte
            NewNetworkStream.Read(bytes, 0, CInt(NewClient.ReceiveBufferSize))
            Dim ClientName As String = System.Text.Encoding.ASCII.GetString(bytes) '.TrimEnd(ChrW(0))

            Dim NewStudent As New Student(ClientName, NewClient, NewNetworkStream)

            StudentsColl.Add(NewStudent)
            ClientsWaiting = tcplistener.Pending
        End While

        Dim Students(StudentsColl.Count - 1) As Student
        Dim index As Integer

        For Each Student As Student In StudentsColl
            Students(index) = Student
            index += 1
        Next

        Return Students

    End Function

    Public Function GetDataType(ByVal networkstream As System.Net.Sockets.NetworkStream) As String
        Dim TypeByte As Integer = networkstream.ReadByte()
        Return ChrW(TypeByte)
    End Function

    Public Function SendTypeByte(ByVal Client As Student, ByVal chrTypeChar As Char) As Boolean
        Dim typebyte() As Byte = {AscW(chrTypeChar)}
        Dim size As Int64 = typebyte.Length
        Dim sizebytes() As Byte = System.BitConverter.GetBytes(size)
        Dim TempBytes As Byte() = CombineByteArrays(typebyte, sizebytes)
        TempBytes = CombineByteArrays(TempBytes, typebyte)
        Try
            Client.NetworkStream.Write(TempBytes, 0, TempBytes.Length)
        Catch ex As Exception
            Return False
            Exit Function
        End Try
        Return True
    End Function

    Public Function SendMessage(ByVal client As Student, ByVal chrMsgType As Char, ByVal strTitle As String, ByVal strMessage As String) As Boolean
        Dim TypeByte() As Byte = {AscW("M")}
        Dim MsgTypeByte() As Byte = {AscW(chrMsgType)}

        Dim TitleBytes() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(strTitle)
        Dim MessageBytes() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(strMessage)
        Dim TitleLength As Int64 = TitleBytes.Length
        Dim MessageLength As Int64 = MessageBytes.Length
        Dim TitleSizeBytes() As Byte = System.BitConverter.GetBytes(TitleLength)
        Dim MessageSizeBytes() As Byte = System.BitConverter.GetBytes(MessageLength)
        Dim PayloadLength As Int64 = MsgTypeByte.Length + TitleSizeBytes.Length + TitleBytes.Length + MessageSizeBytes.Length + MessageBytes.Length
        Dim PayloadSizeBytes() As Byte = System.BitConverter.GetBytes(PayloadLength)

        Dim TempBytes As Byte() = CombineByteArrays(TypeByte, PayloadSizeBytes)
        TempBytes = CombineByteArrays(TempBytes, MsgTypeByte)
        TempBytes = CombineByteArrays(TempBytes, TitleSizeBytes)
        TempBytes = CombineByteArrays(TempBytes, TitleBytes)
        TempBytes = CombineByteArrays(TempBytes, MessageSizeBytes)
        TempBytes = CombineByteArrays(TempBytes, MessageBytes)

        Try
            client.NetworkStream.Write(TempBytes, 0, TempBytes.Length)
        Catch ex As Exception
            Return False
            Exit Function
        End Try
        Return True
    End Function

    Public Function GetStreamImage(ByVal client As Student) As Image
        On Error GoTo CorruptedPacket

        Application.DoEvents()
        Dim ImageBytes As Byte()
        Dim PayloadLengthBA() As Byte = {client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte, client.NetworkStream.ReadByte}
        Dim PayloadLength As Int64 = ByteArrayToInt64(PayloadLengthBA, False)
        client.TcpClient.ReceiveBufferSize = PayloadLength
        Dim NumOfPackets As Integer = Math.Ceiling(PayloadLength / client.TcpClient.ReceiveBufferSize)

        Dim ByteCollection As New Collection
        Dim DisplayImage As Image
        For i As Integer = 1 To NumOfPackets
            Dim newbytes(client.TcpClient.ReceiveBufferSize - 1) As Byte
            client.NetworkStream.Read(newbytes, 0, client.TcpClient.ReceiveBufferSize)
            ByteCollection.Add(newbytes)
        Next

        ImageBytes = CombineByteArrays(ByteCollection)
        DisplayImage = ConvertByteToImage(ImageBytes)

        client.ThumbTimestamp = Now
        Return DisplayImage

        Exit Function

CorruptedPacket:
        Dim buffer(client.TcpClient.ReceiveBufferSize - 1) As Byte
        Do Until client.NetworkStream.DataAvailable = False
            client.NetworkStream.Read(buffer, 0, client.TcpClient.ReceiveBufferSize)
        Loop
    End Function

#Region "Utilities"

    Private Function ByteArrayToInt64(ByVal BA() As Byte, ByVal SwitchEndian As Boolean) As Long
        Dim outVal As Long

        outVal = BitConverter.ToInt64(BA, 0)
        If SwitchEndian = True Then
            outVal = System.Net.IPAddress.NetworkToHostOrder(outVal)
        End If

        Return outVal
    End Function

    Private Function CombineByteArrays(ByVal ByteCollection As Collection) As Byte()
        If ByteCollection.Count > 1 Then
            Dim ReturnBytes As Byte() = CombineByteArrays(ByteCollection(1), ByteCollection(2))
            For i As Integer = 3 To ByteCollection.Count
                ReturnBytes = CombineByteArrays(ReturnBytes, ByteCollection(i))
            Next

            Return ReturnBytes
        Else
            Return ByteCollection(1)
        End If
    End Function

    Private Function CombineByteArrays(ByVal bytes1 As Byte(), ByVal bytes2 As Byte()) As Byte()
        Dim ReturnBytes(bytes1.Length + bytes2.Length - 1) As Byte
        Array.Copy(bytes1, ReturnBytes, bytes1.Length)
        Array.Copy(bytes2, 0, ReturnBytes, bytes1.Length, bytes2.Length)

        Return ReturnBytes
    End Function

    Private Overloads Function ConvertByteToImage(ByVal ImageArray As Byte()) As System.Drawing.Image

        Dim objMS As New System.IO.MemoryStream
        objMS.Write(ImageArray, 0, ImageArray.Length)

        Try
            ConvertByteToImage = System.Drawing.Image.FromStream(objMS)
        Catch ae As System.ArgumentException
            ConvertByteToImage = Nothing
        End Try

    End Function

    Private Function PrintArray(ByVal array As Byte()) As String
        Dim strString As String
        Dim Percent As Integer
        Dim Fraction As Double
        Dim LastPercent As Integer

        For i As Integer = 0 To array.Length - 1
            strString = strString & array(i).ToString & ", "
            Fraction = i / (array.Length - 1)
            Percent = Math.Floor(Fraction * 100)
            If Percent > 0 And Percent Mod 10 = 0 And Percent <> LastPercent Then
                LastPercent = Percent
            End If
            Application.DoEvents()
        Next

        Return strString
    End Function

#End Region 'Utilites

End Module