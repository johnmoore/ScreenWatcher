Public Class ClassStudent

#Region "Fields"
    Private fldFullName As String
    Private fldUserName As String
    Private fldMachineName As String
    Private fldNetworkName As String
    'Private fldComputerName As String

    Private fldImage As Image
    Private fldNetworkStream As System.Net.Sockets.NetworkStream
    Private fldTcpClient As System.Net.Sockets.TcpClient
    Private fldThumbTimestamp As DateTime

#End Region 'Fields

#Region "Properties"

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

    'Public ReadOnly Property ComputerName() As String
    '    Get
    '        Return fldComputerName
    '    End Get
    'End Property

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
    End Sub

#End Region     'Methods

End Class
