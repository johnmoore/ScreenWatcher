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

Imports System
Imports System.IO
Public Structure sctSection
    Dim sectionimg As Image
    Dim x As Integer
    Dim y As Integer
End Structure

Public Class ServerScreenFunctions
    Const SECWIDTH As Integer = 0
    Const SECHEIGHT As Integer = 1

    Const SCRNWIDTH As Integer = 0
    Const SCRNHEIGHT As Integer = 1

    Dim sectionSize(1) As Integer
    Dim screenSize(1) As Integer

    Dim sections(7, 5) As Image

    Private Sub Initialize(ByVal intSectionWidth As Integer, ByVal intSectionHeight As Integer)
        sectionSize(SECWIDTH) = intSectionWidth
        sectionSize(SECHEIGHT) = intSectionHeight
        screenSize(SCRNWIDTH) = intSectionWidth * 8
        screenSize(SCRNHEIGHT) = intSectionHeight * 6
    End Sub

    Public Function GetScreen() As Image
        Dim ScreenImg As Image = New Bitmap(screenSize(SCRNWIDTH), screenSize(SCRNHEIGHT))
        Dim gMain As Graphics = Graphics.FromImage(ScreenImg)
        Dim intScreenW, intScreenH As Integer
        For intScreenW = 1 To 8
            For intScreenH = 1 To 6
                Dim srcRect, destRect As Rectangle
                srcRect = New Rectangle(0, 0, sectionSize(SECWIDTH), sectionSize(SECHEIGHT))
                destRect = New Rectangle(((intScreenW - 1) * sectionSize(SECWIDTH)), ((intScreenH - 1) * sectionSize(SECHEIGHT)), sectionSize(SECWIDTH), sectionSize(SECHEIGHT))
                gMain.DrawImage(sections(intScreenW - 1, intScreenH - 1), destRect, srcRect, GraphicsUnit.Pixel)
            Next
        Next
        Return ScreenImg
    End Function

    Public Sub UpdateSection(ByVal section As sctSection)
        Initialize(section.sectionimg.Width, section.sectionimg.Height)
        sections(section.x, section.y) = section.sectionimg
    End Sub

    Public Sub UpdateAllSections(ByVal colSections As Collection)
        For Each section As sctSection In colSections
            sections(section.x, section.y) = section.sectionimg
        Next
    End Sub
End Class

Public Class ClientScreenFunctions
    Declare Auto Function BitBlt Lib "GDI32.DLL" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Int32) As Boolean
    Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As IntPtr) As Integer

    Const SRCCOPY As Object = &HCC0020
    Const SECWIDTH As Integer = 0
    Const SECHEIGHT As Integer = 1

    Dim dc1 As New IntPtr(CreateDC("DISPLAY", Nothing, Nothing, CType(Nothing, IntPtr)))
    Dim g1 As Graphics = Graphics.FromHdc(dc1)
    Dim screenshot As Image = New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, g1)
    Dim g2 As Graphics = Graphics.FromImage(screenshot)
    Dim sectionSize(1) As Integer
    Dim secOld(7, 5), secNew(7, 5) As Image

    Private Function PerformCapture() As Image
        dc1 = g1.GetHdc()
        Dim dc2 As IntPtr = g2.GetHdc()
        BitBlt(dc2, 0, 0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, dc1, 0, 0, SRCCOPY)
        g1.ReleaseHdc(dc1)
        g2.ReleaseHdc(dc2)
        Return screenshot
    End Function

    Private Function CaptureRect(ByVal sx As Integer, ByVal sy As Integer, ByVal ex As Integer, ByVal ey As Integer) As Image
        Dim screenshot As Image = New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, g1)
        Dim g2 As Graphics = Graphics.FromImage(screenshot)
        dc1 = g1.GetHdc()
        Dim dc2 As IntPtr = g2.GetHdc()
        BitBlt(dc2, 0, 0, ex - sx, ey - sy, dc1, sx, sy, SRCCOPY)
        g1.ReleaseHdc(dc1)
        g2.ReleaseHdc(dc2)
        Return screenshot
    End Function

    Private Function ImageRect(ByVal imgMain As Image, ByVal sx As Integer, ByVal sy As Integer, ByVal ex As Integer, ByVal ey As Integer) As Image
        Dim imgRect As New Bitmap(ex - sx, ey - sy)
        Dim gMain As Graphics = Graphics.FromImage(imgRect)
        gMain.DrawImage(imgMain, New Rectangle(0, 0, ex - sx, ey - sy), New Rectangle(sx, sy, ex - sx, ey - sy), GraphicsUnit.Pixel)
        Return imgRect
    End Function

    Private Function MergeImages(ByVal srcImgMain As Image, ByVal srcImgSub As Image, ByVal intX As Integer, ByVal intY As Integer) As Image
        Dim merged As New Bitmap(srcImgMain.Width, srcImgMain.Width)
        Dim gMain As Graphics = Graphics.FromImage(merged)
        gMain.DrawImage(srcImgMain, 0, 0)
        gMain.DrawImage(srcImgSub, intX, intY)
        Return merged
    End Function

    Public Function ThumbnailCapture(ByVal intWidth As Integer, ByVal intHeight As Integer) As Image
        Dim screenshot As Image = PerformCapture()
        Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
        dummyCallBack = New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        screenshot = screenshot.GetThumbnailImage(intWidth, intHeight, dummyCallBack, System.IntPtr.Zero)
        Return screenshot
    End Function

    Private Function ThumbnailCallback() As Boolean
        Return True
    End Function

    Private Function LoadSections(ByVal screenshot As Image) As Image(,)
        Dim LoadedSections(7, 5) As Image
        Dim intCountW, intCountH As Integer
        For intCountW = 1 To 8
            For intCountH = 1 To 6
                LoadedSections(intCountW - 1, intCountH - 1) = ImageRect(screenshot, ((intCountW - 1) * sectionSize(SECWIDTH)), ((intCountH - 1) * sectionSize(SECHEIGHT)), (intCountW * sectionSize(SECWIDTH)), (intCountH * sectionSize(SECHEIGHT)))
            Next
        Next
        Return LoadedSections
    End Function

    Public Function GetChangedSections() As Collection
        Dim ChangedSections As New Collection
        Dim sctMain As sctSection
        secNew = LoadSections(PerformCapture)
        Dim intCountW, intCountH As Integer
        For intCountW = 0 To 7
            For intCountH = 0 To 5
                If ImageToString(secOld(intCountW, intCountH)) <> ImageToString(secNew(intCountW, intCountH)) Then
                    sctMain.sectionimg = secNew(intCountW, intCountH)
                    sctMain.x = intCountW
                    sctMain.y = intCountH
                    ChangedSections.Add(sctMain)
                End If
            Next
        Next
        secOld = secNew.Clone
        Return ChangedSections
    End Function

    Private Function ImageToString(ByVal img As Image) As String
        Dim ms As New MemoryStream
        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        Dim msa As Byte()
        msa = ms.ToArray
        Return System.Text.ASCIIEncoding.ASCII.GetString(msa)
    End Function

    Public Function GetAllSections(Optional ByVal newCapture As Boolean = False) As Collection
        If newCapture = False Then
            Return ImgToSection(secOld)
        Else
            Return ImgToSection(LoadSections(PerformCapture()))
        End If
    End Function

    Private Function ImgToSection(ByVal sections(,) As Image) As Collection
        Dim colSections As New Collection
        Dim sctMain As sctSection
        Dim intCountW, intCountH As Integer
        For intCountW = 0 To 7
            For intCountH = 0 To 5
                sctMain.sectionimg = sections(intCountW, intCountH)
                sctMain.x = intCountW
                sctMain.y = intCountH
                colSections.Add(sctMain)
            Next
        Next
        Return colSections
    End Function

    Public Sub New()
        sectionSize(SECWIDTH) = Screen.PrimaryScreen.Bounds.Width / 8
        sectionSize(SECHEIGHT) = Screen.PrimaryScreen.Bounds.Height / 6
        secOld = LoadSections(PerformCapture)
    End Sub

End Class