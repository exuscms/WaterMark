Attribute VB_Name = "WinDIctl"
'WINDI CONTROL MODULE
'(c) 2002-2003 By Marco Samy - marco_s2@hotmail.com
'for Lesson 2
Option Explicit
'API calls
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Need Type by BITMAPINFO
Public Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type
'This type is used in both 2 API function to handle image data
Public Type BITMAPINFO
  Header As BITMAPINFOHEADER
  Bits() As Byte
End Type
'Variables
Private SourceWidth As Long, DestWidth As Long
Private SourceHeight As Long, DestHeight As Long
Private SourceBuffer As BITMAPINFO
Private DestBuffer As BITMAPINFO
'GetDIs
'an easy way to use get dibits
'remeber you must use a PictureBox with no borders [BorderStyle = 0-None]
Public Sub GetDIs(ByRef Source As PictureBox, Bytes() As Byte, Optional ByRef ByteWid As Long, Optional ByRef ByteHeigh As Long)
  'Store a few attributes of the pictures for increased speed
  SourceWidth = Source.Width
  SourceHeight = Source.Height
  'Allocate array for source picture's bits
  ReDim SourceBuffer.Bits(3, SourceWidth - 1, SourceHeight - 1)
  With SourceBuffer.Header
    .biSize = 40
    .biWidth = SourceWidth
    .biHeight = -SourceHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * SourceWidth * SourceHeight
  End With
  'Get source pictures bits
  GetDIBits Source.hdc, Source.Image.Handle, 0, SourceHeight, SourceBuffer.Bits(0, 0, 0), SourceBuffer, 0&
'Putting Bytes inside the buffer variable
Bytes() = SourceBuffer.Bits()
ByteWid = SourceWidth - 1
ByteHeigh = SourceHeight - 1
End Sub
'SetDIs
'an easy way to use set dibits
'remeber you must use a PictureBox with no borders [BorderStyle = 0-None]
Public Sub SetDIs(Dest As PictureBox, Bytes() As Byte)
  DestWidth = Dest.Width
  DestHeight = Dest.Height
'Load Into Memory
ReDim DestBuffer.Bits(3, DestWidth - 1, DestHeight - 1)
'Putting Bytes inside the buffer variable
DestBuffer.Bits() = Bytes()
'Settting Buffer Header Information
  With DestBuffer.Header
    .biSize = 40
    .biWidth = DestWidth
    .biHeight = -DestHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * DestWidth * DestHeight
  End With
  'Load destination bits to destination picture
SetDIBits Dest.hdc, Dest.Image.Handle, 0, DestHeight, DestBuffer.Bits(0, 0, 0), DestBuffer, 0&
End Sub
Public Function GetBITMAPINFO(ByRef Source As PictureBox, ByRef iBITMAP As BITMAPINFO)
  'Store a few attributes of the pictures for increased speed
  SourceWidth = Source.Width
  SourceHeight = Source.Height
  'Allocate array for source picture's bits
  ReDim iBITMAP.Bits(3, SourceWidth - 1, SourceHeight - 1)
  With iBITMAP.Header
    .biSize = 40
    .biWidth = SourceWidth
    .biHeight = -SourceHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * SourceWidth * SourceHeight
  End With
  'Get source pictures bits
  GetDIBits Source.hdc, Source.Image.Handle, 0, SourceHeight, iBITMAP.Bits(0, 0, 0), iBITMAP, 0&
'Putting Bytes inside the buffer variable
End Function
'SetDIs
'an easy way to use set dibits
'remeber you must use a PictureBox with no borders [BorderStyle = 0-None]
Public Function SetBITMAPINFO(Dest As PictureBox, ByRef iBITMAP As BITMAPINFO)
  DestWidth = Dest.Width
  DestHeight = Dest.Height
'Settting Buffer Header Information
  With iBITMAP.Header
    .biSize = 40
    .biWidth = DestWidth
    .biHeight = -DestHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * DestWidth * DestHeight
  End With
  'Load destination bits to destination picture
SetDIBits Dest.hdc, Dest.Image.Handle, 0, DestHeight, iBITMAP.Bits(0, 0, 0), iBITMAP, 0&
End Function
