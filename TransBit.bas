Attribute VB_Name = "TransBit"
Public Type BITMAPFILEHEADER    '14 bytes
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER   '40 bytes
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

Public Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Public Type BITMAPINFO_8
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 255) As RGBQUAD
End Type

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
  ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetDIBits_8 Lib "gdi32" Alias "SetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
  ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_8, _
  ByVal wUsage As Long) As Long

'Constants
Public Const DIB_RGB_COLORS As Long = 0

Public picWidth As Long, picHeight As Long, PicDC As Long, PicDC_M As Long


