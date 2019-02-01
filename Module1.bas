Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Public Const DIB_PAL_COLORS = 1
Public Const DIB_PAL_INDICES = 2
Public Const DIB_PAL_LOGINDICES = 4
Public Const DIB_PAL_PHYSINDICES = 2
Public Const DIB_RGB_COLORS = 0
Public Const SRCCOPY = &HCC0020
Public Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

Public Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Bits() As Byte             '(Colors)
End Type

Sub RotBlt(Destdc As Long, Angle As Currency, x&, y&, W&, H&, ImgHandle&, Optional TranspColor&, Optional Alpha As Currency = 1, Optional pScale As Currency = 1, Optional px% = -32767, Optional py% = -32767)
  'Angle given is in rads
  Dim P() As Byte
  Dim ProcessedBits() As Byte
  Dim dx As Currency, dy As Currency, tx As Currency, ty As Currency
  Dim ix As Integer, iy As Integer
  Dim Tmp&, CX&, CY&, XX&, YY&
  Dim TR As Byte, TB As Byte, TG As Byte
  Dim D() As Byte
  Dim BackDC As Long
  Dim BackBmp As BITMAPINFO
  Dim iBitmap As Long
  Dim TopL As Currency, TopR As Currency, BotL As Currency, BotR As Currency
  Dim TopLV As Currency, TopRV As Currency, BotLV As Currency, BotRV As Currency
  Dim pSin As Currency, PCos As Currency
  Dim PicBmp As BITMAPINFO, PicDC As Long
  'Get the maximum width and heigth any rotation can produce
  If px = -32767 Then
     Tmp = Int(Sqr(W * W + H * H)) * pScale
  Else
     If py = -32767 Then py = (H / 2)  'pivot y
     Tmp = Int(Sqr(W * W + H * H) + Sqr((px - W / 2) ^ 2 + (py - H / 2) ^ 2)) * pScale
  End If
    'Set the rotation axis default values
  If px = -32767 Then px = (W / 2)  'pivot x
  If py = -32767 Then py = (H / 2)  'pivot y
  
  'Prepare the pixel arrays
  ReDim D(3, Tmp - 1, Tmp - 1) 'Holds The background image
  ReDim P(3, W - 1, H - 1)     'Holds the source image
  ReDim ProcessedBits(3, Tmp - 1, Tmp - 1) 'Holds the rotated result
   
  '[Create a Context - Copy the Backgroung - Get Background pixels]
  With BackBmp.Header
      .biBitCount = 4 * 8
      .biPlanes = 1
      .biSize = 40
      .biWidth = Tmp
      .biHeight = -Tmp
  End With

  'Create a context
  BackDC = CreateCompatibleDC(0)
  'Create a blank picture on the BackBmp standards (W,H,bitdebth)
  iBitmap = CreateDIBSection(BackDC, BackBmp, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
  'Copy the picture in to the context to make the context useable just like a picturebox
  SelectObject BackDC, iBitmap
  'Copy the background to the context
  BitBlt BackDC, 0, 0, Tmp, Tmp, Destdc, x - Tmp / 2, y - Tmp / 2, SRCCOPY
  'Analyze the background pixels and save to the array D(ColorIndex,X,Y)
  GetDIBits BackDC, iBitmap, 0, Tmp, D(0, 0, 0), BackBmp, DIB_RGB_COLORS
  
  '[Get SourceImage Pixels]
  With PicBmp.Header
      .biBitCount = 4 * 8
      .biPlanes = 1
      .biSize = 40
      .biWidth = W
      .biHeight = -H
  End With
  'Create a context
  PicDC = CreateCompatibleDC(0)
  'Copy the sourceimage in to the context to make the context useable (no need to create a new dibsection since the SourcePicture image is compatible)
  SelectObject PicDC, ImgHandle
  'Analize the sourceimage pixels and save to the array D
  GetDIBits PicDC, ImgHandle, 0, H, P(0, 0, 0), PicBmp, DIB_RGB_COLORS
  
  'Get the min values to scan
  CX = Int((Tmp - W) / 2)
  CY = Int((Tmp - H) / 2)
  
  'Convert to R,G,B format the transparent color
  TR = TranspColor And &HFF&
  TG = (TranspColor And &HFF00&) / &H100&
  TB = (TranspColor And &HFF0000) / &H10000
  
  'Precalculate the trigonometry
  PCos = Cos(Angle) / pScale
  pSin = Sin(Angle) / pScale

  'Loop through all pixels of the source image
  For XX = -CX To Tmp - CX - 1
   For YY = -CY To Tmp - CY - 1
      'Get the rotation translation (gives the SourceImage coordinate for each DestImage x,y)
      tx = (XX - px) * PCos - (YY - py) * pSin + px
      ty = (XX - px) * pSin + (YY - py) * PCos + py
      
      'Get nearest to the left pixel
      ix = Int(tx)
      iy = Int(ty)
      
      'Get the digits after the decimal point
      dx = Abs(tx - ix)
      dy = Abs(ty - iy)
      
      'Color the destination with the background color
      ProcessedBits(0, XX + CX, YY + CY) = D(0, XX + CX, YY + CY)
      ProcessedBits(1, XX + CX, YY + CY) = D(1, XX + CX, YY + CY)
      ProcessedBits(2, XX + CX, YY + CY) = D(2, XX + CX, YY + CY)


      If tx >= 0 And ix + 1 < W Then
       If ty >= 0 And iy + 1 < H Then

           'These variables hold Alpha value if the source pixel is non-transparent
           'If it's transparent they hold zero
           TopLV = -CBool(P(0, ix, iy) <> TR Or P(1, ix, iy) <> TG Or P(2, ix, iy) <> TB) * Alpha
           TopRV = -CBool(P(0, ix + 1, iy) <> TR Or P(1, ix + 1, iy) <> TG Or P(2, ix + 1, iy) <> TB) * Alpha
           BotLV = -CBool(P(0, ix, iy + 1) <> TR Or P(1, ix, iy + 1) <> TG Or P(2, ix, iy + 1) <> TB) * Alpha
           BotRV = -CBool(P(0, ix + 1, iy + 1) <> TR Or P(1, ix + 1, iy + 1) <> TG Or P(2, ix + 1, iy + 1) <> TB) * Alpha
           
           'The SourcePixel color maybe a combination of upto four pixels as tx and ty are not integers
           'The intersepted (by the current calculated source pixel) area each pixel involved (see .doc for more info)
           TopL = (1 - dx) * (1 - dy)
           TopR = dx * (1 - dy)
           BotL = (1 - dx) * dy
           BotR = dx * dy
        
           'Simplified explanation of the routine combination:
           'Alphablending (alpha being a real value from 0 to 1): DestColor = SourceImageColor * Alpha + BackImageColor * (1-Alpha)
           'Antialiasing: DestColor = SourceTopLeftPixel * TopLeftAreaIntersectedBySourcePixel +SourceTopRightPixel * TopRightAreaIntersectedBySourcePixel + bottomleft... + bottomrigth...
           
           'The AntiAliased Alpha assigment of colors
           ProcessedBits(0, XX + CX, YY + CY) = (P(0, ix, iy) * TopLV + D(0, XX + CX, YY + CY) * (1 - TopLV)) * TopL + (P(0, ix + 1, iy) * TopRV + D(0, XX + CX, YY + CY) * (1 - TopRV)) * TopR + (P(0, ix, iy + 1) * BotLV + D(0, XX + CX, YY + CY) * (1 - BotLV)) * BotL + (P(0, ix + 1, iy + 1) * BotRV + D(0, XX + CX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(1, XX + CX, YY + CY) = (P(1, ix, iy) * TopLV + D(1, XX + CX, YY + CY) * (1 - TopLV)) * TopL + (P(1, ix + 1, iy) * TopRV + D(1, XX + CX, YY + CY) * (1 - TopRV)) * TopR + (P(1, ix, iy + 1) * BotLV + D(1, XX + CX, YY + CY) * (1 - BotLV)) * BotL + (P(1, ix + 1, iy + 1) * BotRV + D(1, XX + CX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(2, XX + CX, YY + CY) = (P(2, ix, iy) * TopLV + D(2, XX + CX, YY + CY) * (1 - TopLV)) * TopL + (P(2, ix + 1, iy) * TopRV + D(2, XX + CX, YY + CY) * (1 - TopRV)) * TopR + (P(2, ix, iy + 1) * BotLV + D(2, XX + CX, YY + CY) * (1 - BotLV)) * BotL + (P(2, ix + 1, iy + 1) * BotRV + D(2, XX + CX, YY + CY) * (1 - BotRV)) * BotR
       End If
      End If
   Next
  Next
  
  'Draw the pixel array
  StretchDIBits Destdc, x - Tmp / 2, y - Tmp / 2, Tmp, Tmp, 0, 0, Tmp, Tmp, ProcessedBits(0, 0, 0), BackBmp, DIB_RGB_COLORS, SRCCOPY
  'Clear the variables
  Erase D
  Erase ProcessedBits
  Erase P
  DeleteObject iBitmap
  DeleteDC PicDC
  DeleteDC BackDC
End Sub


