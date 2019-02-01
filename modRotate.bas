Attribute VB_Name = "modRotate"
'* CODED BY: BattleStorm
'* EMAIL: battlestorm@cox.net
'* UPDATED: 08/02/2002
'* PURPOSE: Rotates a picture from one
'*     picturebox to another using an
'*     angle specified in degrees from
'*     -359.999° to 359.999°.
'* COPYRIGHT: This program and source
'*     code is freeware and can be copied
'*     and/or distributed as long as you
'*     mention the original author. I am
'*     not responsible for any harm as the
'*     outcome of using any of this code.

'* Please note that you can enter angles less
'* than -359.999° and greater than 359.999°, but
'* it would serve no purpose. Code will not
'* return an error if valid ranges are not used.
'* It will simply over rotate the picture and the
'* result would be the same.

'* Program uses 3 types of pixel setting:
'*     1. Point and Pset - Slow (Average 600 ms)
'*     2. GetPixel and SetPixel - Medium (Average 500 ms)
'*     3. GetDiBits and SetDiBits - Very, Very Fast (Average 35 ms)

Option Explicit

'API calls
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'Types
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
  Header As BITMAPINFOHEADER
  Bits() As Byte
End Type

'Variables
Private biRct As RECT
Private x As Long, y As Long
Private biW As Long, biH As Long
Private cx As Double, cy As Double
Private dx As Double, dy As Double
Private cosa As Double, sina As Double
Private rx As Double, ry As Double
Private irx As Integer, iry As Integer
Private drx As Double, dry As Double
Private xin As Double, yin As Double
Private pcol As Long
Private r1 As Integer, g1 As Integer, b1 As Integer
Private r2 As Integer, g2 As Integer, b2 As Integer
Private r3 As Integer, g3 As Integer, b3 As Integer
Private r4 As Integer, g4 As Integer, b4 As Integer
Private ir1 As Integer, ig1 As Integer, ib1 As Integer
Private ir2 As Integer, ig2 As Integer, ib2 As Integer
Private R As Integer, G As Integer, B As Integer
Private SourceWidth As Long, DestWidth As Long
Private SourceHeight As Long, DestHeight As Long
Private SourceBuffer As BITMAPINFO
Private DestBuffer As BITMAPINFO

'Constants
Private Const Deg2Rad As Double = 0.017453292519943

'Rotate from source to destination in angle of degrees
'with or without Anti-Alias using Point and Pset
Public Sub PointRotate(ByRef Source As PictureBox, ByRef Dest As PictureBox, ByVal Angle As Double, ByVal AntiAlias As Boolean)
  'Store a few attributes of the pictures for increased speed
  SourceWidth = Source.Width
  SourceHeight = Source.Height
  DestWidth = Dest.Width
  DestHeight = Dest.Height

  'Get center of source picture
  cx = SourceWidth * 0.5
  cy = SourceHeight * 0.5
  
  'Get center of destination picture
  dx = DestWidth * 0.5
  dy = DestHeight * 0.5
  
  'Convert angle to Sin/Cos radians
  cosa = Cos(Angle * Deg2Rad * -1)
  sina = Sin(Angle * Deg2Rad * -1)
  
  'Get bounds of source picture
  biW = SourceWidth - 1
  biH = SourceHeight - 1
  SetRect biRct, 0, 0, biW, biH

  'Clear dectination picture
  Dest.Cls
  For y = 0 To DestHeight - 1
    'Destination Y to calculate
    yin = y - dy
    For x = 0 To DestWidth - 1
      'Destination X to calculate
      xin = x - dx
      
      'Rotate destination X, Y according to angle in radians
      rx = xin * cosa - yin * sina + cx
      ry = xin * sina + yin * cosa + cy
      
      'Round of rotated pixels X, Y coordinates
      irx = Int(rx)
      iry = Int(ry)
      
      'If rotated pixel is within bounds of destination
      If (PtInRect(biRct, irx, iry)) Then
      
        'Convert pixel to destination
        drx = rx - irx
        dry = ry - iry
        
        'If Anti-Alias switch is on
        If AntiAlias Then
          'Get rotated pixel
          pcol = Source.Point(irx, iry)
          UnRGB pcol, r1, g1, b1
          
          'Get rotated pixels right neighbor
          pcol = Source.Point(irx + 1, iry)
          UnRGB pcol, r2, g2, b2
          
          'Get rotated pixels lower neighbor
          pcol = Source.Point(irx, iry + 1)
          UnRGB pcol, r3, g3, b3
          
          'Get rotated pixels lower right neighbor
          pcol = Source.Point(irx + 1, iry + 1)
          UnRGB pcol, r4, g4, b4
          
          'Interpolate pixels along Y axis
          ib1 = b1 * (1 - dry) + b3 * dry
          ig1 = g1 * (1 - dry) + g3 * dry
          ir1 = r1 * (1 - dry) + r3 * dry
          ib2 = b2 * (1 - dry) + b4 * dry
          ig2 = g2 * (1 - dry) + g4 * dry
          ir2 = r2 * (1 - dry) + r4 * dry
    
          'Interpolate pixels along X axis
          B = ib1 * (1 - drx) + ib2 * drx
          G = ig1 * (1 - drx) + ig2 * drx
          R = ir1 * (1 - drx) + ir2 * drx
          
          'Check for valid color range
          If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
          If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
          If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
          
          'Plot interpolated pixel to destination picture
          Dest.PSet (x, y), RGB(R, G, B)
        Else
          'Get rotated pixel
          pcol = Source.Point(irx, iry)
          
          'Plot rotated pixel to destination picture
          Dest.PSet (x, y), pcol
        End If
      End If
    Next x
  Next y
End Sub

'Rotate from source to destination in angle of degrees
'with or without Anti-Alias using GetPixel and SetPixel
Public Sub PixelRotate(ByRef Source As PictureBox, ByRef Dest As PictureBox, ByVal Angle As Double, ByVal AntiAlias As Boolean)
  'Store a few attributes of the pictures for increased speed
  SourceWidth = Source.Width
  SourceHeight = Source.Height
  DestWidth = Dest.Width
  DestHeight = Dest.Height

  'Get center of source picture
  cx = SourceWidth * 0.5
  cy = SourceHeight * 0.5
  
  'Get center of destination picture
  dx = DestWidth * 0.5
  dy = DestHeight * 0.5
  
  'Convert angle to Sin/Cos radians
  cosa = Cos(Angle * Deg2Rad * -1)
  sina = Sin(Angle * Deg2Rad * -1)
  
  'Get bounds of source picture
  biW = SourceWidth - 1
  biH = SourceHeight - 1
  SetRect biRct, 0, 0, biW, biH

  'Clear dectination picture
  Dest.Cls
  For y = 0 To DestHeight - 1
    'Destination Y to calculate
    yin = y - dy
    For x = 0 To DestWidth - 1
      'Destination X to calculate
      xin = x - dx
      
      'Rotate destination X, Y according to angle in radians
      rx = xin * cosa - yin * sina + cx
      ry = xin * sina + yin * cosa + cy
      
      'Round of rotated pixels X, Y coordinates
      irx = Int(rx)
      iry = Int(ry)
      
      'If rotated pixel is within bounds of destination
      If (PtInRect(biRct, irx, iry)) Then
      
        'Convert pixel to destination
        drx = rx - irx
        dry = ry - iry
        
        'If Anti-Alias switch is on
        If AntiAlias Then
          'Get rotated pixel
          pcol = GetPixel(Source.hdc, irx, iry)
          UnRGB pcol, r1, g1, b1
          
          'Get rotated pixels right neighbor
          pcol = GetPixel(Source.hdc, irx + 1, iry)
          UnRGB pcol, r2, g2, b2
          
          'Get rotated pixels lower neighbor
          pcol = GetPixel(Source.hdc, irx, iry + 1)
          UnRGB pcol, r3, g3, b3
          
          'Get rotated pixels lower right neighbor
          pcol = GetPixel(Source.hdc, irx + 1, iry + 1)
          UnRGB pcol, r4, g4, b4
          
          'Interpolate pixels along Y axis
          ib1 = b1 * (1 - dry) + b3 * dry
          ig1 = g1 * (1 - dry) + g3 * dry
          ir1 = r1 * (1 - dry) + r3 * dry
          ib2 = b2 * (1 - dry) + b4 * dry
          ig2 = g2 * (1 - dry) + g4 * dry
          ir2 = r2 * (1 - dry) + r4 * dry
    
          'Interpolate pixels along X axis
          B = ib1 * (1 - drx) + ib2 * drx
          G = ig1 * (1 - drx) + ig2 * drx
          R = ir1 * (1 - drx) + ir2 * drx
          
          'Check for valid color range
          If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
          If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
          If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
          
          'Plot interpolated pixel to destination picture
          SetPixel Dest.hdc, x, y, RGB(R, G, B)
        Else
          'Get rotated pixel
          pcol = GetPixel(Source.hdc, irx, iry)
          
          'Plot rotated pixel to destination picture
          SetPixel Dest.hdc, x, y, pcol
        End If
      End If
    Next x
  Next y
End Sub

'Rotate from source to destination in angle of degrees
'with or without Anti-Alias using GetDiBits and SetDiBits
Public Sub BitRotate(ByRef Source As PictureBox, ByRef Dest As PictureBox, ByVal Angle As Double, ByVal AntiAlias As Boolean)
  'Store a few attributes of the pictures for increased speed
  SourceWidth = Source.Width
  SourceHeight = Source.Height
  DestWidth = Dest.Width
  DestHeight = Dest.Height
  
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
  
  'Allocate array for dest picture's bits
  ReDim DestBuffer.Bits(3, DestWidth - 1, DestHeight - 1)
  With DestBuffer.Header
    .biSize = 40
    .biWidth = DestWidth
    .biHeight = -DestHeight
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * DestWidth * DestHeight
  End With
  

  'Get center of source picture
  cx = SourceWidth * 0.5
  cy = SourceHeight * 0.5
  
  'Get center of destination picture
  dx = DestWidth * 0.5
  dy = DestHeight * 0.5
  
  'Convert angle to Sin/Cos radians
  cosa = Cos(Angle * Deg2Rad * -1)
  sina = Sin(Angle * Deg2Rad * -1)
  
  'Get bounds of source picture
  biW = SourceWidth - 1
  biH = SourceHeight - 1
  SetRect biRct, 0, 0, biW, biH

  'Clear dectination picture
  Dest.Cls
  For y = 0 To DestHeight - 1
    'Destination Y to calculate
    yin = y - dy
    For x = 0 To DestWidth - 1
      'Destination X to calculate
      xin = x - dx
      
      'Rotate destination X, Y according to angle in radians
      rx = xin * cosa - yin * sina + cx
      ry = xin * sina + yin * cosa + cy
      
      'Round of rotated pixels X, Y coordinates
      irx = Int(rx)
      iry = Int(ry)
      
      'If rotated pixel is within bounds of destination
      If (PtInRect(biRct, irx, iry)) Then
      
        'Convert pixel to destination
        drx = rx - irx
        dry = ry - iry
        
        'If Anti-Alias switch is on
        If AntiAlias Then
          'Get rotated pixel
          r1 = SourceBuffer.Bits(2, irx, iry)
          g1 = SourceBuffer.Bits(1, irx, iry)
          b1 = SourceBuffer.Bits(0, irx, iry)
          
          'Get rotated pixels right neighbor
          r2 = SourceBuffer.Bits(2, irx + 1, iry)
          g2 = SourceBuffer.Bits(1, irx + 1, iry)
          b2 = SourceBuffer.Bits(0, irx + 1, iry)
          
          'Get rotated pixels lower neighbor
          r3 = SourceBuffer.Bits(2, irx, iry + 1)
          g3 = SourceBuffer.Bits(1, irx, iry + 1)
          b3 = SourceBuffer.Bits(0, irx, iry + 1)
          
          'Get rotated pixels lower right neighbor
          r4 = SourceBuffer.Bits(2, irx + 1, iry + 1)
          g4 = SourceBuffer.Bits(1, irx + 1, iry + 1)
          b4 = SourceBuffer.Bits(0, irx + 1, iry + 1)
          
          'Interpolate pixels along Y axis
          ib1 = b1 * (1 - dry) + b3 * dry
          ig1 = g1 * (1 - dry) + g3 * dry
          ir1 = r1 * (1 - dry) + r3 * dry
          ib2 = b2 * (1 - dry) + b4 * dry
          ig2 = g2 * (1 - dry) + g4 * dry
          ir2 = r2 * (1 - dry) + r4 * dry
    
          'Interpolate pixels along X axis
          B = ib1 * (1 - drx) + ib2 * drx
          G = ig1 * (1 - drx) + ig2 * drx
          R = ir1 * (1 - drx) + ir2 * drx
          
          'Check for valid color range
          If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
          If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
          If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
          
          'Plot interpolated pixel to destination picture
          DestBuffer.Bits(2, x, y) = R
          DestBuffer.Bits(1, x, y) = G
          DestBuffer.Bits(0, x, y) = B
        Else
          'Get rotated pixel
          R = SourceBuffer.Bits(2, irx, iry)
          G = SourceBuffer.Bits(1, irx, iry)
          B = SourceBuffer.Bits(0, irx, iry)
          
          'Plot rotated pixel to destination picture
          DestBuffer.Bits(2, x, y) = R
          DestBuffer.Bits(1, x, y) = G
          DestBuffer.Bits(0, x, y) = B
        End If
      End If
    Next x
  Next y
  'Load destination bits to destination picture
  SetDIBits Dest.hdc, Dest.Image.Handle, 0, DestHeight, DestBuffer.Bits(0, 0, 0), DestBuffer, 0&
End Sub

'Break apart RGB values
Private Sub UnRGB(RGBCol As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
  Red = RGBCol And 255
  Green = Int(RGBCol / 256) And 255
  Blue = Int(RGBCol / 65536) And 255
End Sub

