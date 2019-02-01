Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Long, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Global Const SRCCOPY = &HCC0020

Private Function ZHex(lHex As Long, iZeros As Integer) As String
  'Returns a HEX string of specified length (pad zeros on left)
  ZHex = Right$(String$(iZeros - 1, "0") & Hex$(lHex), iZeros)
End Function

Public Function MakeHexRGB(r As Long, G As Long, B As Long) As String
  'Returns hex value for rgb color values
  MakeHexRGB = ZHex(r, 2) & ZHex(G, 2) & ZHex(B, 2)
End Function

Public Function MakeHexLong(lngColor As Long) As String
  Dim r As Long, G As Long, B As Long
  r = RGBRed(lngColor)
  G = RGBGreen(lngColor)
  B = RGBBlue(lngColor)
  'Returns hex value for a long color value
  MakeHexLong = ZHex(r, 2) & ZHex(G, 2) & ZHex(B, 2)
End Function

Public Function RGBRed(RGBCol As Long) As Integer
  'Returns the Red component from an RGB Color
  RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
  'Returns the Green component from an RGB Color
  RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
  'Returns the Blue component from an RGB Color
  RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function


