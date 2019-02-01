Attribute VB_Name = "txt3d"
'3D-fonts by Stephan Swertvaegher
'<><><><><><><><><><><><><><><><><><><><><><><><><><>
'please make suggestions:

'stephan.swertvaegher@planetinternet.be
'                    or
'gumming@compaqnet.be
'--------------------------------------------------------------------------------------------
'some global stuff
Public TMidX%, TMidY%, qq%, xx%, yy As Single, Color&
'getpixel and setpixel is a lot faster
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Declarations for the EnumFontProc-function
Public Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const LF_FACESIZE = 32
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

'-------------------------------------------------------------------
'Function to read fonts very quickly
'-------------------------------------------------------------------
Public Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As LOGFONT, FontName As String, ZeroPos As Long
    CopyMemory LF, ByVal lplf, LenB(LF)
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, Chr$(0))
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    'store them in a combo
    frmtext.Combo1.AddItem FontName
    EnumFontProc = 1
End Function
'-------------------------------------------------------------------
'Sub to make 3D out of a font (.ttf for best results)
'
'The sub is easy to understand and is a bit "straight forward"
'There's no return
'Syntax:
'Print3D Form, Destination, Distance, R1, G1, B1, R2, G2, B2, ShiftX, ShiftY, Text, Check, Texture - Source
'Form:  the form where everything happens
'           is only used to get the textwidth of the text, and has no other purpose...
'Destination: picturebox or form
'Distance: The width of the 3D (from 5 to 50 for best results)
'R1.......B2: The first and second color to make a gradient, all values form 0-255
'ShiftX, ShiftY: for the 3D effect; values from 1-10
'Text: The text to put in 3D
'Check: True or False
'       True: with texture
'        False: no texture
' Texture-Source: a picturebox with the texture in it (can be any size)
'Remarks:
'        The destination and texture-souce must be in scalemode 3 (pixels)
'        and autoredraw = true
'        The autosize of the texture-source must be set to true
'-------------------------------------------------------------------
Public Sub Print3D(Frm As Object, Ob As Object, Dist%, R1 As Single, G1 As Single, B1 As Single, R2 As Single, G2 As Single, B2 As Single, Pxx As Single, Pyy As Single, Txt$, Ch As Boolean, Optional Ob2 As Object)
On Error Resume Next
Dim Sr As Single, Sg As Single, Sb As Single
Ob.Cls
'do 3D
Pxx = Pxx / 10
Pyy = Pyy / 10
'make sure text is always centerred
TMidX = (Ob.Width / 2) - (Ob.TextWidth(Frm.Text1.Text) / 2)
TMidY = (Ob.Height / 2) - (Ob.TextHeight(Frm.Text1.Text) / 2)
TMidX = TMidX - ((Pxx * Dist) / 2)
TMidY = TMidY - ((Pyy * Dist) / 2)
Sr = (R2 - R1) / Dist
Sg = (G2 - G1) / Dist
Sb = (B2 - B1) / Dist
'print a lot of text
For xx = 0 To Dist - 1
    Ob.CurrentX = TMidX + (xx * Pxx)
    Ob.CurrentY = TMidY + (xx * Pyy)
    R1 = R1 + Sr
    G1 = G1 + Sg
    B1 = B1 + Sb
    'the values cannot be < 0
        If Int(R1) < 0 Then R1 = 0
        If Int(G1) < 0 Then G1 = 0
        If Int(B1) < 0 Then B1 = 0
    Ob.ForeColor = RGB(Int(R1), Int(G1), Int(B1))
    Ob.Print Txt
Next xx
'Now look if texture is wanted
If Ch = False Then Exit Sub
'Yep, do texture
Dim tx%, ty%
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
If GetPixel(Ob.hdc, xx, yy) = RGB(Int(R1), Int(G1), Int(B1)) Then
Color = GetPixel(Ob2.hdc, tx, ty)
SetPixel Ob.hdc, xx, yy, Color
End If
ty = ty + 1
If ty = Ob2.Height - 1 Then ty = 0
Next yy
ty = 0
tx = tx + 1
If tx = Ob2.Width - 1 Then tx = 0
Next xx
Ob.Refresh
End Sub
'-------------------------------------------------------------------
'That's all folks !


