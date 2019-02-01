Attribute VB_Name = "ModRotate"
Type BITMAPINFOHEADER '40 bytes
BmSize As Long
BmWidth As Long
BmHeight As Long
BmPlanes As Integer
BmBitCount As Integer
BmCompression As Long
BmSizeImage As Long
BmXPelsPerMeter As Long
BmYPelsPerMeter As Long
BmClrUsed As Long
BmClrImportant As Long
End Type
Type BITMAPINFO
BmHeader As BITMAPINFOHEADER
End Type
'VB 16
Declare Sub GetDIBits Lib "GDI" (ByVal hDC%, ByVal hBitmap%, ByVal nStartScan%, ByVal nNumScans%, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage%)
Declare Sub SetDIBits Lib "GDI" (ByVal hDC%, ByVal hBitmap%, ByVal nStartScan%, ByVal nNumScans%, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage%)
'VB 32


Sub RotatePicDI(SrcPic As PictureBox, DestPic As PictureBox, A As Double)
Dim SrcInfo As BITMAPINFO, DesInfo As BITMAPINFO
Dim X&, Y&, CA As Double, SA As Double, nX&, nY&
Dim sW&, sH&, sW2&, sH2&, dW&, dH&, dW2&, dH2&
'RotatePicDI Picture1, Picture2, 45
Const Pi = 0.017453292519943
CA = Cos(A * Pi * -1): SA = Sin(A * Pi * -1)
sW = SrcPic.ScaleWidth
sH = SrcPic.ScaleHeight
dW = DestPic.ScaleWidth
dH = DestPic.ScaleHeight
sW2 = sW / 2: sH2 = sH / 2
dW2 = dW / 2: dH2 = dH / 2
SrcInfo.BmHeader.BmSize = 40 'Always
SrcInfo.BmHeader.BmWidth = sW 'Width
SrcInfo.BmHeader.BmHeight = -sH 'If You Want to Start Top-Botttom Put -Height
SrcInfo.BmHeader.BmPlanes = 1 'Always
SrcInfo.BmHeader.BmBitCount = 32 ' Can Be 16, 24,

SrcInfo.BmHeader.BmSizeImage = 3 * sW * sH
'If You Change The BitCount To 16 Or 24
'You Have To Change The SrcPix And DesPix Values
'Example: ReDim SrcPix("0,1,2,3,4" , sW - 1, sH - 1) As Long
'I Think For VB32 Users If You Have BitCount 32
'You Have To Change SrcPix And DesPix Values To 3,W,H
'(ReDim SrcPix(3, sW - 1, sH - 1) As Byte)
'Or (ReDim SrcPix(0 To 2, sW - 1, sH - 1) As Byte)
'This Should Get You The Red,Green,Blue Values
'2=Red,1=Green,0=Blue | 3=Red,2=Green,1=Blue
LSet DesInfo = SrcInfo 'Copy SrcInfo to DesInfo
DesInfo.BmHeader.BmWidth = dW 'Width
DesInfo.BmHeader.BmHeight = -dH 'If You Want to
DesInfo.BmHeader.BmSizeImage = 3 * dW * dH

ReDim SrcPix(0, sW - 1, sH - 1) As Long
ReDim DesPix(0, dW - 1, dH - 1) As Long
'Dont work try this
'ReDim SrcPix(0, sW - 1, sH - 1) As Byte
'ReDim DesPix(0, dW - 1, dH - 1) As Byte
'Or this
'ReDim SrcPix(0 To 2, sW - 1, sH - 1) As Byte
'ReDim DesPix(0 To 2, dW - 1, dH - 1) As Byte
'Also You Might Have To Change
'Pic.Image To Pic.Image.Handle
'Call GetDIBits(SrcPic.hDC, SrcPic.Image.Handle, 0&, sH, SrcPix(0, 0, 0), SrcInfo, 0&)
Call GetDIBits(SrcPic.hDC, SrcPic.Image, 0&, sH, SrcPix(0, 0, 0), SrcInfo, 0&)
For Y = 0 To dH - 1
For X = 0 To dW - 1
nX = CA * (X - dW2) - SA * (Y - dH2) + sW2
nY = SA * (X - dW2) + CA * (Y - dH2) + sH2
If nX > -1 And nY > -1 And nX < sW And nY < sH Then
DesPix(0, X, Y) = SrcPix(0, nX, nY)
'VB32 Might Have To Use This
'DesPix(1, X, Y) = SrcPix(1, nX, nY)
'DesPix(2, X, Y) = SrcPix(2, nX, nY)
End If
Next
Next
'Call SetDIBits(DestPic.hDC, DestPic.Image.Handle, 0&, dH, DesPix(0, 0, 0), DesInfo, 0&)
Call SetDIBits(DestPic.hDC, DestPic.Image, 0&, dH, DesPix(0, 0, 0), DesInfo, 0&)
DestPic.Picture = DestPic.Image
End Sub
