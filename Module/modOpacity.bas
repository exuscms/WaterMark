Attribute VB_Name = "modOpacity"
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal blend As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Dim BF As BLENDFUNCTION
Dim lBF As Long
Private Const AC_SRC_OVER As Byte = &H0

Public Sub DoAlphablend(SrcPicBox As PictureBox, DestPicBox As PictureBox, AlphaVal As Integer)
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = AlphaVal
        .AlphaFormat = 0
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend DestPicBox.hdc, 0, 0, DestPicBox.ScaleWidth, DestPicBox.ScaleHeight, SrcPicBox.hdc, 0, 0, SrcPicBox.ScaleWidth, SrcPicBox.ScaleHeight, lBF
End Sub
