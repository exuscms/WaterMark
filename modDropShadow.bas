Attribute VB_Name = "modDropShadow"
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub DrawDropShadow(objPic As PictureBox, sWidth As Long)
'Procedure develops a "drop shadow" around a picturebox (sWidth in Pixels).
'Righthand and lower edge sections of the object container are stored in temporary
'Device Contexts and used as "background(s)" to create a shadow effect. These
'sections are added to the existing picturebox to create an integral effect.
Dim lDC As Long, rDC As Long
Dim lBMP As Long, rBMP As Long
Dim X As Long, Y As Long, uColor As Long
    
'Limit the shadow offset due to errors in the algorithm
If sWidth > 14 Then sWidth = 14

With objPic
    .Visible = False
    'Capture the right and lower "shadow sections"
    '- create a Device Context to store the lower edge container section
    lDC = CreateCompatibleDC(.hDC)
    lBMP = CreateCompatibleBitmap(.hDC, .Width + sWidth, sWidth)
    SelectObject lDC, lBMP
    '- capture the lower edge "shadow" section
    BitBlt lDC, 0, 0, .Width + sWidth, sWidth, .Container.hDC, .Left, .Top + .Height, vbSrcCopy
        
    '- create a Device Context to store the right edge container section
    rDC = CreateCompatibleDC(.hDC)
    rBMP = CreateCompatibleBitmap(.hDC, sWidth, .Height + sWidth)
    SelectObject rDC, rBMP
    '- capture the right edge "shadow" section
    BitBlt rDC, 0, 0, sWidth, .Height + sWidth, .Container.hDC, .Left + .Width, .Top, vbSrcCopy
    
    '- create a Device Context to store the generated image
    genDC = CreateCompatibleDC(.hDC)
    genBMP = CreateCompatibleBitmap(.hDC, .Width + sWidth, .Height + sWidth)
    SelectObject genDC, genBMP
    '- capture the original image
    BitBlt genDC, 0, 0, .Width - sWidth, .Height - sWidth, .hDC, 0, 0, vbSrcCopy
    
    'Enlarge objPic
    .Width = .Width + sWidth
    .Height = .Height + sWidth
    
    'Note: The algorithm used to produce the gradient shadow effect was
    'developed using a concept from the old vbSmart.com website. I have never
    'had the time to optimize it and do not like the distortions that occur
    'with larger (>12 pixel) shadow widths. Please post any improvements!
    
    ' - Simulate a shadow on right edge...
    For X = 1 To sWidth
        For Y = 0 To 3
            uColor = GetPixel(rDC, sWidth - X, Y)
            SetPixel .hDC, .Width - X, Y, uColor
        Next Y
        For Y = 4 To 7
            uColor = GetPixel(rDC, sWidth - X, Y)
            If X + Y <= .Height Then
                SetPixel .hDC, .Width - X, Y, uMask(3 * X * (Y - 3), uColor)
            End If
        Next Y
        For Y = 8 To .Height - 5
            uColor = GetPixel(rDC, sWidth - X, Y)
            If X + Y <= .Height Then
                SetPixel .hDC, .Width - X, Y, uMask(15 * X, uColor)
            End If
        Next Y
        For Y = .Height - 5 To .Height - 1
            uColor = GetPixel(rDC, sWidth - X, Y)
            If X + Y <= .Height + 3 Then
                SetPixel .hDC, .Width - X, Y, uMask(-3 * X * (Y - .Height), uColor)
            End If
        Next Y
    Next X
    
    ' - Simulate a shadow on the bottom edge...
    For Y = 1 To sWidth
        For X = 0 To 3
            uColor = GetPixel(lDC, X, sWidth - Y)
            SetPixel .hDC, X, .Height - Y, uColor
        Next X
        For X = 4 To 7
            uColor = GetPixel(lDC, X, sWidth - Y)
            SetPixel .hDC, X, .Height - Y, uMask(3 * (X - 3) * Y, uColor)
        Next X
        For X = 8 To .Width - 5
            uColor = GetPixel(lDC, X, sWidth - Y)
            If X + Y <= .Width Then
                SetPixel .hDC, X, .Height - Y, uMask(15 * Y, uColor)
            End If
        Next X
    Next Y
    .Visible = True
End With

ExitSub:
' - Release the resources
    DeleteDC lDC
    DeleteObject lBMP
    DeleteDC rDC
    DeleteObject rBMP
    
End Sub

Private Function uMask(lScale As Long, lColor As Long) As Long
'Function splits a color into its RGB components and transforms the
'color using a psuedo scale 0..255
Dim R As Long, G As Long, B As Long
Dim sColor As String
    
    sColor = CStr(Format(Hex(lColor), "000000"))
    ' Extract the component values
    R = CLng("&H" & Right(sColor, 2))
    G = CLng("&H" & Mid(sColor, 3, 2))
    B = CLng("&H" & Left(sColor, 2))
    
    'Create Fade effect
    R = R - R * lScale / 255  'pTransform(lScale, R)
    If R < 0 Then R = 0
    G = G - G * lScale / 255 'pTransform(lScale, G)
    If G < 0 Then G = 0
    B = B - B * lScale / 255 'pTransform(lScale, B)
    If B < 0 Then B = 0
    
    uMask = RGB(R, G, B)
    
End Function



