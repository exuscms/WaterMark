Attribute VB_Name = "modRegionImage"
Option Explicit

Private Const RGN_OR As Long = 2
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'

Public Function RegionFromImage(fRgnForm As Object, picBackground As PictureBox, ByVal lBackColor As Long) As Long

    Dim hRegion         As Long
    Dim lpicHeight      As Long
    Dim lpicWidth       As Long
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lStart          As Long
    Dim hTempRegion     As Long
    Dim lResult         As Long
    
    hRegion = CreateRectRgn(0, 0, 0, 0)
    
    With picBackground
        lpicHeight = .Height / Screen.TwipsPerPixelY
        lpicWidth = .Width / Screen.TwipsPerPixelX
        
        For lRow = 0 To lpicHeight - 1
            lCol = 0
            
            Do While (lCol < lpicWidth)
                Do While (lCol < lpicWidth) And (GetPixel(.hdc, lCol, lRow) = lBackColor)
                    lCol = lCol + 1
                Loop
                
                If (lCol < lpicWidth) Then
                    lStart = lCol
                    
                    Do While (lCol < lpicWidth) And _
                             (GetPixel(.hdc, lCol, lRow) <> lBackColor)
                        lCol = lCol + 1
                    Loop
                    
                    If (lCol > lpicWidth) Then lCol = lpicWidth
                    
                    hTempRegion = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                    lResult = CombineRgn(hRegion, hRegion, hTempRegion, RGN_OR)
                    
                    DeleteObject hTempRegion
                End If
            Loop
        Next
    End With
    
    RegionFromImage = hRegion
    SetWindowRgn fRgnForm.hwnd, hRegion, True
End Function



