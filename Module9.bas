Attribute VB_Name = "RegionImage"
Option Explicit

Private Const RGN_OR                        As Long = 2

Private Declare Function SetWindowRgn Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal hRgn As Long, _
     ByVal bRedraw As Boolean) _
    As Long
    
Private Declare Function CreateRectRgn Lib "gdi32" _
    (ByVal X1 As Long, _
     ByVal Y1 As Long, _
     ByVal X2 As Long, _
     ByVal Y2 As Long) _
    As Long
    
Private Declare Function CombineRgn Lib "gdi32" _
    (ByVal hDestRgn As Long, _
     ByVal hSrcRgn1 As Long, _
     ByVal hSrcRgn2 As Long, _
     ByVal nCombineMode As Long) _
    As Long
  
Public Declare Function GetPixel Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long) _
    As Long
    
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) _
    As Long
'

Public Function RegionFromImage(fRgnForm As Object, _
                                picBackground As PictureBox, _
                                ByVal lBackColor As Long) _
                               As Long
                               
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



