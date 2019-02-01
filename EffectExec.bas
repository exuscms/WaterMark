Attribute VB_Name = "EffectExec"
'FX Functions Module
'Copyright (c) By Marco Samy - 5/2003

'Enumerating Available Effects
Public Enum Effects
[None] = 1
[Fade] = 2
[Dissolove] = 3
[Horizontal Bars] = 4
[Vertical Bars] = 5
[Box IN] = 6
[Box OUT] = 7
[Pixelate OUT IN] = 8
[Chess Boxes] = 9
[Move Left - Right] = 10
[Move Right - Left] = 11
[Move Down - Up] = 12
[Move Up - Down] = 13
[Brightness IN OUT] = 14
[Diffuse] = 15
[Blackness IN OUT] = 16
[TV] = 17
[Random Horizontal Lines] = 18
[Random Vertical Lines] = 19
[Wipe Vertical] = 20
[Wip Horizontal] = 21
[Slide Up] = 22
[Slide Down] = 23
End Enum
'Begin Executable Functions
Function exeNone(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
If sValTo2 > 50 Then DestBits = Bits2 Else DestBits = Bits1
End Function
Function exeFade(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = (Bits1.Bits(0, X, Y) * (100 - sValTo2) / 100) + (Bits2.Bits(0, X, Y) * (sValTo2) / 100)
                DestBits.Bits(1, X, Y) = (Bits1.Bits(1, X, Y) * (100 - sValTo2) / 100) + (Bits2.Bits(1, X, Y) * (sValTo2) / 100)
                DestBits.Bits(2, X, Y) = (Bits1.Bits(2, X, Y) * (100 - sValTo2) / 100) + (Bits2.Bits(2, X, Y) * (sValTo2) / 100)
        Next X
    Next Y
End Function
Function exeDissolove(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long
    For Y = 0 To -Bits1.Header.biHeight - 1 Step ((100 - sValTo2) / 8) + 1
        For X = 0 To Bits1.Header.biWidth - 1 Step ((100 - sValTo2) / 8) + 1
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
End Function
Function exeVerticalBars(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long, Z As Long
  For Z = 1 To Fix(sValTo2 / 10)
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1 - Fix(sValTo2 / 10) Step 10
                TempBits.Bits(0, X + (Z - 1), Y) = Bits2.Bits(0, X + (Z - 1), Y)
                TempBits.Bits(1, X + (Z - 1), Y) = Bits2.Bits(1, X + (Z - 1), Y)
                TempBits.Bits(2, X + (Z - 1), Y) = Bits2.Bits(2, X + (Z - 1), Y)
        Next X
    Next Y
  Next Z
End Function
Function exeHorizontalBars(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long, Z As Long
  For Z = 1 To Fix(sValTo2 / 10)
    For Y = 0 To -Bits1.Header.biHeight - 1 - Fix(sValTo2 / 10) Step 10
        For X = 0 To Bits1.Header.biWidth - 1
                TempBits.Bits(0, X, Y + (Z - 1)) = Bits2.Bits(0, X, Y + (Z - 1))
                TempBits.Bits(1, X, Y + (Z - 1)) = Bits2.Bits(1, X, Y + (Z - 1))
                TempBits.Bits(2, X, Y + (Z - 1)) = Bits2.Bits(2, X, Y + (Z - 1))
        Next X
    Next Y
  Next Z
End Function
'stop here
Function exeBoxIN(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
'////////////////////////////////COMMENTED AREA/////////////////////////////////////////////////////////////////////////////////////
'to make box in we have to make 4 steps
'(1)-the top bar is ( max x , y from 0 to ((0.5 * (svalto2/100)* height)
'(2)-the left bar is (x from 0 to (0.5 * (svalto2/100)* width) , y from ((0.5 * (svalto2/100)* height) to ( (0.5 * ((100-svalto2)/100)* height)
'(3)-the bottom bar is ( max x , y from ( (0.5 * ((100-svalto2)/100)* height) to max
'(4)-the left bar is (x from ((0.5 * ((100-svalto2/100))* width) to max , y from ((0.5 * (svalto2/100)* height) to ( (0.5 * ((100-svalto2)/100)* height)
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
'here we make 1
    Dim X As Long, Y As Long
    For Y = 0 To (((0.5 * sValTo2) / 100) * (-Bits1.Header.biHeight - 1))
        For X = 0 To Bits1.Header.biWidth - 1
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y

'here we make 2
     For Y = (((0.5 * sValTo2) / 100) * (-Bits1.Header.biHeight - 1)) To (((100 - (0.5 * sValTo2)) / 100) * (-Bits1.Header.biHeight - 1))
        For X = 0 To (0.5 * (sValTo2 / 100) * (Bits1.Header.biWidth - 1))
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y

'here we make 3
    For Y = (((100 - (0.5 * sValTo2)) / 100) * (-Bits1.Header.biHeight - 1)) To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y

'here we make 4
    For Y = (((0.5 * sValTo2) / 100) * (-Bits1.Header.biHeight - 1)) To (((100 - (0.5 * sValTo2)) / 100) * (-Bits1.Header.biHeight - 1))
        For X = (Bits1.Header.biWidth - 1) - (((sValTo2) / 200) * (Bits1.Header.biWidth - 1)) To Bits1.Header.biWidth - 1
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y

End Function
Function exeBoxOUT(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
'////////////////////////////////COMMENTED AREA/////////////////////////////////////////////////////////////////////////////////////
'we will draw a box in the center with the second photo
'y from (((0.5 * sValTo2) / 100) * (-Bits1.Header.biHeight - 1)) to
'        (((0.5 * sValTo2) / 100) * (-Bits1.Header.biHeight - 1)) + (((100 - (0.5 * sValTo2)) / 100) * (-Bits1.Header.biHeight - 1))
'x from (((0.5 * sValTo2) / 100) * (Bits1.Header.biwidth - 1)) to
'        (((0.5 * sValTo2) / 100) * (-Bits1.Header.biwidth - 1)) + (((100 - (0.5 * sValTo2)) / 100) * (-Bits1.Header.biwidth - 1))
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim Wid As Long, Hie As Long, bWid As Long, bHie As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = Wid * sValTo2 / 100: bHie = Hie * sValTo2 / 100
    Dim X As Long, Y As Long
    For Y = ((Hie - bHie) / 2) To ((Hie - bHie) / 2) + bHie
        For X = ((Wid - bWid) / 2) To ((Wid - bWid) / 2) + bWid
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
End Function
Function exePixelateOUTIN(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
On Error Resume Next
    Dim X As Long, Y As Long
'draw pixlate out
  If sValTo2 < 50 Then
  sValTo2 = sValTo2 / 2
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = Bits1.Bits(0, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
                DestBits.Bits(1, X, Y) = Bits1.Bits(1, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
                DestBits.Bits(2, X, Y) = Bits1.Bits(2, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
        Next X
    Next Y
  Else
'draw pixleate in
  sValTo2 = 100 - sValTo2
'this to make it (IN)
  
  sValTo2 = sValTo2 / 2
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, (Fix((X + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2), (Fix((Y + 1) / (sValTo2 + 1)) * (sValTo2 + 1)) + (sValTo2 / 2))
        Next X
    Next Y
End If
End Function

Function exeChessBoxes(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
'8 x 8 screen boxes
    Dim X As Long, Y As Long, Z As Long, StartBlock As Single
    Dim Wid As Long, Hie As Long, bWid As Long, bHie As Long
    
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = Wid / 8: bHie = Hie / 8
    
    
    For I = StartBlock To 7 Step 1
        For Z = StartBlock To 7 Step 1
            For X = 0 To bWid * (sValTo2 / 100)
                For Y = 0 To bHie * (sValTo2 / 100)
                    TempBits.Bits(0, (I * bWid) + X, (Z * bHie) + Y) = Bits2.Bits(0, (I * bWid) + X, (Z * bHie) + Y)
                    TempBits.Bits(1, (I * bWid) + X, (Z * bHie) + Y) = Bits2.Bits(1, (I * bWid) + X, (Z * bHie) + Y)
                    TempBits.Bits(2, (I * bWid) + X, (Z * bHie) + Y) = Bits2.Bits(2, (I * bWid) + X, (Z * bHie) + Y)
                Next
            Next
        Next
    Next

End Function
Function exeMoveLeftRight(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim Wid As Long, Hie As Long, bWid As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = (sValTo2 / 100) * Wid
    
    Dim X As Long, Y As Long
    For Y = 0 To Hie
        For X = 0 To bWid
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, Wid - bWid + X, Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, Wid - bWid + X, Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, Wid - bWid + X, Y)
        Next X
    Next Y
End Function
Function exeMoveRightLeft(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim Wid As Long, Hie As Long, bWid As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = (sValTo2 / 100) * Wid
    
    Dim X As Long, Y As Long
    For Y = 0 To Hie
        For X = Wid - bWid To Wid
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X - (Wid - bWid), Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X - (Wid - bWid), Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X - (Wid - bWid), Y)
        Next X
    Next Y
End Function
Function exeMoveUpDown(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim Wid As Long, Hie As Long, bWid As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = (sValTo2 / 100) * Hie
    
    Dim X As Long, Y As Long
    For Y = 0 To bWid
        For X = 0 To Wid
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Hie - bWid + Y)
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Hie - bWid + Y)
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Hie - bWid + Y)
        Next X
    Next Y
End Function
Function exeMoveDownUp(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim Wid As Long, Hie As Long, bWid As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    bWid = (sValTo2 / 100) * Hie
    
    Dim X As Long, Y As Long
    For Y = Hie - bWid To Hie
        For X = 0 To Wid
                TempBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y - (Hie - bWid))
                TempBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y - (Hie - bWid))
                TempBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y - (Hie - bWid))
        Next X
    Next Y
End Function
Function exeBrightnessINOUT(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long
  If sValTo2 <= 50 Then
  sValTo2 = sValTo2 * 2
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = (Bits1.Bits(0, X, Y) * ((100 - sValTo2) / 100)) + (255 * (sValTo2) / 100)
                DestBits.Bits(1, X, Y) = (Bits1.Bits(1, X, Y) * ((100 - sValTo2) / 100)) + (255 * (sValTo2) / 100)
                DestBits.Bits(2, X, Y) = (Bits1.Bits(2, X, Y) * ((100 - sValTo2) / 100)) + (255 * (sValTo2) / 100)
        Next X
    Next Y
  Else
  sValTo2 = 100 - sValTo2
     For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = (Bits2.Bits(0, X, Y) * (100 - sValTo2) / 100) + (255 * (sValTo2) / 100)
                DestBits.Bits(1, X, Y) = (Bits2.Bits(1, X, Y) * (100 - sValTo2) / 100) + (255 * (sValTo2) / 100)
                DestBits.Bits(2, X, Y) = (Bits2.Bits(2, X, Y) * (100 - sValTo2) / 100) + (255 * (sValTo2) / 100)
        Next X
    Next Y
  End If
End Function
Function exeBlacknessINOUT(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long
  If sValTo2 <= 50 Then
  sValTo2 = sValTo2 * 2
    For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = (Bits1.Bits(0, X, Y) * ((100 - sValTo2) / 100))
                DestBits.Bits(1, X, Y) = (Bits1.Bits(1, X, Y) * ((100 - sValTo2) / 100))
                DestBits.Bits(2, X, Y) = (Bits1.Bits(2, X, Y) * ((100 - sValTo2) / 100))
        Next X
    Next Y
  Else
  sValTo2 = 100 - sValTo2
     For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
                DestBits.Bits(0, X, Y) = (Bits2.Bits(0, X, Y) * (100 - sValTo2) / 100)
                DestBits.Bits(1, X, Y) = (Bits2.Bits(1, X, Y) * (100 - sValTo2) / 100)
                DestBits.Bits(2, X, Y) = (Bits2.Bits(2, X, Y) * (100 - sValTo2) / 100)
        Next X
    Next Y
  End If
End Function
Function exeDiffuse(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
    Dim X As Long, Y As Long
    Dim GPx As Long, GPy As Long
If sValTo2 < 50 Then
        For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
            GPx = Int(Rnd * sValTo2) - DiffuseX + X
            GPy = Int(Rnd * sValTo2) - DiffuseY + Y
            If GPx >= 0 And GPx < Bits1.Header.biWidth Then
            If GPy >= 0 And GPy < -Bits1.Header.biHeight Then
                TempBits.Bits(0, GPx, GPy) = Bits1.Bits(0, X, Y)
                TempBits.Bits(1, GPx, GPy) = Bits1.Bits(1, X, Y)
                TempBits.Bits(2, GPx, GPy) = Bits1.Bits(2, X, Y)
            End If
            End If
        Next X
    Next Y
Else
sValTo2 = 50 - (sValTo2 - 50)
        For Y = 0 To -Bits1.Header.biHeight - 1
        For X = 0 To Bits1.Header.biWidth - 1
            GPx = Int(Rnd * sValTo2) - DiffuseX + X
            GPy = Int(Rnd * sValTo2) - DiffuseY + Y
            If GPx >= 0 And GPx < Bits1.Header.biWidth Then
            If GPy >= 0 And GPy < -Bits1.Header.biHeight Then
                TempBits.Bits(0, GPx, GPy) = Bits2.Bits(0, X, Y)
                TempBits.Bits(1, GPx, GPy) = Bits2.Bits(1, X, Y)
                TempBits.Bits(2, GPx, GPy) = Bits2.Bits(2, X, Y)
            End If
            End If
        Next X
    Next Y
End If
End Function
Function exeTV(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
    Dim v1 As Single, V2 As Single
    Dim FastDarken(255) As Single, FastLighten(255) As Single
    Dim X As Long, Y As Long, V As Single
If sValTo2 <= 50 Then
sValTo2 = sValTo2 * 2
DestBits = Bits1
Else
DestBits = Bits2
sValTo2 = 50 - (sValTo2 - 50)
End If

    v1 = sValTo2 / 100
    
    For X = 0 To 255
        FastDarken(X) = X + (X * v1)
        FastLighten(X) = X - (X * v1)
        If FastDarken(X) < 0 Then FastDarken(X) = 0
        If FastLighten(X) < 0 Then FastLighten(X) = 0
        If FastDarken(X) > 255 Then FastDarken(X) = 255
        If FastLighten(X) > 255 Then FastLighten(X) = 255
    Next X
    
    On Error Resume Next

    For Y = 0 To -DestBits.Header.biHeight - 1 Step 2
        For X = 0 To DestBits.Header.biWidth - 1
            DestBits.Bits(0, X, Y) = FastDarken(DestBits.Bits(0, X, Y))
            DestBits.Bits(1, X, Y) = FastDarken(DestBits.Bits(1, X, Y))
            DestBits.Bits(2, X, Y) = FastDarken(DestBits.Bits(2, X, Y))
            DestBits.Bits(0, X, Y + 1) = FastLighten(DestBits.Bits(0, X, Y + 1))
            DestBits.Bits(1, X, Y + 1) = FastLighten(DestBits.Bits(1, X, Y + 1))
            DestBits.Bits(2, X, Y + 1) = FastLighten(DestBits.Bits(2, X, Y + 1))
        Next X
    Next Y
End Function
Function exeRandomVerticalLines(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
On Error Resume Next
    Dim X As Long, Y As Long
    Dim LineAdd As Integer
    For X = 0 To Bits1.Header.biWidth - 1 Step 100 / (sValTo2 / 1)
       LineAdd = Int(Rnd * ((100 / (sValTo2 / 1) - 1)))
       For Y = 0 To -Bits1.Header.biHeight - 1
                TempBits.Bits(0, X + LineAdd, Y) = Bits2.Bits(0, X + LineAdd, Y)
                TempBits.Bits(1, X + LineAdd, Y) = Bits2.Bits(1, X + LineAdd, Y)
                TempBits.Bits(2, X + LineAdd, Y) = Bits2.Bits(2, X + LineAdd, Y)
        Next Y
    Next X
End Function
Function exeRandomHorizontalLines(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, TempBits As BITMAPINFO, sValTo2 As Integer)
On Error Resume Next
 Dim X As Long, Y As Long
    Dim LineAdd As Integer
    For Y = 0 To -Bits1.Header.biHeight - 1 Step 100 / (sValTo2 / 1)
        LineAdd = Int(Rnd * ((100 / (sValTo2 / 1) - 1)))
        For X = 0 To Bits1.Header.biWidth - 1
                TempBits.Bits(0, X, Y + LineAdd) = Bits2.Bits(0, X, Y + LineAdd)
                TempBits.Bits(1, X, Y + LineAdd) = Bits2.Bits(1, X, Y + LineAdd)
                TempBits.Bits(2, X, Y + LineAdd) = Bits2.Bits(2, X, Y + LineAdd)
        Next X
    Next Y
End Function
Function exeWipeVertical(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
Dim Wid As Long, Hie As Long, SideWid As Long
    Dim X As Long, Y As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    SideWid = (sValTo2 / 200) * Wid
    For Y = 0 To Hie
        For X = 0 To SideWid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
    
    For Y = 0 To Hie
        For X = (Wid - SideWid) To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
End Function
Function exeWipeHorizontal(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
Dim Wid As Long, Hie As Long, SideHie As Long
    Dim X As Long, Y As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    SideHie = (sValTo2 / 200) * Hie
    For Y = 0 To SideHie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
    
    For Y = (Hie - SideHie) To Hie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y)
        Next X
    Next Y
End Function
Function exeSlideUp(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
On Error Resume Next
Dim Wid As Long, Hie As Long, SideHie As Long
    Dim X As Long, Y As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    SideHie = ((100 - sValTo2) / 100) * (Hie)
If sValTo2 <= 50 Then
    
    For Y = 0 To SideHie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits1.Bits(0, X, (Hie - SideHie) + Y)
                DestBits.Bits(1, X, Y) = Bits1.Bits(1, X, (Hie - SideHie) + Y)
                DestBits.Bits(2, X, Y) = Bits1.Bits(2, X, (Hie - SideHie) + Y)
        Next X
    Next Y
    
    For Y = SideHie To Hie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, (SideHie - (Hie - SideHie)) + (Y - SideHie) + 1)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, (SideHie - (Hie - SideHie)) + (Y - SideHie) + 1)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, (SideHie - (Hie - SideHie)) + (Y - SideHie) + 1)
        Next X
    Next Y
    
Else
    For Y = SideHie To Hie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y - SideHie)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y - SideHie)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y - SideHie)
        Next X
    Next Y
    
    For Y = 0 To SideHie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits1.Bits(0, X, (((SideHie))) + (Y))
                DestBits.Bits(1, X, Y) = Bits1.Bits(1, X, (((SideHie))) + (Y))
                DestBits.Bits(2, X, Y) = Bits1.Bits(2, X, (((SideHie))) + (Y))
        Next X
    Next Y
    
End If
End Function
Function exeSlideDown(Bits1 As BITMAPINFO, Bits2 As BITMAPINFO, DestBits As BITMAPINFO, sValTo2 As Integer)
'On Error Resume Next
Dim Wid As Long, Hie As Long, SideHie As Long
    Dim X As Long, Y As Long
    Hie = (-Bits1.Header.biHeight - 1): Wid = Bits1.Header.biWidth - 1
    SideHie = ((100 - sValTo2) / 100) * (Hie)
If sValTo2 <= 50 Then
    
    For Y = (Hie - SideHie) To Hie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits1.Bits(0, X, Y - (Hie - SideHie))
                DestBits.Bits(1, X, Y) = Bits1.Bits(1, X, Y - (Hie - SideHie))
                DestBits.Bits(2, X, Y) = Bits1.Bits(2, X, Y - (Hie - SideHie))
        Next X
    Next Y
    
    For Y = 0 To (Hie - SideHie)
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, (Hie - SideHie) + (Y))
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, (Hie - SideHie) + (Y))
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, (Hie - SideHie) + (Y))
        Next X
    Next Y

    
Else
    
    For Y = 0 To (Hie - SideHie)
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits2.Bits(0, X, Y + SideHie)
                DestBits.Bits(1, X, Y) = Bits2.Bits(1, X, Y + SideHie)
                DestBits.Bits(2, X, Y) = Bits2.Bits(2, X, Y + SideHie)
        Next X
    Next Y
    
    For Y = (Hie - SideHie) To Hie
        For X = 0 To Wid
                DestBits.Bits(0, X, Y) = Bits1.Bits(0, X, ((Hie - SideHie) - SideHie) + (Y - (Hie - SideHie)))
                DestBits.Bits(1, X, Y) = Bits1.Bits(1, X, ((Hie - SideHie) - SideHie) + (Y - (Hie - SideHie)))
                DestBits.Bits(2, X, Y) = Bits1.Bits(2, X, ((Hie - SideHie) - SideHie) + (Y - (Hie - SideHie)))
        Next X
    Next Y
   
End If
End Function
