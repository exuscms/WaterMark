VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Water Mark Master"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   10860
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10860
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox TempPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   25
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicBack 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   10845
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10875
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   1875
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CommandButton Command1 
            Caption         =   "확인"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox Cred 
            Caption         =   "빨"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Value           =   1  '확인
            Width           =   495
         End
         Begin VB.CheckBox Cgreen 
            Caption         =   "초"
            Height          =   255
            Left            =   720
            TabIndex        =   16
            Top             =   120
            Value           =   1  '확인
            Width           =   495
         End
         Begin VB.CheckBox Cblue 
            Caption         =   "파"
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            Top             =   120
            Value           =   1  '확인
            Width           =   615
         End
      End
      Begin VB.PictureBox Backup 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  '없음
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   0
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox WaterMark 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  '없음
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   0
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1455
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   0
         Left            =   0
         MousePointer    =   8  'NW SE 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   1
         Left            =   240
         MousePointer    =   7  'N S크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   2
         Left            =   480
         MousePointer    =   6  'NE SW 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   3
         Left            =   480
         MousePointer    =   9  'W E 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   4
         Left            =   480
         MousePointer    =   8  'NW SE 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   5
         Left            =   240
         MousePointer    =   7  'N S크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   6
         Left            =   0
         MousePointer    =   6  'NE SW 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox picPatrate 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   7
         Left            =   0
         MousePointer    =   9  'W E 크기 조정
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.TextBox Inputs 
         BorderStyle     =   0  '없음
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Text            =   "TEXT"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CDBackGround 
         Left            =   1200
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CDWaterMark 
         Left            =   960
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Watermark2 
         Height          =   1455
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "TEXT"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuBackGroundOpen 
         Caption         =   "배경 열기(&B)"
      End
      Begin VB.Menu mnuWaterMark 
         Caption         =   "워터마크 열기(&W)"
      End
      Begin VB.Menu mnu234 
         Caption         =   "투명 워터마크 열기(&W)"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "그림 저장(&S)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "프로그램 종료(&Q)"
      End
   End
   Begin VB.Menu mnuEffect 
      Caption         =   "효과(&E)"
      Begin VB.Menu mnuText 
         Caption         =   "글자(&X)"
      End
      Begin VB.Menu mnuT 
         Caption         =   "투명도(&T)"
         Begin VB.Menu mnuO 
            Caption         =   "불투명(&O)"
         End
         Begin VB.Menu mnuTM 
            Caption         =   "배경투명(&R)"
         End
      End
      Begin VB.Menu mnuSpecial 
         Caption         =   "특수 효과(&S)"
         Begin VB.Menu mnuOil 
            Caption         =   "수체화(&T)"
         End
         Begin VB.Menu mnuPixel 
            Caption         =   "픽셀화(&P)"
         End
         Begin VB.Menu mnuBlack 
            Caption         =   "흑백화(&B)"
         End
         Begin VB.Menu mnuColor 
            Caption         =   "색체 변경(&C)"
         End
         Begin VB.Menu mnu3ds 
            Caption         =   "3D 효과(&3)"
         End
         Begin VB.Menu mnuLight 
            Caption         =   "밝기 제거(&L)"
         End
         Begin VB.Menu mnuBlacks 
            Caption         =   "어두움 제거(&B)"
         End
         Begin VB.Menu mnumodern 
            Caption         =   "모던 아트(&M)"
         End
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "엣지(&E)"
         Begin VB.Menu mnuOri 
            Caption         =   "가로 엣지(&H)"
         End
         Begin VB.Menu mnuOrib1 
            Caption         =   "가로 흑백 엣지(&H)"
         End
         Begin VB.Menu mnuOri2 
            Caption         =   "세로 엣지(&V)"
         End
         Begin VB.Menu mnuOrib2 
            Caption         =   "세로 흑백 엣지(&H)"
         End
         Begin VB.Menu mnu3d 
            Caption         =   "3D엣지(&3)"
         End
      End
      Begin VB.Menu mnuETC 
         Caption         =   "기타(&E)"
         Begin VB.Menu mnuMulti 
            Caption         =   "멀티플라이(&M)"
         End
         Begin VB.Menu mnuDivine 
            Caption         =   "디바인(&D)"
         End
         Begin VB.Menu mnuLights 
            Caption         =   "밝게(&L)"
         End
         Begin VB.Menu mnuBlackss 
            Caption         =   "어둡게(&B)"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "색상 뒤집기(&I)"
         End
         Begin VB.Menu mnuFog 
            Caption         =   "구름(&F)"
         End
         Begin VB.Menu mnuX 
            Caption         =   "X 블러(&X)"
         End
         Begin VB.Menu mnuY 
            Caption         =   "Y 블러(&Y)"
         End
         Begin VB.Menu mnuShadow 
            Caption         =   "그림자(&S)"
         End
      End
      Begin VB.Menu mnuM 
         Caption         =   "거울(&M)"
         Begin VB.Menu mnuLR 
            Caption         =   "왼쪽에서 오른쪽으로(&L)"
         End
         Begin VB.Menu mnuRL 
            Caption         =   "오른쪽에서 왼쪽으로(&R)"
         End
         Begin VB.Menu mnuTD 
            Caption         =   "위에서 아래로(&T)"
         End
         Begin VB.Menu mnuDownTop 
            Caption         =   "아래에서 위로(&D)"
         End
      End
      Begin VB.Menu mnuF 
         Caption         =   "뒤집기(&F)"
         Begin VB.Menu mnuH 
            Caption         =   "좌우 전환(&H)"
         End
         Begin VB.Menu mnuV 
            Caption         =   "상하 전환(&V)"
         End
      End
      Begin VB.Menu mnuOption 
         Caption         =   "옵션(&O)"
      End
   End
   Begin VB.Menu mnuMovie 
      Caption         =   "동영상(&M)"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const Red As Integer = 3
Private Const Green As Integer = 2
Private Const Blue As Integer = 1

Dim Pixel
Dim Pixel2

Dim Rred
Dim Ggreen
Dim Bblue

Dim RR1
Dim GG1
Dim BB1

Dim RR2
Dim GG2
Dim BB2

Dim RR3
Dim GG3
Dim BB3

Dim Q As String
Dim Q2 As String

Dim Temp As Integer
Dim Temp2 As Integer

Dim XXX As Integer
Dim YYY As Integer

Dim XX As Integer
Dim YY As Integer

Dim RR As Integer
Dim RG As Integer
Dim RB As Integer

Dim CurX
Dim CurY

Dim JB As Byte

Dim token    As Long
Dim Counter As Long

Public ark As Integer
Public ark2 As Integer
Public actark As Integer
Public actark2 As Integer
Public tmpX, tmpy As Long ' Temp X, Y
Public Obj_S As Control
Public MouseD_p As Boolean
Public bMoving As Boolean
Public xStart, yStart As Long

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long     '이미지 가로 크기
    biHeight As Long    '이미지 세로 크기
    biPlanes As Integer
    biBitCount As Integer   '픽셀별 색상의 비트수
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long      '파일 사이즈(byte)
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long   '실제 이미지 데이터가 시작되는 주소
End Type

Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
     'bmiColors As RGBQUAD
End Type

Private Sub GetRGB(ByVal col As String)
On Error Resume Next
    Bblue = col \ (256 ^ 2)
    Ggreen = (col - Bblue * 256 ^ 2) \ 256
    Rred = (col - Bblue * 256 ^ 2 - Ggreen * 256) '\ 256
End Sub

Private Sub Command1_Click()
Picture2.Visible = False
End Sub

Private Sub Form_Resize()
PicBack.Width = Me.ScaleWidth
PicBack.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
FreeGDIPlus token
End
End Sub

Private Sub Inputs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    label(Index).Caption = Inputs(Index).Text
    Unload Inputs(Index)
End If
End Sub

Private Sub label_DblClick(Index As Integer)

Load Inputs(Index)
Inputs(Index).Text = label(Index).Caption
Inputs(Index).Width = label(Index).Width
Inputs(Index).Height = label(Index).Height
Inputs(Index).Left = label(Index).Left
Inputs(Index).Top = label(Index).Top
Inputs(Index).Visible = True

End Sub

Private Sub label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Position_P label(Index)
    Set Obj_S = label(Index)
    xStart = x: yStart = y
    bMoving = True
    actark2 = Index
End If

End Sub

Private Sub label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim lft As Long, tp As Long

If bMoving Then
    lft = label(Index).Left + x - xStart
    tp = label(Index).Top + y - yStart
    If lft <= 0 Then lft = 0
    If tp <= 0 Then tp = 0
    label(Index).Move lft, tp
    Position_P label(Index)
    Set Obj_S = label(Index)
End If

End Sub

Private Sub label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
bMoving = False
End Sub

Private Sub mnu234_Click()

    CDWaterMark.DialogTitle = "워터마크 불러오기"
    CDWaterMark.ShowOpen
    
    If CDWaterMark.filename <> "" Then
        ark = ark + 1
        Load Watermark2(ark)
        Watermark2(ark).Visible = True
        Watermark2(ark).Picture = LoadPicture(CDWaterMark.filename)
        Watermark2(ark).Stretch = True
    End If
    
End Sub

Private Sub mnu3d_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q2 = InputBox("3D 그리드 값을 입력하시오", "", "4")
    If Q2 = "" Then Exit Sub
    Q = InputBox("밝기를 입력하시오 (높을수록 어두워짐)", "", "10")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1 Step Q2 + 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1 Step Q2 + 1
    
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    
    GetRGB Pixel
    Rred = Rred - Q
    Ggreen = Ggreen - Q
    Bblue = Bblue - Q
    
    
    For Counter = 1 To Q2
    SetPixelV WaterMark(actark).hdc, XXX + Counter, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    For Counter = 1 To Q2
    SetPixelV WaterMark(actark).hdc, XXX - Counter, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    For Counter = 1 To Q2
    SetPixelV WaterMark(actark).hdc, XXX, YYY + Counter, RGB(Rred, Ggreen, Bblue)
    Next
    For Counter = 1 To Q2
    SetPixelV WaterMark(actark).hdc, XXX, YYY - Counter, RGB(Rred, Ggreen, Bblue)
    Next
    
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnu3ds_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오(낮을 수록 효과가 깊어짐)", "", "6")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    Temp = (Rred + Ggreen + Bblue)
    Temp = Temp / 3
    
    WaterMark(actark).ForeColor = RGB(Rred, Ggreen, Bblue)
    WaterMark(actark).Line (XXX, YYY)-(XXX, YYY - (Temp / Q))
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuBlack_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오 (0-255, 값이 높을수록 그림이 어두워짐)", "", "127")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    
    GetRGB Pixel
    
    Temp = (Rred + Ggreen + Bblue)
    Temp = (Temp / 3)
    
    If Val(Temp) >= Q Then
    Pixel = vbWhite
    Else
    Pixel = vbBlack
    End If
    
    
    
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuBlacks_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "1,5")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Rred > 128 Then
    RR1 = Rred - 128
    Else
    RR1 = 128 - Rred
    End If
    
    If Ggreen > 128 Then
    GG1 = Ggreen - 128
    Else
    GG1 = 128 - Ggreen
    End If
    
    If Bblue > 128 Then
    BB1 = Bblue - 128
    Else
    BB1 = 128 - Bblue
    End If
    
    RR1 = RR1 / Q
    GG1 = GG1 / Q
    BB1 = BB1 / Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(RR1, GG1, BB1)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuBlackss_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "30")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then Rred = Rred - Q
    If Cgreen.Value = 1 Then Ggreen = Ggreen - Q
    If Cblue.Value = 1 Then Bblue = Bblue - Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuColor_Click()
On Error GoTo ja
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    CM.CancelError = True
    CM.ShowColor
    GetRGB CM.Color
    RR3 = Rred
    GG3 = Ggreen
    BB3 = Bblue
    On Error Resume Next
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    Temp = (Rred + Ggreen + Bblue)
    Temp = Temp / 3
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB((RR3 + Temp), (GG3 + Temp), (BB3 + Temp))
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
    Exit Sub
ja:
    Exit Sub
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuDivine_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "1,5")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then Rred = Rred / Q
    If Cgreen.Value = 1 Then Ggreen = Ggreen / Q
    If Cblue.Value = 1 Then Bblue = Bblue / Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuDownTop_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To (WaterMark(actark).ScaleHeight / 2) - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, WaterMark(actark).ScaleHeight - YYY)
    SetPixelV WaterMark(actark).hdc, XXX, YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuFog_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "30")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then
    If Val(Rred) > 127 Then
    Rred = Rred - Q
    If Rred < 127 Then Rred = 127
    Else
    Rred = Rred + Q
    If Rred > 127 Then Rred = 127
    End If
    End If
    If Cgreen.Value = 1 Then
    If Val(Ggreen) > 127 Then
    Ggreen = Ggreen - Q
    If Ggreen < 127 Then Ggreen = 127
    Else
    Ggreen = Ggreen + Q
    If Ggreen > 127 Then Ggreen = 127
    End If
    End If
    
    
    If Cblue.Value = 1 Then
    If Val(Bblue) > 127 Then
    Bblue = Bblue - Q
    If Bblue < 127 Then Bblue = 127
    Else
    Bblue = Bblue + Q
    If Bblue > 127 Then Bblue = 127
    End If
    End If
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuH_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    TempPic.Width = WaterMark(actark).Width
    TempPic.Height = WaterMark(actark).Height
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    SetPixelV TempPic.hdc, WaterMark(actark).ScaleWidth - (XXX + 1), YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    BitBlt WaterMark(actark).hdc, 0, 0, TempPic.ScaleWidth - 1, TempPic.ScaleHeight - 1, TempPic.hdc, 0, 0, &HCC0020
    
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuInvert_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then Rred = 255 - Rred
    If Cgreen.Value = 1 Then Ggreen = 255 - Ggreen
    If Cblue.Value = 1 Then Bblue = 255 - Bblue
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuLight_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "3")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Rred > 128 Then
    RR1 = Rred - 128
    Else
    RR1 = 128 - Rred
    End If
    
    If Ggreen > 128 Then
    GG1 = Ggreen - 128
    Else
    GG1 = 128 - Ggreen
    End If
    
    If Bblue > 128 Then
    BB1 = Bblue - 128
    Else
    BB1 = 128 - Bblue
    End If
    
    RR1 = RR1 * Q
    GG1 = GG1 * Q
    BB1 = BB1 * Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(RR1, GG1, BB1)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuLights_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "30")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then Rred = Rred + Q
    If Cgreen.Value = 1 Then Ggreen = Ggreen + Q
    If Cblue.Value = 1 Then Bblue = Bblue + Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuLR_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To (WaterMark(actark).ScaleWidth / 2) - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    SetPixelV WaterMark(actark).hdc, WaterMark(actark).ScaleWidth - XXX, YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnumodern_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Rred > 127 Then
    RR1 = 255 - Rred
    Rred = Rred / RR1
    Else
    RR1 = 255 - Rred
    Rred = Rred * RR1
    End If
    
    If Ggreen > 127 Then
    GG1 = 255 - Ggreen
    Ggreen = Ggreen / GG1
    Else
    GG1 = 255 - Ggreen
    Ggreen = Ggreen * GG1
    End If
    
    If Bblue > 127 Then
    BB1 = 255 - Bblue
    Bblue = Bblue / BB1
    Else
    BB1 = 255 - Bblue
    Bblue = Bblue * BB1
    End If
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuMovie_Click()
frmavi.Show
End Sub

Private Sub mnuMulti_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "1,5")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    
    If Cred.Value = 1 Then Rred = Rred * Q
    If Cgreen.Value = 1 Then Ggreen = Ggreen * Q
    If Cblue.Value = 1 Then Bblue = Bblue * Q
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuO_Click()
DoAlphablend PicBack, WaterMark(actark), 30
End Sub

Private Sub mnuoil_Click()
frmoil.Show
End Sub

Private Sub mnuOption_Click()
Picture2.Visible = True
End Sub

Private Sub mnuOri_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오 (값이 높을수록 밝아짐)", "", "4")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    
    Pixel2 = GetPixel(WaterMark(actark).hdc, XXX + 2, YYY)
    Pixel = GetPixel(WaterMark(actark).hdc, XXX + 1, YYY)
    
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    
    If RR1 = RR2 Then RR3 = 0
    If RR1 > RR2 Then
    RR3 = RR1 - RR2
    Else
    RR3 = RR2 - RR1
    End If
    
    If GG1 = GG2 Then GG3 = 0
    If GG1 > GG2 Then
    GG3 = GG1 - GG2
    Else
    GG3 = GG2 - GG1
    End If
    
    If BB1 = BB2 Then BB3 = 0
    If BB1 > BB2 Then
    BB3 = BB1 - BB2
    Else
    BB3 = BB2 - BB1
    End If
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(RR3 * Q, GG3 * Q, BB3 * Q)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuOri2_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오 (높을수록 밝아짐)", "", "4")
    If Q = "" Then Exit Sub
    
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    
    Pixel2 = GetPixel(WaterMark(actark).hdc, XXX, YYY + 2)
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY + 1)
    
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    
    If RR1 = RR2 Then RR3 = 0
    If RR1 > RR2 Then
    RR3 = RR1 - RR2
    Else
    RR3 = RR2 - RR1
    End If
    
    If GG1 = GG2 Then GG3 = 0
    If GG1 > GG2 Then
    GG3 = GG1 - GG2
    Else
    GG3 = GG2 - GG1
    End If
    
    If BB1 = BB2 Then BB3 = 0
    If BB1 > BB2 Then
    BB3 = BB1 - BB2
    Else
    BB3 = BB2 - BB1
    End If
    
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, RGB(RR3 * Q, GG3 * Q, BB3 * Q)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuOrib1_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오 (0-255, 높을수록 엣지가 없어짐)", "", "7")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    
    Pixel2 = GetPixel(WaterMark(actark).hdc, XXX + 2, YYY)
    Pixel = GetPixel(WaterMark(actark).hdc, XXX + 1, YYY)
    
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    
    Temp = (RR1 + GG1 + BB1)
    Temp = (Temp / 3)
    
    Temp2 = (RR2 + GG2 + BB2)
    Temp2 = (Temp2 / 3)
    
    If Temp = Temp2 Then Pixel = vbWhite
    If Val(Temp) > Val(Temp2) Then
    If Val(Temp) - Val(Temp2) >= Q Then
    Pixel = vbBlack
    Else
    Pixel = vbWhite
    End If
    Else
    If Val(Temp2) - Val(Temp) >= Q Then
    Pixel = vbBlack
    Else
    Pixel = vbWhite
    End If
    End If
    
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
End Sub

Private Sub mnuOrib2_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오 (0-255, 높을수록 엣지가 없어짐)", "", "7")
    If Q = "" Then Exit Sub
    
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    
    Pixel2 = GetPixel(WaterMark(actark).hdc, XXX, YYY + 2)
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY + 1)
    
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    
    Temp = (RR1 + GG1 + BB1)
    Temp = (Temp / 3)
    
    Temp2 = (RR2 + GG2 + BB2)
    Temp2 = (Temp2 / 3)
    
    If Temp = Temp2 Then Pixel = vbWhite
    If Val(Temp) > Val(Temp2) Then
    If Val(Temp) - Val(Temp2) >= Q Then
    Pixel = vbBlack
    Else
    Pixel = vbWhite
    End If
    Else
    If Val(Temp2) - Val(Temp) >= Q Then
    Pixel = vbBlack
    Else
    Pixel = vbWhite
    End If
    End If
    
    
    SetPixelV WaterMark(actark).hdc, XXX, YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuPixel_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    Q = InputBox("값을 입력하시오", "", "5")
    If Q = "" Then Exit Sub
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1 Step Q
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1 Step Q
    Pixel = GetPixel(WaterMark(actark).hdc, XXX + 1, YYY + 1)
    WaterMark(actark).Line (XXX, YYY)-(XXX + Q, YYY + Q), Pixel, BF
    
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuRL_Click()
If Not actark = 0 Then
WaterMark(actark).ScaleMode = 3
For YYY = 0 To WaterMark(actark).ScaleHeight - 1
For XXX = 0 To (WaterMark(actark).ScaleWidth / 2) - 1
Pixel = GetPixel(WaterMark(actark).hdc, WaterMark(actark).ScaleWidth - XXX, YYY)
SetPixelV WaterMark(actark).hdc, XXX, YYY, Pixel
Next
WaterMark(actark).Refresh
Next
WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuSave_Click()
Dim i As Long
For i = 0 To 7
Me.picPatrate(i).Visible = False
Next i
PicBack.ScaleMode = 3
SaveFormImageToFile frmmain, Picture1, App.Path & "\" & Date & Hour(Time) & Minute(Time) & Second(Time) & ".jpg"
MsgBox App.Path & "\" & Date & Hour(Time) & Minute(Time) & Second(Time) & ".jpg"
PicBack.ScaleMode = 1
End Sub

Public Sub SaveFormImageToFile(ByRef ContainerForm As Form, ByRef PictureBoxControl As PictureBox, ByVal ImageFileName As String)
Dim FormInsideWidth As Long
Dim FormInsideHeight As Long
Dim PictureBoxLeft As Long
Dim PictureBoxTop As Long
Dim PictureBoxWidth As Long
Dim PictureBoxHeight As Long
Dim FormAutoRedrawValue As Boolean

With PictureBoxControl
    'Set PictureBox properties
    .Visible = False
    .AutoRedraw = True
    .Appearance = 0 ' Flat
    .AutoSize = False
    .BorderStyle = 0 'No border
    
    'Store PictureBox Original Size and location Values
    PictureBoxHeight = .Height: PictureBoxWidth = .Width: PictureBoxLeft = .Left: PictureBoxTop = .Top
    
    'Make PictureBox to size to inside of form.
    .Align = vbAlignTop: .Align = vbAlignLeft
    DoEvents
    
    FormInsideHeight = .Height: FormInsideWidth = .Width
    
    'Restore PictureBox Original Size and location Values
    .Align = vbAlignNone
    .Height = FormInsideHeight: .Width = FormInsideWidth: .Left = PictureBoxLeft: .Top = PictureBoxTop
    
    FormAutoRedrawValue = ContainerForm.AutoRedraw
    ContainerForm.AutoRedraw = False
    DoEvents
    
    'Copy Form Image to Picture Box
    BitBlt .hdc, 0, 0, FormInsideWidth / Screen.TwipsPerPixelX, FormInsideHeight / Screen.TwipsPerPixelY, ContainerForm.hdc, 0, 0, vbSrcCopy
    DoEvents
    SaveJPG .Image, ImageFileName, 70
    DoEvents
    
    ContainerForm.AutoRedraw = FormAutoRedrawValue
    DoEvents
End With
End Sub

Private Sub mnuTD_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To (WaterMark(actark).ScaleHeight / 2) - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    SetPixelV WaterMark(actark).hdc, XXX, WaterMark(actark).ScaleHeight - YYY, Pixel
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuText_Click()
ark2 = ark2 + 1
Load label(ark2)
label(ark2).Visible = True
End Sub

Private Sub mnuTM_Click()
If Not actark = 0 Then
    RegionFromImage WaterMark(actark), WaterMark(actark), GetPixel(WaterMark(actark).hdc, 0, 0)
End If
End Sub

Private Sub mnuV_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    TempPic.Width = WaterMark(actark).Width
    TempPic.Height = WaterMark(actark).Height
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    SetPixelV TempPic.hdc, XXX, WaterMark(actark).ScaleHeight - (YYY + 1), Pixel
    Next
    WaterMark(actark).Refresh
    Next
    BitBlt WaterMark(actark).hdc, 0, 0, TempPic.ScaleWidth - 1, TempPic.ScaleHeight - 1, TempPic.hdc, 0, 0, &HCC0020
    
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuX_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    If XXX < WaterMark(actark).ScaleWidth - 3 Then Pixel2 = GetPixel(WaterMark(actark).hdc, XXX + 2, YYY)
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    If Cred.Value = 1 Then Rred = (RR1 + RR2) / 2
    If Cgreen.Value = 1 Then Ggreen = (GG1 + GG2) / 2
    If Cblue.Value = 1 Then Bblue = (BB1 + BB2) / 2
    SetPixelV WaterMark(actark).hdc, XXX + 1, YYY, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub mnuY_Click()
On Error Resume Next
WaterMark(actark).ScaleMode = 3
If Not actark = 0 Then
    For XXX = 0 To WaterMark(actark).ScaleWidth - 1
    For YYY = 0 To WaterMark(actark).ScaleHeight - 1
    
    Pixel = GetPixel(WaterMark(actark).hdc, XXX, YYY)
    GetRGB Pixel
    RR1 = Rred
    GG1 = Ggreen
    BB1 = Bblue
    If YYY < WaterMark(actark).ScaleHeight - 3 Then Pixel2 = GetPixel(WaterMark(actark).hdc, XXX, YYY + 2)
    GetRGB Pixel2
    RR2 = Rred
    GG2 = Ggreen
    BB2 = Bblue
    If Cred.Value = 1 Then Rred = (RR1 + RR2) / 2
    If Cgreen.Value = 1 Then Ggreen = (GG1 + GG2) / 2
    If Cblue.Value = 1 Then Bblue = (BB1 + BB2) / 2
    SetPixelV WaterMark(actark).hdc, XXX, YYY + 1, RGB(Rred, Ggreen, Bblue)
    Next
    WaterMark(actark).Refresh
    Next
    WaterMark(actark).Refresh
End If
WaterMark(actark).ScaleMode = 1
End Sub

Private Sub PicBack_Click()
Dim i As Long
For i = 0 To 7
Me.picPatrate(i).Visible = False
Next i
actark = 0
End Sub

Public Sub picPatrate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
tmpX = x
tmpy = y
MouseD_p = True
End Sub

Public Sub picPatrate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
 
  If Obj_S.Width < 100 Then
   Obj_S.Width = 101
   MouseD_p = False
   Exit Sub
   End If
  If Obj_S.Height < 100 Then
   Obj_S.Height = 101
   MouseD_p = False
   Exit Sub
  End If
  If MouseD_p = True Then
   'Colt Stanga sus
   If Index = 0 Then
    Obj_S.Move Obj_S.Left + (x - tmpX), Obj_S.Top + (y - tmpy), Obj_S.Width - x + tmpX, Obj_S.Height - y + tmpy
    picPatrate(0).Move picPatrate(0).Left + (x - tmpX), picPatrate(0).Top + (y - tmpy)
    Position_P Obj_S, 0
   End If
   ' Mijloc sus
   If Index = 1 Then
   Obj_S.Move Obj_S.Left, Obj_S.Top + (y - tmpy), Obj_S.Width, Obj_S.Height - y + tmpy
   picPatrate(1).Move picPatrate(1).Left, picPatrate(1).Top + (y - tmpy)
   Position_P Obj_S, 1
   End If
   'Colt Dreapta sus
   If Index = 2 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top + (y - tmpy), Obj_S.Width + x - tmpX, Obj_S.Height - y + tmpy
    picPatrate(2).Move picPatrate(2).Left + (x - tmpX), picPatrate(2).Top + (y - tmpy)
   Position_P Obj_S, 2
   End If
   'Mijloc Dreapta
   If Index = 3 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width + x - tmpX, Obj_S.Height
    picPatrate(3).Move picPatrate(3).Left + (x - tmpX), picPatrate(3).Top
   Position_P Obj_S, 3
   End If
   'Colt dreapta jos
   If Index = 4 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width + x - tmpX, Obj_S.Height + y - tmpy
    picPatrate(4).Move picPatrate(4).Left + (x - tmpX), picPatrate(4).Top + (y - tmpy)
   Position_P Obj_S, 4
   End If
   'Mijloc jos
   If Index = 5 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width, Obj_S.Height + y - tmpy
    picPatrate(5).Move picPatrate(5).Left, picPatrate(5).Top + (y - tmpy)
   Position_P Obj_S, 5
   End If
   'Colt Stanga jos
   If Index = 6 Then
    Obj_S.Move Obj_S.Left + (x - tmpX), Obj_S.Top, Obj_S.Width - x + tmpX, Obj_S.Height + y - tmpy
    picPatrate(6).Move picPatrate(6).Left + (x - tmpX), picPatrate(6).Top + (y - tmpy)
    Position_P Obj_S, 6
   End If
   'Mijloc Jos
   If Index = 7 Then
    Obj_S.Move Obj_S.Left + (x - tmpX), Obj_S.Top, Obj_S.Width - x + tmpX, Obj_S.Height
    picPatrate(7).Move picPatrate(7).Left + (x - tmpX), picPatrate(7).Top
    Position_P Obj_S, 7
   End If
  End If

WaterMark(actark) = Resize(Backup(actark).Picture.Handle, Backup(actark).Picture.Type, WaterMark(actark).Width / Screen.TwipsPerPixelX, WaterMark(actark).Height / Screen.TwipsPerPixelY, vbBlack, False)
'If WaterMark(actark).Tag = 1 Then
'    RegionFromImage WaterMark(actark), WaterMark(actark), GetPixel(WaterMark(actark).hdc, 0, 0)
'End If
WaterMark(actark).ScaleMode = 3
Dim clr As Long, w As Long, h As Long
w = WaterMark(actark).ScaleWidth
h = WaterMark(actark).ScaleHeight
clr = WaterMark(actark).Point(0, 0)
Call TransparentBlt(Me.hdc, ScaleWidth - w, 0, w, h, WaterMark(actark).hdc, 0, 0, w, h, clr)
Me.Refresh
WaterMark(actark).ScaleMode = 1
End Sub

Public Sub picPatrate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 MouseD_p = False
End Sub

Public Function Position_P(ByRef obj As Control, Optional Ignorea As Integer = -1)

   With frmmain
    'Colt Stanga sus
    If Ignorea <> 0 Then
    .picPatrate(0).Move obj.Left - 115, obj.Top - 115
    .picPatrate(0).Visible = True
    End If
    'Mijloc sus
    If Ignorea <> 1 Then
    .picPatrate(1).Move obj.Left + obj.Width / 2 - 50, obj.Top - 115
    .picPatrate(1).Visible = True
    End If
    'Colt Dreapta sus
    If Ignorea <> 2 Then
    .picPatrate(2).Move obj.Left + obj.Width + 15, obj.Top - 115
    .picPatrate(2).Visible = True
    End If
    'Mijloc Dreapta
    If Ignorea <> 3 Then
    .picPatrate(3).Move obj.Left + obj.Width + 15, obj.Top + obj.Height / 2 - 50
    .picPatrate(3).Visible = True
    End If
    'Colt Dreapta jos
    If Ignorea <> 4 Then
    .picPatrate(4).Move obj.Left + obj.Width + 15, obj.Top + obj.Height + 15
    .picPatrate(4).Visible = True
    End If
    'Mijloc jos
    If Ignorea <> 5 Then
    .picPatrate(5).Move obj.Left + obj.Width / 2 - 50, obj.Top + obj.Height + 15
    .picPatrate(5).Visible = True
    End If
    'Colt Stanga jos
    If Ignorea <> 6 Then
    .picPatrate(6).Move obj.Left - 115, obj.Top + obj.Height + 15
    .picPatrate(6).Visible = True
    End If
    'Mijloc Stanga
    If Ignorea <> 7 Then
    .picPatrate(7).Move obj.Left - 115, obj.Top + obj.Height / 2 - 15
    .picPatrate(7).Visible = True
    End If
    
   End With
  
End Function

Private Sub Form_Load()
    token = InitGDIPlus
    ark = 0
    actark = 0
End Sub

Private Sub mnuBackGroundOpen_Click()

    CDBackGround.DialogTitle = "배경 불러오기"
    CDBackGround.ShowOpen
    
    If CDBackGround.filename <> "" Then
        PicBack.Picture = LoadPictureGDIPlus(CDBackGround.filename)
        Me.Width = PicBack.Width
        Me.Height = PicBack.Height
    End If
    
End Sub

Private Sub mnuWaterMark_Click()

    CDWaterMark.DialogTitle = "워터마크 불러오기"
    CDWaterMark.ShowOpen
    
    If CDWaterMark.filename <> "" Then
        ark = ark + 1
        Load WaterMark(ark)
        Load Backup(ark)
        WaterMark(ark).Visible = True
        WaterMark(ark).Picture = LoadPictureGDIPlus(CDWaterMark.filename)
        Backup(ark).Picture = WaterMark(ark).Picture
    End If
    
End Sub

Private Sub WaterMark_DblClick(Index As Integer)
On Error Resume Next
Unload WaterMark(Index)
WaterMark(Index).Visible = False
Dim i As Long
For i = 0 To 7
Me.picPatrate(i).Visible = False
Next i
End Sub

Private Sub WaterMark_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Position_P WaterMark(Index)
    Set Obj_S = WaterMark(Index)
    xStart = x: yStart = y
    bMoving = True
    actark = Index
End If

End Sub

Private Sub WaterMark_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim lft As Long, tp As Long

If bMoving Then
    lft = WaterMark(Index).Left + x - xStart
    tp = WaterMark(Index).Top + y - yStart
    If lft <= 0 Then lft = 0
    If tp <= 0 Then tp = 0
    WaterMark(Index).Move lft, tp
    Position_P WaterMark(Index)
    Set Obj_S = WaterMark(Index)
End If

End Sub

Private Sub WaterMark_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMoving = False
End Sub





Private Sub watermark2_DblClick(Index As Integer)
On Error Resume Next
Unload Watermark2(Index)
Watermark2(Index).Visible = False
Dim i As Long
For i = 0 To 7
Me.picPatrate(i).Visible = False
Next i
End Sub

Private Sub watermark2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Position_P Watermark2(Index)
    Set Obj_S = Watermark2(Index)
    xStart = x: yStart = y
    bMoving = True
    actark = Index
End If

End Sub

Private Sub watermark2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim lft As Long, tp As Long

If bMoving Then
    lft = Watermark2(Index).Left + x - xStart
    tp = Watermark2(Index).Top + y - yStart
    If lft <= 0 Then lft = 0
    If tp <= 0 Then tp = 0
    Watermark2(Index).Move lft, tp
    Position_P Watermark2(Index)
    Set Obj_S = Watermark2(Index)
End If

End Sub

Private Sub watermark2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
bMoving = False
End Sub
