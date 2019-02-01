VERSION 5.00
Begin VB.Form frmopt 
   BorderStyle     =   5  '크기 조정 가능 도구 창
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Index           =   1
         Left            =   645
         Max             =   100
         Min             =   1
         TabIndex        =   3
         Top             =   585
         Value           =   50
         Width           =   2895
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Index           =   0
         Left            =   645
         Max             =   100
         TabIndex        =   2
         Top             =   285
         Value           =   100
         Width           =   2895
      End
      Begin VB.HScrollBar vScroll1 
         Height          =   270
         Left            =   645
         Max             =   360
         TabIndex        =   1
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "크기"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "투명"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "회전"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "회전"
      Height          =   180
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Rad As Currency = 1.74532925199433E-02

Private Sub HScroll1_Change(index%)
vScroll1_Change
End Sub

Private Sub HScroll1_Scroll(index%)
vScroll1_Change
End Sub

Private Sub vScroll1_Change()
    DoEvents
    frmmain.Effect.Cls
    frmmain.Effect.Picture = frmmain.WaterMark(frmmain.actark).Picture
    RotBlt frmmain.Effect.hdc, vScroll1.Value * Rad, 128, 128, frmmain.PicBack.Width, frmmain.PicBack.Height, frmmain.PicBack.Image.handle, &HFF00FF, HScroll1(0).Value / 100, HScroll1(1).Value / 50
    frmmain.Effect.Refresh
    frmmain.WaterMark(frmmain.actark).Picture = frmmain.Effect.Picture
End Sub

Private Sub vScroll1_Scroll()
vScroll1_Change
End Sub

