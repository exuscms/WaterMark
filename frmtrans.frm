VERSION 5.00
Begin VB.Form frmOil 
   BorderStyle     =   1  '단일 고정
   Caption         =   "수채화 효과"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "frmtrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   6525
   StartUpPosition =   2  '화면 가운데
   Begin VB.HScrollBar ScrBr 
      Height          =   255
      Left            =   1320
      Max             =   5
      Min             =   1
      TabIndex        =   1
      Top             =   120
      Value           =   3
      Width           =   1635
   End
   Begin VB.HScrollBar ScrSmooth 
      Height          =   255
      Left            =   4080
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "브러쉬 크기 :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "매끄러움 :"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmoil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ScrSmooth_Change()
On Error Resume Next
If frmmain.actark = 0 Then
    frmmain.WaterMark(frmmain.actark).ScaleMode = 3
    PicOilPaint frmmain.PicBack, 0, 0, frmmain.PicBack.ScaleWidth - 1, frmmain.PicBack.ScaleHeight - 1, ScrBr.Value, ScrSmooth.Value
    frmmain.WaterMark(frmmain.actark).ScaleMode = 1
Else
    frmmain.WaterMark(frmmain.actark).ScaleMode = 3
    PicOilPaint frmmain.WaterMark(frmmain.actark), 0, 0, frmmain.WaterMark(frmmain.actark).ScaleWidth - 1, frmmain.WaterMark(frmmain.actark).ScaleHeight - 1, ScrBr.Value, ScrSmooth.Value
    frmmain.WaterMark(frmmain.actark).ScaleMode = 1
End If
End Sub
