VERSION 5.00
Begin VB.Form frmRotate 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "회전"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   180
      TabIndex        =   0
      Top             =   120
      Value           =   45
      Width           =   4455
   End
End
Attribute VB_Name = "frmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
If Not frmmain.actark = 0 Then
    frmmain.Backup(frmmain.actark).ScaleMode = 3
    frmmain.WaterMark(frmmain.actark).ScaleMode = 3
    BitRotate frmmain.Backup(frmmain.actark), frmmain.WaterMark(frmmain.actark), HScroll1.Value, 1
    frmmain.WaterMark(frmmain.actark).ScaleMode = 1
    frmmain.Backup(frmmain.actark).ScaleMode = 1
End If
End Sub
