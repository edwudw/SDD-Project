VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   17295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton button 
      Caption         =   "Start"
      Height          =   495
      Left            =   14160
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Line Line11 
      X1              =   7200
      X2              =   7200
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line10 
      X1              =   6575
      X2              =   6575
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line9 
      X1              =   5950
      X2              =   5950
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line8 
      X1              =   5325
      X2              =   5325
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line7 
      X1              =   4700
      X2              =   4700
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line6 
      X1              =   4075
      X2              =   4075
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line5 
      X1              =   3450
      X2              =   3450
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line4 
      X1              =   2825
      X2              =   2825
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line3 
      X1              =   2200
      X2              =   2200
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line Line2 
      X1              =   7200
      X2              =   1560
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   600
      Y2              =   6240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub button_Click()
FillColor = vbBlack
FillStyle = vbSolid
Line (1675, 6225)-(8000, 2025), vbBlack, B
End Sub

