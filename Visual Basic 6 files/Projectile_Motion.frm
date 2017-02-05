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
   Begin VB.Line Line2 
      X1              =   7320
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
Line (2000, 2000)-(2025, 2025), vbBlack, B
End Sub
