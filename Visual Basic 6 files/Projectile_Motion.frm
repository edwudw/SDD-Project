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
   Begin VB.PictureBox Image1 
      Height          =   4695
      Left            =   3960
      ScaleHeight     =   4635
      ScaleWidth      =   11715
      TabIndex        =   2
      Top             =   960
      Width           =   11775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   11760
      TabIndex        =   1
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton button 
      Caption         =   "Start"
      Height          =   495
      Left            =   14160
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Line Line9 
      X1              =   1445
      X2              =   1575
      Y1              =   575
      Y2              =   575
   End
   Begin VB.Line Line8 
      X1              =   1445
      X2              =   1575
      Y1              =   1225
      Y2              =   1225
   End
   Begin VB.Line Line7 
      X1              =   1445
      X2              =   1575
      Y1              =   1850
      Y2              =   1850
   End
   Begin VB.Line Line6 
      X1              =   1445
      X2              =   1575
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line5 
      X1              =   1445
      X2              =   1575
      Y1              =   3100
      Y2              =   3100
   End
   Begin VB.Line Line4 
      X1              =   1445
      X2              =   1575
      Y1              =   3725
      Y2              =   3725
   End
   Begin VB.Line Line3 
      X1              =   1445
      X2              =   1575
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Line Line2 
      X1              =   1445
      X2              =   1575
      Y1              =   4975
      Y2              =   4975
   End
   Begin VB.Line Line1 
      X1              =   1445
      X2              =   1575
      Y1              =   5600
      Y2              =   5600
   End
   Begin VB.Line xLine0 
      X1              =   1560
      X2              =   1560
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine9 
      X1              =   7200
      X2              =   7200
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine8 
      X1              =   6575
      X2              =   6575
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine7 
      X1              =   5950
      X2              =   5950
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine6 
      X1              =   5325
      X2              =   5325
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine5 
      X1              =   4700
      X2              =   4700
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine4 
      X1              =   4075
      X2              =   4075
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine3 
      X1              =   3450
      X2              =   3450
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine2 
      X1              =   2825
      X2              =   2825
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xLine1 
      X1              =   2200
      X2              =   2200
      Y1              =   6250
      Y2              =   6380
   End
   Begin VB.Line xAxisLine 
      X1              =   7200
      X2              =   1560
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line yAxisLine 
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
Dim xlApp As excel.Application
Set xlApp = New excel.Application
Dim xlWkb As excel.Workbook
Set xlWkb = xlApp.Workbooks.Open("D:\Documents\Book1.xlsx")
Dim xlSht As excel.Worksheet
Set xlSht = xlWkb.Worksheets(1)
Dim xlChart As excel.Chart
Set xlChart = xlWkb.Charts.Add
xlChart.ChartType = xlLine
xlChart.SetSourceData xlSht.Range("A1:B5"), xlColumns
xlChart.Visible = xlSheetVisible
xlChart.Legend.Clear
xlChart.
xlChart.ChartArea.Font.Size = 15
xlChart.ChartArea.Font.Color = vbRed
xlChart.ChartArea.Select
xlChart.ChartArea.Copy
Image1.Picture = Clipboard.GetData(vbCFBitmap)
End Sub

Private Sub Form_Load()

End Sub
