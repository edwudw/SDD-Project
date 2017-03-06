VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20535
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   20535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   735
      Left            =   17760
      TabIndex        =   1
      Top             =   9960
      Width           =   1815
   End
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      Height          =   4695
      Left            =   600
      ScaleHeight     =   4635
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   240
      Width           =   11775
   End
   Begin VB.Label yVelocityLabel 
      Caption         =   "Y Component of Initial Velocity (Uy)"
      Height          =   495
      Left            =   17160
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label xVelocityLabel 
      Caption         =   "X Component of Initial Velocity (Ux)"
      Height          =   495
      Left            =   17160
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label velocityLabel 
      Caption         =   "Initial Velocity"
      Height          =   495
      Left            =   17160
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label rangeLabel 
      Caption         =   "Range"
      Height          =   375
      Left            =   17160
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label timeLabel 
      Caption         =   "Time"
      Height          =   255
      Left            =   17160
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
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
xlChart.ChartArea.Font.Size = 10
xlChart.ChartArea.Font.Color = vbRed
Dim i As Long
For i = 1 To xlChart.FullSeriesCollection.Count
    xlChart.FullSeriesCollection(i).Smooth = True
Next

xlChart.ChartArea.Select
xlChart.ChartArea.Copy
Image1.Picture = Clipboard.GetData(vbCFBitmap)
End Sub
