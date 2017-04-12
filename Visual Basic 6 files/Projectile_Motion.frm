VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   10875
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox heightEndBox 
      Height          =   615
      Left            =   18840
      TabIndex        =   22
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox heightBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   20
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Output 
      Height          =   855
      Left            =   15120
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox angleBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   17
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox timeSpecificBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   15
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox ySpecificVeloBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   14
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox yVelocityBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox xVelocityBox 
      Height          =   615
      Left            =   18840
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox initVeloBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox rangeBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox timeBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
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
   Begin VB.Label heightEndLabel 
      Caption         =   "Height at projectile landing"
      Height          =   495
      Left            =   17160
      TabIndex        =   21
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label heightLabel 
      Caption         =   "Height at projectile launch"
      Height          =   615
      Left            =   17160
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label angleLabel 
      Caption         =   "Angle above Horizontal which Projectile was fired"
      Height          =   615
      Left            =   17160
      TabIndex        =   16
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label timeSpecificLabel 
      Caption         =   "Time at Velocity Above"
      Height          =   495
      Left            =   17160
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label ySpecificVelocityLabel 
      Caption         =   "Y Comoponent of velocity at time below (Vy)"
      Height          =   615
      Left            =   17160
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
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

Call Algorithms4
Dim xlApp As excel.Application
Set xlApp = New excel.Application
Dim xlWkb As excel.Workbook
Set xlWkb = xlApp.Workbooks.Open("D:\Documents\Book1.xlsx")
Dim xlSht As excel.Worksheet
Set xlSht = xlWkb.Worksheets(1)
Dim xlChart As excel.Chart
Set xlChart = xlWkb.Charts.Add
xlChart.ChartType = xlLine
xlChart.SetSourceData xlSht.range("A1:B5"), xlColumns
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

Private Sub Algorithms()
Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single
Dim yVelocity As Single
Dim ySpecificVelocity As Single
Dim overallVelocity As Single

timeSpecific = timeSpecificBox.Text ' CHECK THIS LATER
initVelo = initVeloBox.Text
angle = angleBox.Text

xVelocity = initVelo * Math.Cos((angle / 180) * 3.14159265358979)
yVelocity = initVelo * Math.Sin((angle / 180) * 3.14159265358979)

ySpecificVelocity = yVelocity + (-9.8 * timeSpecific)
overallVelocity = ((ySpecificVelocity) ^ 2 + (xVelocity) ^ 2) ^ (1 / 2)
Output.Text = CStr(overallVelocity)
End Sub

Private Sub Algorithms2()
Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single
Dim yVelocity As Single
Dim ySpecificVelocity As Single
Dim overallVelocity As Single
Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single

height = heightBox.Text
initVelo = initVeloBox.Text
angle = angleBox.Text
xVelocity = initVelo * Math.Cos((angle / 180) * 3.14159265358979)
yVelocity = initVelo * Math.Sin((angle / 180) * 3.14159265358979)
maxHeight = (yVelocity ^ 2 / (2 * 9.8)) + height
timeSpecific = yVelocity / 9.8
timeSpecific2 = (maxHeight / (0.5 * 9.8)) ^ (1 / 2)
time = timeSpecific + timeSpecific2
range = xVelocity * time
Output.Text = CStr(range)
End Sub

Private Sub Algorithms3()
Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single
Dim yVelocity As Single
Dim ySpecificVelocity As Single
Dim overallVelocity As Single
Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim divisor As Single

height = heightBox.Text
initVelo = initVeloBox.Text
time = timeBox.Text

yVelocity = (((-1) * height) - (0.5 * -9.8 * time ^ 2)) / time
divisor = yVelocity / initVelo
angle = Math.Atn(divisor / (-divisor * divisor + 1) ^ 0.5)
xVelocity = initVelo * Math.Cos(angle)
range = xVelocity * time
maxHeight = (yVelocity ^ 2 / (2 * 9.8)) + height
timeSpecific = yVelocity / 9.8
Output.Text = CStr(maxHeight)
End Sub

Private Sub Algorithms4() ' FOR THIS ALGORITHM CHECK HEIGHT DIFFERENCES

Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single
Dim yVelocity As Single
Dim ySpecificVelocity As Single
Dim overallVelocity As Single
Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim divisor As Single
Dim heightEnd As Single
Dim heightDiff As Single

range = rangeBox.Text
time = timeBox.Text
height = heightBox.Text
heightEnd = heightEndBox.Text
heightDiff = heightEnd - height

yVelocity = (heightDiff - (0.5 * -9.8 * time ^ 2)) / time
xVelocity = range / time
angle = Math.Atn(yVelocity / xVelocity)
initVelo = ((xVelocity ^ 2) + (yVelocity ^ 2)) ^ 0.5
maxHeight = (yVelocity ^ 2 / (2 * 9.8)) + height
timeSpecific = yVelocity / 9.8
Output.Text = CStr(timeSpecific)
End Sub

Private Sub Algorithms5()

End Sub
