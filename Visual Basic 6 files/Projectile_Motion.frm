VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox maxHeightBox 
      Height          =   375
      Left            =   18840
      TabIndex        =   22
      Text            =   "0"
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox heightEndBox 
      Height          =   615
      Left            =   18840
      TabIndex        =   20
      Text            =   "0"
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox heightBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   18
      Text            =   "0"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Output 
      Height          =   855
      Left            =   15120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox angleBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   15
      Text            =   "0"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox timeSpecificBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   13
      Text            =   "0"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox yVelocityBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   12
      Text            =   "0"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox xVelocityBox 
      Height          =   615
      Left            =   18840
      TabIndex        =   11
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox initVeloBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   10
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox rangeBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox timeBox 
      Height          =   495
      Left            =   18840
      TabIndex        =   8
      Text            =   "0"
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
   Begin VB.Label maxHeightLabel 
      Caption         =   "Maximum Height"
      Height          =   615
      Left            =   17160
      TabIndex        =   21
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label heightEndLabel 
      Caption         =   "Height at projectile landing"
      Height          =   495
      Left            =   17160
      TabIndex        =   19
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label heightLabel 
      Caption         =   "Height at projectile launch"
      Height          =   615
      Left            =   17160
      TabIndex        =   17
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label angleLabel 
      Caption         =   "Angle above Horizontal which Projectile was fired"
      Height          =   615
      Left            =   17160
      TabIndex        =   14
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label timeSpecificLabel 
      Caption         =   "Time at Max height"
      Height          =   495
      Left            =   17160
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
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
Dim initVelo As Single
Dim angle As Single
Dim range As Single
Dim time As Single
Dim maxHeight As Single
initVelo = initVeloBox.Text
angle = angleBox.Text
range = rangeBox.Text
time = timeBox.Text
maxHeight = maxHeightBox.Text

If range <> 0 Then
    If time <> 0 Then
        Call Algorithms4
    ElseIf angle <> 0 Then
        Call Algorithms6
    ElseIf maxHeight <> 0 Then
        
        Call Algorithms5
    Else
        MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
    End If
ElseIf initVelo <> 0 Then
    If angle <> 0 Then
        Call Algorithms2
    ElseIf time <> 0 Then
        Call Algorithms3
    Else
        MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
    End If
Else
    MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
End If


End Sub

' Private Sub Algorithms()
' Dim timeSpecific As Single
' Dim initVelo As Single
' Dim angle As Single
' Dim xVelocity As Single
' Dim yVelocity As Single
' Dim ySpecificVelocity As Single
' Dim overallVelocity As Single

' timeSpecific = timeSpecificBox.Text ' CHECK THIS LATER
' initVelo = initVeloBox.Text
' angle = angleBox.Text

' xVelocity = initVelo * Math.Cos((angle / 180) * 3.14159265358979)
' yVelocity = initVelo * Math.Sin((angle / 180) * 3.14159265358979)

' ySpecificVelocity = yVelocity + (-9.8 * timeSpecific)
' overallVelocity = ((ySpecificVelocity) ^ 2 + (xVelocity) ^ 2) ^ (1 / 2)
' Output.Text = CStr(overallVelocity)
' End Sub

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
Dim holder

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
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
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
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
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
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Private Sub Algorithms5()

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
maxHeight = maxHeightBox.Text
height = heightBox.Text

yVelocity = Math.Sqr(2 * 9.8 * maxHeight)
timeSpecific = yVelocity / 9.8
timeSpecific2 = Math.Sqr((height + maxHeight) / (0.5 * 9.8))
time = timeSpecific + timeSpecific2
xVelocity = range / time
angle = Math.Atn(yVelocity / xVelocity)
initVelo = ((xVelocity ^ 2) + (yVelocity ^ 2)) ^ 0.5
Output.Text = CStr(xVelocity)
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Private Sub Algorithms6()
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
Dim angleR As Single
Dim test As Single

range = rangeBox.Text
height = heightBox.Text
angle = angleBox.Text
heightDiff = heightEnd - height
angleR = (angle / 180) * 3.14159265358979
test = Math.Sqr((heightDiff - ((300 * Math.Sin(angleR)) / Math.Cos(angleR))) / (0.5 * -9.8)) ' CHANGE NAME OF THIS VARIABLE
initVelo = range / (test * Math.Cos(angleR))
xVelocity = initVelo * Math.Cos(angleR)
yVelocity = initVelo * Math.Sin(angleR)
time = range / xVelocity
maxHeight = (yVelocity ^ 2 / (2 * 9.8)) + height
timeSpecific = yVelocity / 9.8
Output.Text = CStr(maxHeight)
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Function OutputFunc(time As Single, range As Single, initVelo As Single, xVelocity As Single, yVelocity As Single, timeSpecific As Single, angle As Single, maxHeight As Single)
timeBox.Text = time
rangeBox.Text = range
initVeloBox.Text = initVelo
xVelocityBox.Text = xVelocity
yVelocityBox.Text = yVelocity
timeSpecificBox.Text = timeSpecific
angleBox.Text = angle
maxHeightBox.Text = maxHeight
End Function

Function excelGraph(time As Single, yVelocity As Single, height As Single)
Dim xlApp As excel.Application
Set xlApp = New excel.Application
Dim xlWkb As excel.Workbook
Set xlWkb = xlApp.Workbooks.Open("D:\Documents\Book1.xlsx")
Dim xlSht As excel.Worksheet
Set xlSht = xlWkb.Worksheets(1)


Dim timeInterval As Single
Dim i As Integer
Dim times(9) As Single
Dim heights(9) As Single
timeInterval = time / 10
For i = 0 To 9
    times(i) = timeInterval * (i + 1)
    heights(i) = (yVelocity * times(i)) + (0.5 * -9.8 * times(i) ^ 2)
Next
xlSht.Cells(1, 1).Value = "Time"
xlSht.Cells(1, 2).Value = "Height"
xlSht.Cells(2, 1).Value = 0
xlSht.Cells(2, 2).Value = height
Output.Text = times(0)
For i = 3 To 12
    xlSht.Cells(i, 1).Value = times(i - 3)
    xlSht.Cells(i, 2).Value = heights(i - 3) + height
Next
Dim xlChart As excel.Chart
Set xlChart = xlWkb.Charts.Add
xlChart.ChartType = xlLine
xlChart.SetSourceData xlSht.range("A1:B12"), xlColumns
xlChart.Visible = xlSheetVisible
xlChart.Legend.Clear
xlChart.ChartArea.Font.Size = 10
xlChart.ChartArea.Font.Color = vbRed
For i = 1 To xlChart.FullSeriesCollection.Count
    xlChart.FullSeriesCollection(i).Smooth = True
Next

xlChart.ChartArea.Select
xlChart.ChartArea.Copy
Image1.Picture = Clipboard.GetData(vbCFBitmap)
End Function

