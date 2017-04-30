VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame dialogFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   8880
      TabIndex        =   24
      Top             =   3120
      Width           =   7695
      Begin VB.TextBox Dialog1Box 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3120
         TabIndex        =   36
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Dialog2Box 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3120
         TabIndex        =   35
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton dialogButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   4320
         TabIndex        =   34
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Dialog3Box 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3120
         TabIndex        =   33
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame optionFrame 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3375
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton Option3 
            BackColor       =   &H00000000&
            Caption         =   "Option3"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   240
            TabIndex        =   32
            Top             =   2280
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "Range (metres)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Initial velocity (m/s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame labelFrame 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3495
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label heightSDialogLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Height at projectile launch (metres)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label heightEDialogLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Height at projectile landing (metres)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   0
            TabIndex        =   27
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label accelDialogLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Gravitational acceleration (ms^-2)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1095
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   2415
         End
      End
      Begin VB.Label dialogLabel 
         BackColor       =   &H00000000&
         Caption         =   "Select either initial velocity or range and enter in the box the corresponding variable."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   600
         TabIndex        =   37
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Frame mainFrame 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   11295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   23055
      Begin VB.PictureBox Image1 
         AutoSize        =   -1  'True
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4635
         ScaleWidth      =   11715
         TabIndex        =   12
         Top             =   240
         Width           =   11775
      End
      Begin VB.TextBox timeBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox rangeBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox initVeloBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox xVelocityBox 
         Height          =   615
         Left            =   18240
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox yVelocityBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   7
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox timeSpecificBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   6
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox angleBox 
         Height          =   615
         Left            =   18240
         TabIndex        =   5
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox heightBox 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   18240
         TabIndex        =   4
         Text            =   "0"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox heightEndBox 
         Height          =   495
         Left            =   18240
         TabIndex        =   3
         Text            =   "0"
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox maxHeightBox 
         Height          =   375
         Left            =   18240
         TabIndex        =   2
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox accelBox 
         Height          =   375
         Left            =   18240
         TabIndex        =   1
         Text            =   "9.8"
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label timeLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   16560
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label rangeLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   16560
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label velocityLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Initial Velocity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   16560
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label xVelocityLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "X Component of Initial Velocity (Ux)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   16560
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label yVelocityLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Y Component of Initial Velocity (Uy)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   16560
         TabIndex        =   19
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label timeSpecificLabel 
         BackColor       =   &H0000FF00&
         Caption         =   "Time at Max height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   16560
         TabIndex        =   18
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label angleLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Angle above Horizontal which Projectile was fired (in degrees)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   16560
         TabIndex        =   17
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label heightLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Height at projectile launch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   16560
         TabIndex        =   16
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label heightEndLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Height at projectile landing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   16560
         TabIndex        =   15
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label maxHeightLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Maximum Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   16560
         TabIndex        =   14
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label accelLabel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Acceleration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   16560
         TabIndex        =   13
         Top             =   6480
         Width           =   1575
      End
   End
   Begin VB.Label loadingLabel 
      BackColor       =   &H00000000&
      Caption         =   "Please Wait - Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9120
      TabIndex        =   38
      Top             =   3480
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mainFunc()
Dim initVelo As String ' Initial velocity of the projectile
Dim angle As String ' Angle between projectile's launch and horizontal
Dim range As String ' Horizontal distance projectile travels
Dim time As String ' Time of projectile's journey
Dim maxHeight As String ' Maximum height that projectile reaches
Dim timeSpecific As String
Dim xVelocity As String
Dim yVelocity As String

Dim holder

' Assigns information that user enters to variables
initVelo = initVeloBox.Text
angle = angleBox.Text
range = rangeBox.Text
time = timeBox.Text
maxHeight = maxHeightBox.Text
accel = accelBox.Text

If accel = 0 Then ' If this is not in code, program will break, and so this generates a straight line on the graph.
    holder = excelGraph(1000, 0, 0)
' Finds which information user has entered, if equals 0 then user has not entered information regarding that variable
ElseIf range <> "" Then ' All projectile motion questions either give range or intial velocity
    If time <> "" Then
        Call Algorithm3 ' Decides which algorithm should be used based on data given by user
    ElseIf angle <> "" Then
        Call Algorithm5
    ElseIf maxHeight <> "" Then
        
        Call Algorithm4
    Else
        MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
    End If
ElseIf initVelo <> "" Then
    If angle <> "" Then
        Call Algorithm1
    ElseIf time <> "" Then
        Call Algorithm2
    Else
        MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
    End If
Else
    MsgBox ("Cannot solve projectile motion problem. Please enter more variables.")
End If


End Sub

Private Sub Algorithm1()
Dim timeSpecific As Single ' time taken to reach maximum height
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single ' horizontal component of INITIAL velocity of projectile
Dim yVelocity As Single ' vertical component of INITIAL velocity of projectile
Dim height As Single
Dim timeSpecific2 As Single ' time taken between reaching maximum height and landing of projectile
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim accel As Single
Dim angleR As Single
Dim holder ' Due to Visual Baic 6 limitations, an empty variable is required to call functions

height = heightBox.Text ' Getting variables from user
initVelo = initVeloBox.Text
angle = angleBox.Text
accel = accelBox.Text

angleR = (angle / 180) * 3.14159265358979
xVelocity = initVelo * Math.Cos(angleR) ' The Math.Cos and Math.Sin functions require the angle to be in radians
yVelocity = initVelo * Math.Sin(angleR) ' Uses Sin and Cos to find horizontal and vertical components of initial velocity
maxHeight = (yVelocity ^ 2 / (2 * accel)) + height ' rearranges v^2 = u^2 + 2as to find maxHeight (v is 0 at maxHeight)
timeSpecific = yVelocity / accel ' as v is 0 at maxHeight, v = u + at becomes t = u / a
timeSpecific2 = (maxHeight / (0.5 * accel)) ^ (1 / 2) ' If only journey after maxHeight is considered, u in s = ut + 0.5at^2 is 0
time = timeSpecific + timeSpecific2 ' time before maxHeight is reached and time after, when added, becomes total time of journey
range = xVelocity * time ' One of equations of projectile motion
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight) ' outputs all variables to user
holder = excelGraph(time, yVelocity, height) ' Outputs some variables to Graph function to display graph
End Sub

Private Sub Algorithm2() ' include heightEnd and heightDiff
Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim angleR As Single
Dim xVelocity As Single
Dim yVelocity As Single
Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim divisor As Single ' temporary variable required to calculate inverse sin
Dim accel As Single
height = heightBox.Text
initVelo = initVeloBox.Text
time = timeBox.Text
accel = accelBox.Text

yVelocity = (((-1) * height) - (0.5 * -accel * time ^ 2)) / time ' Rearrange s = ut + 0.5at^2
divisor = yVelocity / initVelo ' inverse sin of yVelocity / initVelo will find angle
angleR = Math.Atn(divisor / (-divisor * divisor + 1) ^ 0.5) ' VB6 does not support inverse sin, so inverse tan required to calculate inverse sin
angle = (angleR * 180) / 3.1415926
xVelocity = initVelo * Math.Cos(angleR) ' As in previous algorithm
range = xVelocity * time
maxHeight = (yVelocity ^ 2 / (2 * accel)) + height
timeSpecific = yVelocity / accel
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Private Sub Algorithm3()

Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim angleR As Single
Dim xVelocity As Single
Dim yVelocity As Single

Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim divisor As Single
Dim heightEnd As Single
Dim heightDiff As Single
Dim accel As Single

range = rangeBox.Text
time = timeBox.Text
height = heightBox.Text
heightEnd = heightEndBox.Text
heightDiff = heightEnd - height
accel = accelBox.Text

yVelocity = (heightDiff - (0.5 * -accel * time ^ 2)) / time ' rearrange s = ut + 0.5at^2
xVelocity = range / time
angleR = Math.Atn(yVelocity / xVelocity) ' As x and y components of initial velocity are found, angle can be found using tan
initVelo = ((xVelocity ^ 2) + (yVelocity ^ 2)) ^ 0.5 ' Pythagoras theorem to find initVelo
maxHeight = (yVelocity ^ 2 / (2 * accel)) + height ' Same as previous algorithm
timeSpecific = yVelocity / accel
angle = (angleR * 180) / 3.1415926
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Private Sub Algorithm4()

Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim angleR As Single
Dim xVelocity As Single
Dim yVelocity As Single


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
accel = accelBox.Text

yVelocity = Math.Sqr(2 * accel * maxHeight) ' v^2 = u^2 + 2as except v = 0 at max height, so is u^2 = -2as rearranged
timeSpecific = yVelocity / accel
timeSpecific2 = Math.Sqr((height + maxHeight) / (0.5 * accel)) ' s = ut + 0.5at^2 except u = 0 at maxHeight
time = timeSpecific + timeSpecific2 ' Below is same as previous algorithms
xVelocity = range / time
angleR = Math.Atn(yVelocity / xVelocity)
angle = (angleR * 180) / 3.1415926
initVelo = ((xVelocity ^ 2) + (yVelocity ^ 2)) ^ 0.5
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Private Sub Algorithm5()
Dim timeSpecific As Single
Dim initVelo As Single
Dim angle As Single
Dim xVelocity As Single
Dim yVelocity As Single


Dim height As Single
Dim timeSpecific2 As Single
Dim maxHeight As Single
Dim time As Single
Dim range As Single
Dim divisor As Single
Dim heightEnd As Single
Dim heightDiff As Single
Dim angleR As Single
Dim temp As Single

range = rangeBox.Text
height = heightBox.Text
angle = angleBox.Text
accel = accelBox.Text
heightEnd = heightEndBox.Text
heightDiff = heightEnd - height

angleR = (angle / 180) * 3.14159265358979
time = Math.Sqr((heightDiff - ((range * Math.Sin(angleR)) / Math.Cos(angleR))) / (0.5 * -accel))  ' range = xVelocity * time, so range can be substituted into s = ut + 0.5at^2
initVelo = range / (time * Math.Cos(angleR)) ' time is substituted into range = xVelocity * time and rearranged to find value for initVelo
xVelocity = initVelo * Math.Cos(angleR) ' Below is same as previous algorithms
yVelocity = initVelo * Math.Sin(angleR)
maxHeight = (yVelocity ^ 2 / (2 * accel)) + height
timeSpecific = yVelocity / accel
holder = OutputFunc(time, range, initVelo, xVelocity, yVelocity, timeSpecific, angle, maxHeight)
holder = excelGraph(time, yVelocity, height)
End Sub

Function OutputFunc(time As Single, range As Single, initVelo As Single, xVelocity As Single, yVelocity As Single, timeSpecific As Single, angle As Single, maxHeight As Single)
timeBox.Text = time ' Sets all boxes to their appropriate values
rangeBox.Text = range
initVeloBox.Text = initVelo
xVelocityBox.Text = xVelocity
yVelocityBox.Text = yVelocity
timeSpecificBox.Text = timeSpecific
angleBox.Text = angle
maxHeightBox.Text = maxHeight
End Function

Function excelGraph(time As Single, yVelocity As Single, height As Single)
Dim xlApp As excel.Application ' Below are required variables needed for working with Excel
Set xlApp = New excel.Application ' Instance of Excel application created and set
Dim xlWkb As excel.Workbook ' Instance of an Excel workbook created and set
Set xlWkb = xlApp.Workbooks.Open("D:\Documents\Book1.xlsx")
Dim xlSht As excel.Worksheet
Set xlSht = xlWkb.Worksheets(1) ' Instance of an Excel worksheet within the workbook created and set


Dim timeInterval As Single
Dim i As Integer
Dim times(9) As Single
Dim heights(9) As Single
timeInterval = time / 10 ' finds amount of time between each time interval (required to set points on the graph)
For i = 0 To 9
    times(i) = timeInterval * (i + 1) ' create 10 points of time with interval between each point
    heights(i) = (yVelocity * times(i)) + (0.5 * -(accelBox.Text) * times(i) ^ 2) ' uses s = ut + 0.5at^2 to find height at each point of time
Next
xlSht.Cells(1, 1).Value = "Time" ' Adding data to the excel workbook including all times and heights
xlSht.Cells(1, 2).Value = "Height"
xlSht.Cells(2, 1).Value = 0
xlSht.Cells(2, 2).Value = height
For i = 3 To 12
    xlSht.Cells(i, 1).Value = times(i - 3)
    xlSht.Cells(i, 2).Value = heights(i - 3) + height
Next

Dim xlChart As excel.Chart ' Creates excel chart
Set xlChart = xlWkb.Charts.Add
xlChart.ChartType = xlLine
xlChart.SetSourceData xlSht.range("A1:B12"), xlColumns ' Read data from workbook
xlChart.Visible = xlSheetVisible
xlChart.Legend.Clear
xlChart.ChartArea.Font.Size = 10 ' Size of font on graph
xlChart.ChartArea.Font.Color = vbRed ' Color of font on graph
For i = 1 To xlChart.FullSeriesCollection.Count
    xlChart.FullSeriesCollection(i).Smooth = True ' Makes line connecting points on graph curved
Next

xlChart.ChartArea.Select
xlChart.ChartArea.Copy ' Copies excel chart
Image1.Picture = Clipboard.GetData(vbCFBitmap) ' Reads clipboard and displays picture in clipboard which is the excel chart
End Function

Private Sub dialogButton_Click()
If dialogLabel.Caption = "Select either initial velocity or range and enter in the box the corresponding variable." Then
    If Option1.Value = True Then
        If IsNumeric(Dialog1Box.Text) = False Then
            MsgBox ("Please enter a positive number.")
        Else
            initVeloBox.Text = Dialog1Box.Text
            dialogLabel.Caption = "Select either time or angle and enter in the box the corresponding variable."
            Option1.Caption = "Angle (degrees)"
            Option2.Caption = "Time (seconds)"
            Dialog1Box.Text = ""
        End If
    ElseIf Option2.Value = True Then
        If IsNumeric(Dialog2Box.Text) = False Then
            MsgBox ("Please enter a positive number.")
        Else
            rangeBox.Text = Dialog2Box.Text
            dialogLabel.Caption = "Select either time, angle or maximum height and enter in the box the corresponding variable."
            Option3.Enabled = True
            Option3.Visible = True
            Dialog3Box.Visible = True
            Option1.Caption = "Angle (degrees)"
            Option2.Caption = "Time (seconds)"
            Option3.Caption = "Maximum Height (metres)"
            Dialog2Box.Text = ""
        End If
    Else
        MsgBox ("Error - No option was selected.")
    End If
ElseIf dialogLabel.Caption = "Select either time or angle and enter in the box the corresponding variable." Then
    If Option1.Value = True Then
        If IsNumeric(Dialog1Box.Text) = False Then
            MsgBox ("Please enter a positive number.")
        Else
            angleBox.Text = Dialog1Box.Text
            Dialog1Box.Text = ""
        End If
    ElseIf Option2.Value = True Then
        If IsNumeric(Dialog2Box.Text) = False Then
            MsgBox ("Please enter a positive number")
        Else
            timeBox.Text = Dialog2Box.Text
            Dialog2Box.Text = ""
        End If
    Else
        MsgBox ("Error - No option was selected")
    End If
    optionFrame.Visible = False
    labelFrame.Visible = True
    Dialog1Box.Enabled = True
    Dialog2Box.Enabled = True
    Dialog3Box.Enabled = True
    Option3.Visible = True
    Dialog3Box.Visible = True
    dialogLabel.Caption = "Enter in the box the heights at projectile launch and landing and the gravitational acceleration."
    Dialog1Box.Text = "0"
    Dialog2Box.Text = "0"
    Dialog3Box.Text = "9.8"
ElseIf dialogLabel.Caption = "Select either time, angle or maximum height and enter in the box the corresponding variable." Then
    If Option1.Value = True Then
        If IsNumeric(Dialog1Box.Text) = False Then
            MsgBox ("Please enter a positive number")
        Else
            angleBox.Text = Dialog1Box.Text
            Dialog1Box.Text = ""
        End If
    ElseIf Option2.Value = True Then
        If IsNumeric(Dialog2Box.Text) = False Then
            MsgBox ("Please enter a positive number")
        Else
            timeBox.Text = Dialog2Box.Text
            Dialog2Box.Text = ""
        End If
    ElseIf Option3.Value = True Then
        If IsNumeric(Dialog3Box.Text) = False Then
            MsgBox ("Please enter a positive number")
        Else
            maxHeightBox.Text = Dialog3Box.Text
            Dialog3Box.Text = ""
        End If
    Else
        MsgBox ("Error - No option was selected.")
    End If
    optionFrame.Visible = False
    labelFrame.Visible = True
    Dialog1Box.Enabled = True
    Dialog2Box.Enabled = True
    Dialog3Box.Enabled = True
    dialogLabel.Caption = "Enter in the box the heights at projectile launch and landing and the gravitational acceleration."
    Dialog1Box.Text = "0"
    Dialog2Box.Text = "0"
    Dialog3Box.Text = "9.8"
ElseIf dialogLabel.Caption = "Enter in the box the heights at projectile launch and landing and the gravitational acceleration." Then
    If IsNumeric(Dialog1Box.Text) = False Or IsNumeric(Dialog2Box.Text) = False Or IsNumeric(Dialog3Box.Text) = False Then
        MsgBox ("Please enter a positive number")
    Else
        heightBox.Text = Dialog1Box.Text
        heightEndBox.Text = Dialog2Box.Text
        accelBox.Text = Dialog3Box.Text
        dialogFrame.Visible = False
        Call mainFunc
        mainFrame.Visible = True
    End If
Else
    MsgBox ("Error - dialogLabel has been edited.")
End If
    
End Sub

Private Sub Option1_Click()
Dialog1Box.Enabled = True
Dialog2Box.Enabled = False
Dialog3Box.Enabled = False
End Sub

Private Sub Option2_Click()
Dialog1Box.Enabled = False
Dialog2Box.Enabled = True
Dialog3Box.Enabled = False
End Sub

Private Sub Option3_Click()
Dialog1Box.Enabled = False
Dialog2Box.Enabled = False
Dialog3Box.Enabled = True
End Sub
