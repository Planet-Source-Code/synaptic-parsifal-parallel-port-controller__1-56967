VERSION 5.00
Begin VB.Form Alarm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alarm Clock"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAlarm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alarm On/Off"
      Height          =   195
      Left            =   6750
      TabIndex        =   30
      Top             =   2205
      Width           =   1410
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3270
      TabIndex        =   35
      Top             =   2445
      Width           =   495
   End
   Begin VB.TextBox txtSec 
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3915
      TabIndex        =   34
      Top             =   2445
      Width           =   495
   End
   Begin VB.TextBox txtHour 
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   33
      Top             =   2445
      Width           =   495
   End
   Begin VB.Timer tmrAlarm01 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1320
      Top             =   3120
   End
   Begin VB.CheckBox chkPM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PM"
      Height          =   315
      Left            =   5220
      TabIndex        =   32
      Top             =   2430
      Width           =   645
   End
   Begin VB.CheckBox chkAM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AM"
      Height          =   240
      Left            =   4575
      TabIndex        =   31
      Top             =   2475
      Width           =   645
   End
   Begin VB.Timer tmrAlarm 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose wich pins to turn on when the alarm goes off"
      Height          =   1905
      Left            =   105
      TabIndex        =   2
      Top             =   60
      Width           =   8505
      Begin VB.CheckBox D01 
         BackColor       =   &H00FF0000&
         Caption         =   "02"
         Height          =   195
         Left            =   7005
         TabIndex        =   27
         Top             =   615
         Width           =   525
      End
      Begin VB.CheckBox D04 
         BackColor       =   &H00FF0000&
         Caption         =   "05"
         Height          =   195
         Left            =   5205
         TabIndex        =   26
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox D06 
         BackColor       =   &H00FF0000&
         Caption         =   "07"
         Height          =   195
         Left            =   4005
         TabIndex        =   25
         Top             =   615
         Width           =   555
      End
      Begin VB.CheckBox D08 
         BackColor       =   &H00FF0000&
         Caption         =   "09"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2805
         TabIndex        =   24
         Top             =   615
         Width           =   555
      End
      Begin VB.CheckBox S01 
         BackColor       =   &H00FF0000&
         Caption         =   "10"
         Enabled         =   0   'False
         Height          =   210
         Left            =   2205
         TabIndex        =   23
         Top             =   615
         Width           =   540
      End
      Begin VB.CheckBox S03 
         BackColor       =   &H00FF0000&
         Caption         =   "12"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1005
         TabIndex        =   22
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox D05 
         BackColor       =   &H00FF0000&
         Caption         =   "06"
         Height          =   195
         Left            =   4605
         TabIndex        =   21
         Top             =   615
         Width           =   540
      End
      Begin VB.CheckBox D03 
         BackColor       =   &H00FF0000&
         Caption         =   "04"
         Height          =   195
         Left            =   5805
         TabIndex        =   20
         Top             =   615
         Width           =   525
      End
      Begin VB.CheckBox C01 
         BackColor       =   &H00FF0000&
         Caption         =   "01"
         Enabled         =   0   'False
         Height          =   210
         Left            =   7605
         TabIndex        =   19
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox D07 
         BackColor       =   &H00FF0000&
         Caption         =   "08"
         Height          =   195
         Left            =   3405
         TabIndex        =   18
         Top             =   615
         Width           =   555
      End
      Begin VB.CheckBox S02 
         BackColor       =   &H00FF0000&
         Caption         =   "11"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1605
         TabIndex        =   17
         Top             =   615
         Width           =   540
      End
      Begin VB.CheckBox D02 
         BackColor       =   &H00FF0000&
         Caption         =   "03"
         Height          =   195
         Left            =   6405
         TabIndex        =   16
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox S04 
         BackColor       =   &H00FF0000&
         Caption         =   "13"
         Enabled         =   0   'False
         Height          =   195
         Left            =   405
         TabIndex        =   15
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox G07 
         BackColor       =   &H00FF0000&
         Caption         =   "24"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1245
         TabIndex        =   14
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G05 
         BackColor       =   &H00FF0000&
         Caption         =   "22"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2445
         TabIndex        =   13
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G04 
         BackColor       =   &H00FF0000&
         Caption         =   "21"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3045
         TabIndex        =   12
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G02 
         BackColor       =   &H00FF0000&
         Caption         =   "19"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4245
         TabIndex        =   11
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G01 
         BackColor       =   &H00FF0000&
         Caption         =   "18"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4845
         TabIndex        =   10
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox C02 
         BackColor       =   &H00FF0000&
         Caption         =   "14"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7245
         TabIndex        =   9
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G08 
         BackColor       =   &H00FF0000&
         Caption         =   "25"
         Enabled         =   0   'False
         Height          =   195
         Left            =   645
         TabIndex        =   8
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G06 
         BackColor       =   &H00FF0000&
         Caption         =   "23"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1845
         TabIndex        =   7
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox G03 
         BackColor       =   &H00FF0000&
         Caption         =   "20"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3645
         TabIndex        =   6
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox C04 
         BackColor       =   &H00FF0000&
         Caption         =   "17"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5445
         TabIndex        =   5
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox C03 
         BackColor       =   &H00FF0000&
         Caption         =   "16"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6045
         TabIndex        =   4
         Top             =   975
         Width           =   495
      End
      Begin VB.CheckBox S05 
         BackColor       =   &H00FF0000&
         Caption         =   "15"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6645
         TabIndex        =   3
         Top             =   975
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   510
         X2              =   510
         Y1              =   960
         Y2              =   1290
      End
      Begin VB.Line Line2 
         X1              =   5340
         X2              =   5340
         Y1              =   975
         Y2              =   1275
      End
      Begin VB.Line Line3 
         X1              =   510
         X2              =   5355
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Ground"
         Height          =   210
         Left            =   2355
         TabIndex        =   28
         Top             =   1275
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   1590
         Left            =   75
         Picture         =   "Alarm.frx":0000
         Stretch         =   -1  'True
         Top             =   255
         Width           =   8400
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   135
      Top             =   3120
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HH          MM          SS"
      Height          =   270
      Left            =   2610
      TabIndex        =   40
      Top             =   2835
      Width           =   1800
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is the time that the alarm is set for"
      Height          =   510
      Left            =   6405
      TabIndex        =   39
      Top             =   2835
      Width           =   2055
   End
   Begin VB.Label lblAlarm 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6390
      TabIndex        =   38
      Top             =   2505
      Width           =   2070
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3780
      TabIndex        =   37
      Top             =   2475
      Width           =   150
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3135
      TabIndex        =   36
      Top             =   2475
      Width           =   150
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm"
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2625
      TabIndex        =   29
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label lblTimeCap 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   2100
      Width           =   645
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "WST_Germ"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   2430
      Width           =   2085
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name - Parallel Port Controller
'Author - Josh Gaby
'Date - 28/10/2004

Option Explicit

'Declare variables
Dim D As Integer


Public Function Data_Port()
'Convert binary into integer
'                           eg:
'                                0  |  0  |  0  |  1  |  1  |  0  |  0  |  1
'                               128 | 064 | 032 | 016 | 008 | 004 | 002 | 001
'                                                 016 + 008       +       001  = 25
If D08.Value = 1 Then D = D + 128
If D07.Value = 1 Then D = D + 64
If D06.Value = 1 Then D = D + 32
If D05.Value = 1 Then D = D + 16
If D04.Value = 1 Then D = D + 8
If D03.Value = 1 Then D = D + 4
If D02.Value = 1 Then D = D + 2
If D01.Value = 1 Then D = D + 1

End Function

Private Sub D01_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D02_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D03_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D04_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D05_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D06_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D07_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub D08_Click()
'Call Public Function Data_Port
Call Data_Port
End Sub

Private Sub tmrTime_Timer()

'Display the Time in lblTime's Caption
lblTime.Caption = Time

'Check if chkAlarm is turned on (checked)
If chkAlarm.Value = 1 Then
    
    'If it is then turn on the two alrm timers (tmrAlarm and tmrAlarm01)
    tmrAlarm.Enabled = True
    tmrAlarm01.Enabled = True
    
ElseIf chkAlarm.Value = 0 Then
    
    'if it is not then turn of the two alarm timers 9tmrAlarm anf tmrAlarm01)
    tmrAlarm.Enabled = False
    tmrAlarm01.Enabled = False
    
End If
End Sub

Private Sub tmrAlarm_Timer()

'Declare variables
Dim A As String

'Check if chkAM is turned on (checked)
If chkAM.Value = 1 Then
    
    'If it is then A = a.m.
    A = "a.m."
    
ElseIf chkAM.Value = 0 Then
    'If it is not then do nothing
    
'Check if chkPM is turned on (checked)
ElseIf chkPM.Value = 1 Then
    
    'If it is then A = p.m.
    A = "p.m."

ElseIf chkPM.Value = 0 Then
    'If it isnot then do nothing
    
End If

'Check if the time is the same as the alarm time that you set
If lblTime.Caption = txtHour.Text + ":" + txtMin.Text + ":" + txtSec.Text + " " + A Then
    
    'If it is then send the signal to the data port (&H378) to turn on the pins that you selected (D)
    Out Val("&H378"), Val(D)
    
Else
    'If it is not then do nothing

End If

'Display the alarm time that you set on the form in a seperate label (lblAlarm)
lblAlarm.Caption = txtHour.Text + ":" + txtMin.Text + ":" + txtSec.Text + " " + A

End Sub

Private Sub tmrAlarm01_Timer()

'Declare variables
Dim A As String

'Check if chkAM is turned on (checked)
If chkAM.Value = 1 Then
    
    'If it is then A = a.m.
    A = "a.m."
    
ElseIf chkAM.Value = 0 Then
    'If it is not then do nothing
    
'Check if chkPM is turned on (checked)
ElseIf chkPM.Value = 1 Then
    
    'If it is then A = p.m.
    A = "p.m."
    
ElseIf chkPM.Value = 0 Then
    'if it is not then do nothing
    
End If

'Check if the time is the same as the alarm time that you set
If lblTime.Caption = txtHour.Text + ":" + txtMin.Text + ":" + txtSec.Text + " " + A Then
    
    'If it is then send the signal to the data port (&H378) to turn on the pins that you selected (D)
    Out Val("&H378"), Val(D)
    
Else
    'If it is not then do nothing

End If

'Display the alarm time that you set on the form in a seperate label (lblAlarm)
lblAlarm.Caption = txtHour.Text + ":" + txtMin.Text + ":" + txtSec.Text + " " + A

End Sub
