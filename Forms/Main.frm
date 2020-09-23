VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Parallel Port Control - By Josh Gaby"
   ClientHeight    =   10560
   ClientLeft      =   255
   ClientTop       =   615
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   8910
   Begin VB.Timer tmrResis 
      Interval        =   100
      Left            =   7950
      Top             =   3945
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control Port"
      Height          =   1695
      Left            =   5640
      TabIndex        =   35
      Top             =   1560
      Width           =   2295
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Integer - From Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Binary - From Check Box's"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCb 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblCi 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status Ports"
      Height          =   1695
      Left            =   3240
      TabIndex        =   30
      Top             =   1560
      Width           =   2295
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Integer - From Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Binary - From Check Box's"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblSi 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   105
         TabIndex        =   32
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblSb 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Ports"
      Height          =   1695
      Left            =   840
      TabIndex        =   25
      Top             =   1560
      Width           =   2295
      Begin VB.Label lblDb 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblDi 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   105
         TabIndex        =   28
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Binary - From Check Box's"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Integer -- From Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.CheckBox S05 
      BackColor       =   &H00FF0000&
      Caption         =   "15"
      Height          =   195
      Left            =   6825
      TabIndex        =   24
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox C03 
      BackColor       =   &H00FF0000&
      Caption         =   "16"
      Height          =   195
      Left            =   6225
      TabIndex        =   23
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox C04 
      BackColor       =   &H00FF0000&
      Caption         =   "17"
      Height          =   195
      Left            =   5625
      TabIndex        =   22
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G03 
      BackColor       =   &H00FF0000&
      Caption         =   "20"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3825
      TabIndex        =   21
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G06 
      BackColor       =   &H00FF0000&
      Caption         =   "23"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2025
      TabIndex        =   20
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G08 
      BackColor       =   &H00FF0000&
      Caption         =   "25"
      Enabled         =   0   'False
      Height          =   195
      Left            =   825
      TabIndex        =   19
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox C02 
      BackColor       =   &H00FF0000&
      Caption         =   "14"
      Height          =   195
      Left            =   7425
      TabIndex        =   18
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G01 
      BackColor       =   &H00FF0000&
      Caption         =   "18"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5025
      TabIndex        =   17
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G02 
      BackColor       =   &H00FF0000&
      Caption         =   "19"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4425
      TabIndex        =   16
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G04 
      BackColor       =   &H00FF0000&
      Caption         =   "21"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3225
      TabIndex        =   15
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G05 
      BackColor       =   &H00FF0000&
      Caption         =   "22"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2625
      TabIndex        =   14
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox G07 
      BackColor       =   &H00FF0000&
      Caption         =   "24"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1425
      TabIndex        =   13
      Top             =   975
      Width           =   495
   End
   Begin VB.CheckBox S04 
      BackColor       =   &H00FF0000&
      Caption         =   "13"
      Height          =   195
      Left            =   585
      TabIndex        =   12
      Top             =   615
      Width           =   495
   End
   Begin VB.CheckBox D02 
      BackColor       =   &H00FF0000&
      Caption         =   "03"
      Height          =   195
      Left            =   6585
      TabIndex        =   11
      Top             =   615
      Width           =   495
   End
   Begin VB.CheckBox S02 
      BackColor       =   &H00FF0000&
      Caption         =   "11"
      Height          =   195
      Left            =   1785
      TabIndex        =   10
      Top             =   615
      Width           =   540
   End
   Begin VB.CheckBox D07 
      BackColor       =   &H00FF0000&
      Caption         =   "08"
      Height          =   195
      Left            =   3585
      TabIndex        =   9
      Top             =   615
      Width           =   555
   End
   Begin VB.CheckBox C01 
      BackColor       =   &H00FF0000&
      Caption         =   "01"
      Height          =   210
      Left            =   7785
      TabIndex        =   8
      Top             =   615
      Width           =   495
   End
   Begin VB.CheckBox D03 
      BackColor       =   &H00FF0000&
      Caption         =   "04"
      Height          =   195
      Left            =   5985
      TabIndex        =   7
      Top             =   615
      Width           =   525
   End
   Begin VB.CheckBox D05 
      BackColor       =   &H00FF0000&
      Caption         =   "06"
      Height          =   195
      Left            =   4785
      TabIndex        =   6
      Top             =   615
      Width           =   540
   End
   Begin VB.CheckBox S03 
      BackColor       =   &H00FF0000&
      Caption         =   "12"
      Height          =   195
      Left            =   1185
      TabIndex        =   5
      Top             =   615
      Width           =   495
   End
   Begin VB.CheckBox S01 
      BackColor       =   &H00FF0000&
      Caption         =   "10"
      Height          =   210
      Left            =   2385
      TabIndex        =   4
      Top             =   615
      Width           =   540
   End
   Begin VB.CheckBox D08 
      BackColor       =   &H00FF0000&
      Caption         =   "09"
      Height          =   195
      Left            =   2985
      TabIndex        =   3
      Top             =   615
      Width           =   555
   End
   Begin VB.CheckBox D06 
      BackColor       =   &H00FF0000&
      Caption         =   "07"
      Height          =   195
      Left            =   4185
      TabIndex        =   2
      Top             =   615
      Width           =   555
   End
   Begin VB.CheckBox D04 
      BackColor       =   &H00FF0000&
      Caption         =   "05"
      Height          =   195
      Left            =   5385
      TabIndex        =   1
      Top             =   615
      Width           =   495
   End
   Begin VB.CheckBox D01 
      BackColor       =   &H00FF0000&
      Caption         =   "02"
      Height          =   195
      Left            =   7185
      TabIndex        =   0
      Top             =   615
      Width           =   525
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      Height          =   285
      Left            =   75
      TabIndex        =   44
      Top             =   45
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This tells you the result of the 5bit input of pins 10, 11, 12, 13, 15"
      Height          =   675
      Left            =   5130
      TabIndex        =   43
      Top             =   3690
      Width           =   1695
   End
   Begin VB.Label lblAlarm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alarm Clock"
      Height          =   285
      Left            =   7350
      TabIndex        =   42
      Top             =   45
      Width           =   1485
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ground"
      Height          =   210
      Left            =   2535
      TabIndex        =   41
      Top             =   1275
      Width           =   585
   End
   Begin VB.Line Line3 
      X1              =   690
      X2              =   5535
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line2 
      X1              =   5520
      X2              =   5520
      Y1              =   975
      Y2              =   1275
   End
   Begin VB.Line Line1 
      X1              =   690
      X2              =   690
      Y1              =   960
      Y2              =   1290
   End
   Begin VB.Label lblresis 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5115
      TabIndex        =   40
      Top             =   3450
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   270
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   8400
   End
   Begin VB.Image Image1 
      Height          =   6960
      Left            =   1020
      Picture         =   "Main.frx":241C2
      Top             =   3375
      Width           =   6615
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name - Parallel Port Controller
'Author - Josh Gaby
'Date - 28/10/2004

Option Explicit

Public Function Data_Port()

'8bit binary

'Declare variables
Dim D As Integer
Dim A As String

'Get the binary from check box's D08, D07, D06, D05, D04, D03, D02 and D01
'                           eg:
'                                Off, Off, Off,  On,  On, Off, Off and  On
'                                                =
'                                  0,   0,   0,   1,   1,   0,   0,      1
A = D08.Value & D07.Value & D06.Value & D05.Value & D04.Value & D03.Value & D02.Value & D01.Value

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

'Display the integer on the form
lblDi.Caption = A

'Display the 8bit binary on the form
lblDb.Caption = D

'Send the integer to the parallel port using inpout.dll (The "&H378" tells it that it is a data port and the integer tell it wich pins to turn on)
Out Val("&H378"), Val(D)
End Function

Public Function Status_Port()

'5bit binary

'Declare variables
Dim E As Integer
Dim b As String

'Get the binary from check box's S05, S04, S03, S02 and S01
'                           eg:
'                                Off,  On, Off, Off and  On
'                                                =
'                                  1,   1,   0,   0,      1
b = S05.Value & S04.Value & S03.Value & S02.Value & S01.Value

'Convert binary into integer
'                           eg:
'                                  0  |  1  |  0  |  0  |  1
'                                 016 | 008 | 004 | 002 | 001
'                                       008       +       001  = 9
If S05.Value = 1 Then E = E + 16
If S04.Value = 1 Then E = E + 8
If S03.Value = 1 Then E = E + 4
If S02.Value = 1 Then E = E + 2
If S01.Value = 1 Then E = E + 1

'Display the integer on the form
lblSi.Caption = b

'Display the 5bit binary on the form
lblSb.Caption = E

'Send the integer to the parallel port using inpout.dll (The "&H379" tells it that it is a status port and the integer tell it wich pins to turn on)
Out Val("&H379"), Val(E)
End Function
Public Function Control_Port()

'4bit binary

'Declare variables
Dim F As Integer
Dim C As String

'Get the binary from check box's C04, C03, C02 and C01
'                           eg:
'                                On,   On,  On and  On
'                                                =
'                                 1,   1,   1,      1
C = C04.Value & C03.Value & C02.Value & C01.Value

'Convert binary into integer
'                           eg:
'                                   1  |  1  |  1  |  1
'                                  008 | 004 | 002 | 001
'                                  008 + 004 + 002 + 001  = 15
'By default control ports have a value of 1 so to make the check box's behave the same way as
'the others I had to make them add when the check box's turn Off instead of On
If C04.Value = 0 Then F = F + 8
If C03.Value = 0 Then F = F + 4
If C02.Value = 0 Then F = F + 2
If C01.Value = 0 Then F = F + 1

'Display the integer on the form
'I had to minus F from 15 so that it would show correctly on the form
lblCi.Caption = 15 - F
'Display the 4bit binary on the form
lblCb.Caption = C

'Send the integer to the parallel port using inpout.dll (The "&H37A" tells it that it is a control port and the integer tell it wich pins to turn on)
Out Val("&H37A"), Val(F)
End Function

Private Sub C01_Click()
'Call public Function Control_Port
Call Control_Port
End Sub

Private Sub C02_Click()
'Call public Function Control_Port
Call Control_Port
End Sub

Private Sub C03_Click()
'Call public Function Control_Port
Call Control_Port
End Sub

Private Sub C04_Click()
'Call public Function Control_Port
Call Control_Port
End Sub

Private Sub D01_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D02_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D03_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D04_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D05_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D06_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D07_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub D08_Click()
'Call public Function Data_Port
Call Data_Port
End Sub

Private Sub Form_Load()
'Send all pins the Off signal (0)
Call Data_Port
Call Status_Port
Call Control_Port
End Sub

Private Sub lblAbout_Click()
'Show form About
About.Show
End Sub

Private Sub lblAlarm_Click()
'Show the alarm form
Alarm.Show
End Sub

Private Sub S01_Click()
'Call public Function Satus_Port
Call Status_Port
End Sub

Private Sub S02_Click()
'Call public Function Satus_Port
Call Status_Port
End Sub

Private Sub S03_Click()
'Call public Function Satus_Port
Call Status_Port
End Sub

Private Sub S04_Click()
'Call public Function Satus_Port
Call Status_Port
End Sub

Private Sub S05_Click()
'Call public Function Satus_Port
Call Status_Port
End Sub

Private Sub tmrResis_Timer()
lblresis.Caption = Str(Inp(Val("&H379")))
End Sub
