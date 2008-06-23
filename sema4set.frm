VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSema4Set 
   BorderStyle     =   0  'None
   Caption         =   "Servoset"
   ClientHeight    =   5070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1200
      TabIndex        =   29
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1200
      TabIndex        =   12
      Top             =   480
      Width           =   2895
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   14
         Left            =   1440
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   13
         Left            =   720
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   24
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSer 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   2
      Top             =   2760
      Width           =   4815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "Spd 1"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Spd 0"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "On"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Off"
      Height          =   255
      Left            =   1320
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Servo 4"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Servo 3"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Servo 2"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Servo 1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuCom 
      Caption         =   "Com"
   End
End
Attribute VB_Name = "frmSema4Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer
Dim but As Integer
Dim mode As Integer
Dim com As Integer

Private Sub cmdReset_Click()
Dim n As Integer
mode = 0
cmdRun.Caption = "Run"
For n = 1 To 20
out = Chr(0) + "#" + Format(num, "000")
MSComm1.Output = out
Next
End Sub

Private Sub cmdRun_Click()
Dim n As Integer
Dim out As String

Dim com_in As String
On Error GoTo comset
'com = 4

setport:
dummy = DoEvents
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
MSComm1.CommPort = ComPort
MSComm1.PortOpen = True




'GoTo setport:

mode = 1
cmdRun.Caption = "Setting"
again:
If mode = 0 Then
Exit Sub
End If

out = Chr(0) + Chr(65 + but) + Format(num, "000")

'MSComm1.PortOpen = True

MSComm1.Output = out
dummy = DoEvents
GoTo again

comset:

dummy = MsgBox("Wrong COM port", vbOKOnly, "Error")
Exit Sub
'com_in = InputBox("Set Com Port", "Com port is ", 4)
'If com_in = "" Then
'Exit Sub
'End If
'ComPort = Val(com_in)
'GoTo setport

End Sub

Private Sub cmdStop_Click()
Dim n As Integer
mode = 0
cmdRun.Caption = "Run"
For n = 1 To 20
out = Chr(0) + "@" + Format(num, "000")
MSComm1.Output = out
Next

End Sub

Private Sub Form_Load()
mode = 0
If MSComm1.PortOpen = False Then
Exit Sub
End If
End Sub

Private Sub HScroll1_Change()
num = HScroll1.Value
lblVal.Caption = num
End Sub





Private Sub mnuCom_Click()
ComPort = InputBox("Com port is  ", "Set Com Port", ComPort)
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
MSComm1.CommPort = ComPort
MSComm1.PortOpen = True

End Sub

Private Sub mnuExit_Click()
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
End
End Sub


Private Sub optSer_Click(Index As Integer)
Dim n As Integer
For n = 0 To 15
    If optSer(n) = True Then
    but = n
    Exit Sub
    End If
Next
End Sub
