VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form sema4SetFrame 
   BorderStyle     =   0  'None
   Caption         =   "Servoset"
   ClientHeight    =   4575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton setButton 
      Caption         =   "Set"
      Height          =   375
      Left            =   2880
      TabIndex        =   61
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton restoreButton 
      Caption         =   "Restore"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame servoSettingOptionGroup 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   13
         Left            =   1200
         TabIndex        =   59
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   35
         Left            =   1200
         TabIndex        =   57
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   37
         Left            =   1200
         TabIndex        =   56
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   12
         Left            =   1200
         TabIndex        =   55
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   36
         Left            =   1200
         TabIndex        =   54
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   38
         Left            =   1200
         TabIndex        =   53
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   52
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   39
         Left            =   1200
         TabIndex        =   51
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   15
         Left            =   1200
         TabIndex        =   50
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   47
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   28
         Left            =   840
         TabIndex        =   46
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   29
         Left            =   840
         TabIndex        =   45
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   31
         Left            =   840
         TabIndex        =   44
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   43
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   30
         Left            =   840
         TabIndex        =   42
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   32
         Left            =   840
         TabIndex        =   41
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   40
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   39
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   38
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   37
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   36
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   23
         Left            =   480
         TabIndex        =   35
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   34
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   32
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   26
         Left            =   480
         TabIndex        =   31
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   27
         Left            =   480
         TabIndex        =   29
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   28
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   255
      End
   End
   Begin VB.HScrollBar valueScroller 
      Height          =   255
      LargeChange     =   16
      Left            =   1200
      Max             =   255
      TabIndex        =   2
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton saveButton 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton runButton 
      Caption         =   "Run"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin MSCommLib.MSComm ComPort 
      Left            =   3000
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   128
      OutBufferSize   =   128
   End
   Begin VB.Label valueText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   62
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label valueLabel 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   120
      TabIndex        =   60
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label servo4Label 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   195
      Left            =   2280
      TabIndex        =   49
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo1Label 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   195
      Left            =   1200
      TabIndex        =   48
      Top             =   120
      Width           =   285
   End
   Begin VB.Label onBounce3Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 3"
      Height          =   195
      Left            =   90
      TabIndex        =   27
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label onBounce2Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 2"
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label onBounce1Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 1"
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label offBounce3Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 3"
      Height          =   195
      Left            =   90
      TabIndex        =   24
      Top             =   480
      Width           =   945
   End
   Begin VB.Label offBounce2Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 2"
      Height          =   195
      Left            =   90
      TabIndex        =   23
      Top             =   840
      Width           =   945
   End
   Begin VB.Label offBounce1Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 1"
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label onSpeedLabel 
      AutoSize        =   -1  'True
      Caption         =   "OnSpeed"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label offSpeedLabel 
      AutoSize        =   -1  'True
      Caption         =   "Off Speed"
      Height          =   195
      Left            =   315
      TabIndex        =   20
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label onPositionLabel 
      AutoSize        =   -1  'True
      Caption         =   "On Position"
      Height          =   195
      Left            =   225
      TabIndex        =   19
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label offPositionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Off Position"
      Height          =   195
      Left            =   225
      TabIndex        =   18
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label servo2Label 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo3Label 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servosLabel 
      AutoSize        =   -1  'True
      Caption         =   "Servo:"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
   Begin VB.Menu exitMenu 
      Caption         =   "Exit"
   End
   Begin VB.Menu comPortMenu 
      Caption         =   "Com"
   End
End
Attribute VB_Name = "sema4SetFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum OperatingMode
    running
    setting
End Enum

Dim settingValue(0 To 39) As Integer
Dim settingIndex As Integer
Dim currentMode As OperatingMode
Dim dummy As Integer

Private Sub openComPort()

If True = ComPort.PortOpen Then
    ComPort.PortOpen = False
End If

ComPort.CommPort = ComPortNumber
ComPort.PortOpen = True

End Sub

Private Sub setRunningMode()

currentMode = running
runButton.Enabled = False
setButton.Enabled = True
saveButton.Enabled = False
restoreButton.Enabled = False
servoSettingOptionGroup.Enabled = False
valueScroller.Enabled = False

End Sub

Private Sub setSettingMode()

currentMode = setting
runButton.Enabled = True
setButton.Enabled = False
saveButton.Enabled = True
restoreButton.Enabled = True
servoSettingOptionGroup.Enabled = True
valueScroller.Enabled = True

End Sub

Private Sub sendCommand(commandString As String)

On Error GoTo comPortFailure

Dim n As Integer

For n = 1 To 20
    dummy = DoEvents
    ComPort.Output = Chr(0) + commandString + "000"
Next

Exit Sub

comPortFailure:
dummy = MsgBox("COM Port failed", vbOKOnly, "Error")
comPortMenu_Click

End Sub

Private Sub Form_Load()

For settingIndex = 39 To 0 Step -1
    settingValue(settingIndex) = 127
Next

openComPort
settingIndex = 0
servoSettingOption(settingIndex).Value = True
setRunningMode

End Sub

Private Sub comPortMenu_Click()

selectComPort
openComPort

End Sub

Private Sub exitMenu_Click()

End

End Sub

Private Sub runButton_Click()

setRunningMode

End Sub

Private Sub setButton_Click()

On Error GoTo comPortFailure

setSettingMode

While (setting = currentMode)
    dummy = DoEvents
    ComPort.Output = Chr(0) + Chr(65 + settingIndex) + Format(settingValue(settingIndex), "000")
Wend

Exit Sub

comPortFailure:
dummy = MsgBox("COM Port failed", vbOKOnly, "Error")
comPortMenu_Click

End Sub

Private Sub saveButton_Click()

setRunningMode

sendCommand ("@")

End Sub

Private Sub restoreButton_Click()

setRunningMode

sendCommand ("#")

End Sub

Private Sub servoSettingOption_Click(Index As Integer)

settingIndex = Index
valueScroller.Value = settingValue(settingIndex)

End Sub

Private Sub valueScroller_Change()

settingValue(settingIndex) = valueScroller.Value
valueText.Caption = settingValue(settingIndex)

End Sub

