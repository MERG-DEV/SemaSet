VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sema4SetForm 
   BorderStyle     =   0  'None
   Caption         =   "Servo4Sem4Set"
   ClientHeight    =   4935
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4245
   Icon            =   "sema4set.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox valuetext 
      Height          =   285
      Left            =   720
      MaxLength       =   3
      TabIndex        =   68
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton centerButton 
      Caption         =   "Centre"
      Height          =   375
      Left            =   3000
      TabIndex        =   67
      ToolTipText     =   "Set value to midpoint"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton setallButton 
      Caption         =   "Set &All"
      Height          =   375
      Left            =   3000
      TabIndex        =   64
      ToolTipText     =   "Download all setting values to module"
      Top             =   1320
      Width           =   975
   End
   Begin MSComDlg.CommonDialog settingsFileDialog 
      Left            =   3720
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   38
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   36
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   27
         Left            =   480
         TabIndex        =   35
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   34
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   26
         Left            =   480
         TabIndex        =   33
         Top             =   2880
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
         Index           =   4
         Left            =   480
         TabIndex        =   31
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   30
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   23
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   27
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   26
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   25
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   24
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   32
         Left            =   840
         TabIndex        =   23
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   30
         Left            =   840
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   31
         Left            =   840
         TabIndex        =   20
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   29
         Left            =   840
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   28
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   17
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   15
         Left            =   1200
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   39
         Left            =   1200
         TabIndex        =   15
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   14
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   38
         Left            =   1200
         TabIndex        =   13
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   36
         Left            =   1200
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   12
         Left            =   1200
         TabIndex        =   11
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   37
         Left            =   1200
         TabIndex        =   10
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   35
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   13
         Left            =   1200
         TabIndex        =   7
         Top             =   2160
         Width           =   255
      End
   End
   Begin VB.CommandButton setButton 
      Caption         =   "&Set"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Change setting value interactively"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton resetButton 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Module resets settings from EEPROM"
      Top             =   3240
      Width           =   975
   End
   Begin VB.HScrollBar valueScroller 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   2
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton storeButton 
      Caption         =   "S&tore"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Module stores settings to EEPROM"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton runButton 
      Caption         =   "&Run"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Allow module to run normally"
      Top             =   1800
      Width           =   975
   End
   Begin MSCommLib.MSComm comPort 
      Left            =   3120
      Top             =   -480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   128
      OutBufferSize   =   128
   End
   Begin VB.Frame xtndTravelSelectionGroup 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1080
      TabIndex        =   70
      Top             =   4080
      Width           =   1455
      Begin VB.CheckBox xtndTravelSelection 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   74
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox xtndTravelSelection 
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   73
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox xtndTravelSelection 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   72
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox xtndTravelSelection 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   71
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label connectionText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   66
      Top             =   960
      Width           =   615
   End
   Begin VB.Label connectionLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Connection:"
      Height          =   195
      Left            =   3120
      TabIndex        =   65
      Top             =   720
      Width           =   855
   End
   Begin VB.Label compatabilityLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Compatability:"
      Height          =   195
      Left            =   3000
      TabIndex        =   63
      Top             =   120
      Width           =   975
   End
   Begin VB.Label compatabilityText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Sema4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   62
      Top             =   360
      Width           =   975
   End
   Begin VB.Label servo4Label 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   195
      Left            =   2310
      TabIndex        =   61
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo1Label 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   195
      Left            =   1230
      TabIndex        =   60
      Top             =   120
      Width           =   285
   End
   Begin VB.Label onBounce3Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 3"
      Height          =   195
      Left            =   120
      TabIndex        =   59
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label onBounce2Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 2"
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label onBounce1Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 1"
      Height          =   195
      Left            =   120
      TabIndex        =   57
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label offBounce3Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 3"
      Height          =   195
      Left            =   120
      TabIndex        =   56
      Top             =   480
      Width           =   945
   End
   Begin VB.Label offBounce2Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 2"
      Height          =   195
      Left            =   120
      TabIndex        =   55
      Top             =   840
      Width           =   945
   End
   Begin VB.Label offBounce1Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 1"
      Height          =   195
      Left            =   120
      TabIndex        =   54
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label onSpeedLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "On Speed"
      Height          =   195
      Left            =   360
      TabIndex        =   53
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label offSpeedLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Off Speed"
      Height          =   195
      Left            =   345
      TabIndex        =   52
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label onPositionLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "On Position"
      Height          =   195
      Left            =   255
      TabIndex        =   51
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label offPositionLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Off Position"
      Height          =   195
      Left            =   255
      TabIndex        =   50
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label servo2Label 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   195
      Left            =   1590
      TabIndex        =   49
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo3Label 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   195
      Left            =   1950
      TabIndex        =   48
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servosLabel 
      AutoSize        =   -1  'True
      Caption         =   "Servo:"
      Height          =   195
      Left            =   600
      TabIndex        =   47
      Top             =   120
      Width           =   465
   End
   Begin VB.Label xtndTravelLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Long Travel"
      Height          =   195
      Left            =   120
      TabIndex        =   69
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label valueLabel 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   450
   End
   Begin VB.Menu fileMenu 
      Caption         =   "&File"
      Begin VB.Menu fileNewMenuItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu fileOpenMenuItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu fileSaveMenuItem 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu fileSaveAsMenuItem 
         Caption         =   "Save &As"
      End
      Begin VB.Menu fileExitMenuItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu optionsMenu 
      Caption         =   "&Options"
      Begin VB.Menu optSerPortMenuItem 
         Caption         =   "&Serial Port"
      End
      Begin VB.Menu optCompatSubMenu 
         Caption         =   "&Compatability"
         Begin VB.Menu optCompatServo4MenuItem 
            Caption         =   "Se&rvo4"
         End
         Begin VB.Menu optCompatSema4MenuItem 
            Caption         =   "Sema&4"
         End
         Begin VB.Menu optCompatSema4bMenuItem 
            Caption         =   "Sema4&b"
         End
         Begin VB.Menu optCompatSema4cMenuItem 
            Caption         =   "Sema4&c"
         End
      End
   End
   Begin VB.Menu helpMenu 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu helpAboutMenuItem 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "sema4SetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum OperatingMode
    RUNNING
    SETTING
    OFFLINE
End Enum

' Key code to indicate completion of direct value input to valueText TextBox
Const RTN_KEYCODE As Integer = 13

' Settings file format version
Const SETTINGS_FILE_FORMAT_VERSION As Integer = 0

' Maximum speed for Servo4
Const SERVO4_MAX_SPEED As Integer = 7

' Number of setting values for Servo (on and off, speed and position)
Const SERVO_SETTINGS As Integer = 4

' Number of setting values for Servo4
Const SERVO4_SETTINGS As Integer = 4 * SERVO_SETTINGS

' Number of extra setting values for Sema (3 off and 3 on bounces)
Const SEMA_SETTINGS  As Integer = 6

' Number of setting values for Sema4
Const SEMA4_SETTINGS  As Integer = 4 * (SEMA_SETTINGS + SERVO_SETTINGS)

' Number of times to send a non streaming command or setting string
Const SEND_ITTERATIONS As Integer = 15

' Default value to assign to new setting value
Const DEFAULT_SETTING As Integer = 127

' Transmitted command characters
Const SYNCH_BYTE    As Integer = 0  ' ASCII null
Const SETTING_BASE  As Integer = 65 ' ASCII A
Const TRVL_SETTING  As Integer = 56
Const STORE_COMMAND As String = "@"
Const RESET_COMMAND As String = "#"
Const RUN_COMMAND   As String = "$"

' Compatability names
Const servo4CompatabilityText As String = "Servo4"
Const sema4CompatabilityText  As String = "Sema4"
Const sema4bCompatabilityText As String = "Sema4b"
Const sema4cCompatabilityText As String = "Sema4c"

' Arrays for setting values and commands, layout:
'  Servo 1
'   Off Position, On Position, Off Speed, On Speed
'  Servo 2
'   Off Position, On Position, Off Speed, On Speed
'  Servo 3
'   Off Position, On Position, Off Speed, On Speed
'  Servo 4
'   Off Position, On Position, Off Speed, On Speed
'  Servo 1
'   Off Bounce 1, Off Bounce 2, Off Bounce 3,
'   On Bounce 1,  On Bounce 2,  On Bounce 3,
'  Servo 2
'   Off Bounce 1, Off Bounce 2, Off Bounce 3,
'   On Bounce 1,  On Bounce 2,  On Bounce 3,
'  Servo 3
'   Off Bounce 1, Off Bounce 2, Off Bounce 3,
'   On Bounce 1,  On Bounce 2,  On Bounce 3,
'  Servo 4
'   Off Bounce 1, Off Bounce 2, Off Bounce 3,
'   On Bounce 1,  On Bounce 2,  On Bounce 3
'  Servo extended travel selections
Dim settingValue(0 To (SEMA4_SETTINGS - 1))   As Integer
Dim settingCommand(0 To (SEMA4_SETTINGS - 1)) As Integer

' Define array index constants for certain values
Const OFF_PSTN_NDX_1 As Integer = 0
Const ON_PSTN_NDX_1  As Integer = 1
Const OFF_SPD_NDX_1  As Integer = 2
Const ON_SPD_NDX_1   As Integer = 3
Const OFF_PSTN_NDX_2 As Integer = 4
Const ON_PSTN_NDX_2  As Integer = 5
Const OFF_SPD_NDX_2  As Integer = 6
Const ON_SPD_NDX_2   As Integer = 7
Const OFF_PSTN_NDX_3 As Integer = 8
Const ON_PSTN_NDX_3  As Integer = 9
Const OFF_SPD_NDX_3  As Integer = 10
Const ON_SPD_NDX_3   As Integer = 11
Const OFF_PSTN_NDX_4 As Integer = 12
Const ON_PSTN_NDX_4  As Integer = 13
Const OFF_SPD_NDX_4  As Integer = 14
Const ON_SPD_NDX_4   As Integer = 15

Dim settingIndex    As Integer
Dim settingsChanged As Boolean

Dim settingsFilename As String

Dim currentMode As OperatingMode

Private Sub openComPort(newComPortNumber As Integer)

On Error GoTo commerror

' If COM port currently open close it
If True = comPort.PortOpen Then
    comPort.PortOpen = False
End If

' Set new COM port number and open COM port
comPort.CommPort = newComPortNumber

comPort.PortOpen = True

Exit Sub

commerror:

MsgBox "Unable to open the selected COM Port. Please choose another port.", _
       vbExclamation, _
       "COM Port Error"

End

End Sub

Public Sub selectComPort(oldComPortNumber As Integer)

Dim newComPortName As String
Dim newComPortNumber As Integer

' Prompt user to select COM port for connection
newComPortName = InputBox("Select COM Port", _
                          "COM Port number", _
                          oldComPortNumber)

' Convert entered COM Port string to an integer value
newComPortNumber = CInt(Val(newComPortName))

' Ensure COM port number is greater than 0
If (1 > newComPortNumber) Then
    If True = comPort.PortOpen Then
        comPort.PortOpen = False
    End If
    connectionText.Caption = "Offline"
    setOffline
Else
    connectionText.Caption = "COM" + newComPortName
    openComPort (newComPortNumber)

    setRunningMode
End If

End Sub

Private Sub changeComPort()

selectComPort comPort.CommPort

End Sub

Private Sub comPortFailed()

MsgBox "Error accessing COM port, " + Error, vbOKOnly, "COM Port Error"

changeComPort

End Sub

Private Sub limitSettingValue(testValue As Integer)

If testValue > valueScroller.Max Then
    testValue = valueScroller.Max
End If
If testValue < valueScroller.Min Then
    testValue = valueScroller.Min
End If

End Sub

Private Sub changeSettingValue(newValue As Integer)

limitSettingValue newValue

If newValue <> settingValue(settingIndex) Then
    settingValue(settingIndex) = newValue
    valueScroller.Value = newValue
    valuetext.Text = newValue
    settingsChanged = True
End If

End Sub

Private Sub selectSetting(newSettingIndex As Integer)

settingIndex = newSettingIndex

valueScroller.Max = 255
valueScroller.LargeChange = 18

Select Case settingIndex

Case ON_SPD_NDX_1, OFF_SPD_NDX_1, ON_SPD_NDX_2, OFF_SPD_NDX_2, _
     ON_SPD_NDX_3, OFF_SPD_NDX_3, ON_SPD_NDX_4, OFF_SPD_NDX_4
    
    ' Selected setting is a speed
    If compatabilityText.Caption = servo4CompatabilityText Then
        ' Compatability mode is Servo4, limit maximum speed
        valueScroller.Max = SERVO4_MAX_SPEED
        valueScroller.LargeChange = 1
    End If

End Select
    
limitSettingValue settingValue(settingIndex)

valueScroller.Value = settingValue(settingIndex)
valuetext.Text = settingValue(settingIndex)

servoSettingOption(settingIndex).Value = True

End Sub

Private Sub disableExtendedTravelSelections()

xtndTravelSelection(0).Enabled = False
xtndTravelSelection(1).Enabled = False
xtndTravelSelection(2).Enabled = False
xtndTravelSelection(3).Enabled = False
xtndTravelSelectionGroup.Enabled = False

End Sub

Private Sub enableExtendedTravelSelections()

xtndTravelSelectionGroup.Enabled = True
xtndTravelSelection(0).Enabled = True
xtndTravelSelection(1).Enabled = True
xtndTravelSelection(2).Enabled = True
xtndTravelSelection(3).Enabled = True

End Sub

Private Sub setExtendedTravelSelections(newOptions As Integer)

If (0 <> (&H1 And newOptions)) Then
    xtndTravelSelection(0).Value = vbChecked
Else
    xtndTravelSelection(0).Value = vbUnchecked
End If

If (0 <> (&H2 And newOptions)) Then
    xtndTravelSelection(1).Value = vbChecked
Else
    xtndTravelSelection(1).Value = vbUnchecked
End If

If (0 <> (&H4 And newOptions)) Then
    xtndTravelSelection(2).Value = vbChecked
Else
    xtndTravelSelection(2).Value = vbUnchecked
End If

If (0 <> (&H8 And newOptions)) Then
    xtndTravelSelection(3).Value = vbChecked
Else
    xtndTravelSelection(3).Value = vbUnchecked
End If

End Sub

Private Function getExtendedTravelSelections() As Integer

getExtendedTravelSelections = 0

If (vbChecked = xtndTravelSelection(0).Value) Then
    getExtendedTravelSelections = (&H1 Or getExtendedTravelSelections)
End If

If (vbChecked = xtndTravelSelection(1).Value) Then
    getExtendedTravelSelections = (&H2 Or getExtendedTravelSelections)
End If

If (vbChecked = xtndTravelSelection(2).Value) Then
    getExtendedTravelSelections = (&H4 Or getExtendedTravelSelections)
End If

If (vbChecked = xtndTravelSelection(3).Value) Then
    getExtendedTravelSelections = (&H8 Or getExtendedTravelSelections)
End If

End Function

Private Sub disableAllCompatabilitySelections()

    optCompatServo4MenuItem.Enabled = False
    optCompatSema4MenuItem.Enabled = False
    optCompatSema4bMenuItem.Enabled = False
    optCompatSema4cMenuItem.Enabled = False

End Sub

Private Sub enableAllCompatabilitySelections()

    optCompatServo4MenuItem.Enabled = True
    optCompatSema4MenuItem.Enabled = True
    optCompatSema4bMenuItem.Enabled = True
    optCompatSema4cMenuItem.Enabled = True

End Sub

Private Sub disableServo4CompatabilitySelection()

    ' Allow change of compatability selection, excluding Servo4
    enableAllCompatabilitySelections
    optCompatServo4MenuItem.Enabled = False

End Sub

Private Sub disableSema4cCompatabilitySelection()

    ' Allow change of compatability selection, excluding Sema4b
    enableAllCompatabilitySelections
    optCompatSema4cMenuItem.Enabled = False

End Sub

Private Sub disableSema4bCompatabilitySelection()

    ' Allow change of compatability selection, excluding Sema4b
    enableAllCompatabilitySelections
    optCompatSema4bMenuItem.Enabled = False

End Sub

Private Sub disableSema4CompatabilitySelection()

    ' Allow change of compatability selection, excluding Sema4
    enableAllCompatabilitySelections
    optCompatSema4MenuItem.Enabled = False

End Sub

Private Sub setOffline()

currentMode = OFFLINE

' Allow change of compatability selection, excluding that currently selected
Select Case compatabilityText.Caption
    Case servo4CompatabilityText
        disableServo4CompatabilitySelection
    Case sema4bCompatabilityText
        disableSema4bCompatabilitySelection
    Case sema4cCompatabilityText
        disableSema4cCompatabilitySelection
        enableExtendedTravelSelections
    Case Else
        disableSema4CompatabilitySelection
End Select

setallButton.Enabled = False
runButton.Enabled = False
setButton.Enabled = False
storeButton.Enabled = False
resetButton.Enabled = False
centerButton.Enabled = True
valueScroller.Enabled = True
valuetext.Enabled = True

End Sub

Private Sub setRunningMode()

currentMode = RUNNING

If comPort.PortOpen Then
    ' Enable appropriate Compatability selection
    Select Case compatabilityText.Caption
        Case servo4CompatabilityText
            disableServo4CompatabilitySelection
        Case sema4bCompatabilityText
            disableSema4bCompatabilitySelection
            sendCommand (RUN_COMMAND) ' Ensure module is not in Set mode
        Case sema4cCompatabilityText
            disableSema4cCompatabilitySelection
            sendCommand (RUN_COMMAND) ' Ensure module is not in Set mode
        Case Else
            disableSema4CompatabilitySelection
            sendCommand (RUN_COMMAND) ' Ensure module is not in Set mode
    End Select

    ' Disable selection of Run mode, enable selection of Set mode
    runButton.Enabled = False
    setButton.Enabled = True

    ' Enable sending of commands to Send, Store, and Reset settings
    setallButton.Enabled = True
    storeButton.Enabled = True
    resetButton.Enabled = True

    ' Disable changine of setting value
    centerButton.Enabled = False
    valueScroller.Enabled = False
    valuetext.Enabled = False

    ' Disable changine of extended servo travel selections
    disableExtendedTravelSelections

Else
    ' COM port not available so act just as an offline settings editor
    setOffline
End If

End Sub

Private Sub setSettingMode()

currentMode = SETTING

disableAllCompatabilitySelections

' Enable selection of Run mode, disable selection of Set mode
runButton.Enabled = True
setButton.Enabled = False

' Disable sending of commands to Send, Store, and Reset settings
setallButton.Enabled = False
storeButton.Enabled = False
resetButton.Enabled = False

' Enable changine of setting value
centerButton.Enabled = True
valueScroller.Enabled = True
valuetext.Enabled = True

' Enable changine of extended servo travel selections
enableExtendedTravelSelections

End Sub

Private Sub streamCurrentSetting()
' Continuosly send the currently selected setting value so the module tracks
' changes interactively

If comPort.PortOpen Then
    On Error GoTo comPortFailure

    While (SETTING = currentMode)
        ' Perform event dispatch to keep GUI alive,
        ' allows currentMode to be changed
        DoEvents

        ' Send setting command and value for currently selected setting
        comPort.Output = Chr(SYNCH_BYTE) _
                         + Chr(SETTING_BASE + settingCommand(settingIndex)) _
                         + Format(settingValue(settingIndex), "000")
    Wend
End If

Exit Sub

comPortFailure:
    comPortFailed

End Sub

Private Sub sendCommand(commandCharacter As String, _
                        Optional commandValue As Integer = 0, _
                        Optional sendItterations As Integer = SEND_ITTERATIONS)
' Send the given command, and optionally a value for the command, repeatedly
' a set number of times to allow for garbled reception as link has no handshake

Dim n As Integer

If comPort.PortOpen Then
    On Error GoTo comPortFailure

    For n = 1 To sendItterations
        ' Perform event dispatch to keep GUI alive
        DoEvents

        ' Send command and value
        comPort.Output = Chr(SYNCH_BYTE) _
                         + commandCharacter _
                         + Format(commandValue, "000")
    Next
End If

Exit Sub

comPortFailure:
    comPortFailed

End Sub

Private Sub sendSetting(settingCommand As Integer, _
                        Optional commandValue As Integer = 0, _
                        Optional sendItterations As Integer = SEND_ITTERATIONS)
' Send the command for a setting, and optionally a value for the command, repeatedly
' a set number of times to allow for garbled reception as link has no handshake

If (SETTING = currentMode) Then
    sendCommand Chr(SETTING_BASE + settingCommand), commandValue, sendItterations
End If
                        
End Sub

Private Sub sendCurrentSettings()
' Download all the current settings

sema4SetForm.MousePointer = vbHourglass

Dim sendIndex As Integer

' Walk the array of setting values
For sendIndex = LBound(settingValue) To UBound(settingValue)
    ' Test if option button for setting is enabled
    If servoSettingOption(sendIndex).Enabled Then
        ' Send setting command and value
        sendCommand Chr(SETTING_BASE + settingCommand(sendIndex)), _
                    settingValue(sendIndex)
    End If
Next

If compatabilityText.Caption = sema4cCompatabilityText Then
    sendCommand Chr(SETTING_BASE + TRVL_SETTING), getExtendedTravelSelections
End If

If compatabilityText.Caption <> servo4CompatabilityText Then
    sendCommand (RUN_COMMAND) ' Ensure module leaves Set mode after download
End If

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub checkIfSaveNeeded(Optional beforeAction As String = "overwriting")
' Check if any settings have been changed and if so offer a chance to save
' these before proceeding

If settingsChanged Then
    If vbYes = MsgBox("Settings have changed, save before " + _
                          beforeAction + "?", _
                      vbYesNo) Then
        saveSettings
    End If
End If

End Sub

Private Sub newSettings()
' After checking if current settings need saving change all settings to
' default values

setRunningMode

checkIfSaveNeeded

' Walk the array of setting values restoring all to default value
For settingIndex = LBound(settingValue) To UBound(settingValue)
    settingValue(settingIndex) = DEFAULT_SETTING
    selectSetting settingIndex
Next

setExtendedTravelSelections 0

' Select first setting option control and display corresponding value
selectSetting 0

settingsChanged = False

End Sub

Private Sub loadSettings()
' After checking if current settings need saving load all settings from file

setRunningMode

checkIfSaveNeeded

On Error GoTo errorCancelLoad

' Get name of file to read settings from
settingsFileDialog.ShowOpen
settingsFilename = settingsFileDialog.FileName

If "" = settingsFilename Then
    GoTo errorCancelLoad
End If

sema4SetForm.MousePointer = vbHourglass

' Open the settings file
Open settingsFilename For Input As #1

' Load compatability mode from file and set to same
Dim loadedCompatabilityText As String

Input #1, loadedCompatabilityText

If loadedCompatabilityText <> compatabilityText.Caption Then
    Select Case loadedCompatabilityText
        Case sema4cCompatabilityText
            setSema4cCompatabillity
        Case sema4bCompatabilityText
            setSema4bCompatabillity
        Case sema4CompatabilityText
            setSema4Compatabillity
        Case Else
            setServo4Compatabillity
    End Select
End If

' Check version of format for settings in file in order to support reading
' files written with previous versions of this program, do nothing at present
Dim loadedFileFormatVersion As Integer

Input #1, loadedFileFormatVersion

' Load the setting values from file
settingIndex = LBound(settingValue)

Do Until (EOF(1) Or (UBound(settingValue) < settingIndex))
    Input #1, settingValue(settingIndex)
    settingIndex = 1 + settingIndex
Loop

If (Not EOF(1)) Then
    Input #1, settingIndex
    setExtendedTravelSelections settingIndex
Else
    setExtendedTravelSelections 0
End If

' Close the settings file
Close #1

' Select first setting option control and display corresponding value
selectSetting 0

settingsChanged = False

errorCancelLoad:

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub saveSettings()
' Save the current setting values to file

On Error GoTo errorCancel

' If not already set then get name of file to write settings to
If "" = settingsFilename Then
    settingsFileDialog.ShowSave
    settingsFilename = settingsFileDialog.FileName
End If

If "" = settingsFilename Then
    MsgBox "Filename blank, settings not saved", vbOKOnly, "No filename"
End If

sema4SetForm.MousePointer = vbHourglass

' Open the settings file
Open settingsFilename For Output As #1

' Save current compatability mode and version of format for settings in file
Print #1, compatabilityText.Caption
Print #1, SETTINGS_FILE_FORMAT_VERSION

' Save the setting values to file
Dim outputIndex As Integer

' Walk the array of setting values writing value to file
For outputIndex = LBound(settingValue) To UBound(settingValue)
    Print #1, settingValue(outputIndex)
Next

Print #1, getExtendedTravelSelections

' Close the settings file
Close #1

settingsChanged = False

errorCancel:

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub initialiseServo4SettingCommands()

' Initialise command for each setting
For settingIndex = LBound(settingCommand) To UBound(settingCommand)
    settingCommand(settingIndex) = settingIndex
Next

End Sub

Private Sub initialiseSema4SettingCommands()

initialiseServo4SettingCommands

' Sema4 has alternative commands for equivalent Servo4 settings
settingCommand(OFF_PSTN_NDX_1) = 40
settingCommand(ON_PSTN_NDX_1) = 41
settingCommand(OFF_PSTN_NDX_2) = 42
settingCommand(ON_PSTN_NDX_2) = 43
settingCommand(OFF_PSTN_NDX_3) = 44
settingCommand(ON_PSTN_NDX_3) = 45
settingCommand(OFF_PSTN_NDX_4) = 46
settingCommand(ON_PSTN_NDX_4) = 47

End Sub

Private Sub initialiseSema4bSettingCommands()

initialiseSema4SettingCommands

' Sema4b has alternative commands for equivalent Sema4 settings
settingCommand(OFF_SPD_NDX_1) = 48
settingCommand(ON_SPD_NDX_1) = 49
settingCommand(OFF_SPD_NDX_2) = 50
settingCommand(ON_SPD_NDX_2) = 51
settingCommand(OFF_SPD_NDX_3) = 52
settingCommand(ON_SPD_NDX_3) = 53
settingCommand(OFF_SPD_NDX_4) = 54
settingCommand(ON_SPD_NDX_4) = 55

End Sub

Private Sub setServo4Compatabillity()

initialiseServo4SettingCommands

' Enable the option controls to select settings supported by Servo4
For settingIndex = 0 To (SERVO4_SETTINGS - 1)
    servoSettingOption(settingIndex).Visible = True
    servoSettingOption(settingIndex).Enabled = True
Next

' Disable the option controls to select settings not supported by Servo4
For settingIndex = SERVO4_SETTINGS To (SEMA4_SETTINGS - 1)
    servoSettingOption(settingIndex).Visible = False
    servoSettingOption(settingIndex).Enabled = False
Next
offBounce3Label.Enabled = False
offBounce2Label.Enabled = False
offBounce1Label.Enabled = False
onBounce1Label.Enabled = False
onBounce2Label.Enabled = False
onBounce3Label.Enabled = False

' Disable the selection controls to select extended servo travel not supported by Servo4
xtndTravelSelectionGroup.Visible = False
xtndTravelLabel.Enabled = False

' Select first setting option control and display corresponding value
selectSetting 0

' Update compatability mode display
compatabilityText.Caption = servo4CompatabilityText

disableServo4CompatabilitySelection

End Sub

Private Sub setSema4Compatabillity()

initialiseSema4SettingCommands

' Enable the option controls to select settings supported by Sema4
For settingIndex = 0 To (SEMA4_SETTINGS - 1)
    servoSettingOption(settingIndex).Visible = True
    servoSettingOption(settingIndex).Enabled = True
Next
offBounce3Label.Enabled = True
offBounce2Label.Enabled = True
offBounce1Label.Enabled = True
onBounce1Label.Enabled = True
onBounce2Label.Enabled = True
onBounce3Label.Enabled = True

' Disable the selection controls to select extended servo travel not supported by Servo4
xtndTravelSelectionGroup.Visible = False
xtndTravelSelectionGroup.Enabled = False
xtndTravelLabel.Enabled = False

' Select first setting option control and display corresponding value
selectSetting 0

' Update compatability mode display
compatabilityText.Caption = sema4CompatabilityText

disableSema4CompatabilitySelection

End Sub

Private Sub setSema4bCompatabillity()

' Sema4b is a derivative of Sema4
setSema4Compatabillity

' Difference is in commands used for certain values
initialiseSema4bSettingCommands

' Select first setting option control and display corresponding value
selectSetting 0

' Update compatability mode display
compatabilityText.Caption = sema4bCompatabilityText

disableSema4bCompatabilitySelection

End Sub

Private Sub setSema4cCompatabillity()

' Sema4c is a derivative of Sema4b
setSema4bCompatabillity

' Enable the selection controls to select extended servo travel supported by Servo4c
xtndTravelSelectionGroup.Visible = True
xtndTravelSelectionGroup.Enabled = True
xtndTravelLabel.Enabled = True

' Select first setting option control and display corresponding value
selectSetting 0

' Update compatability mode display
compatabilityText.Caption = sema4cCompatabilityText

disableSema4cCompatabilitySelection

End Sub

Private Sub convertSpeedToServo4()

settingValue(OFF_SPD_NDX_1) = settingValue(OFF_SPD_NDX_1) / 16
settingValue(ON_SPD_NDX_1) = settingValue(ON_SPD_NDX_1) / 16
settingValue(OFF_SPD_NDX_2) = settingValue(OFF_SPD_NDX_2) / 16
settingValue(ON_SPD_NDX_2) = settingValue(ON_SPD_NDX_2) / 16
settingValue(OFF_SPD_NDX_3) = settingValue(OFF_SPD_NDX_3) / 16
settingValue(ON_SPD_NDX_3) = settingValue(ON_SPD_NDX_3) / 16
settingValue(OFF_SPD_NDX_4) = settingValue(OFF_SPD_NDX_4) / 16
settingValue(ON_SPD_NDX_4) = settingValue(ON_SPD_NDX_4) / 16

End Sub

Private Sub convertSpeedFromServo4()

settingValue(OFF_SPD_NDX_1) = settingValue(OFF_SPD_NDX_1) * 16
settingValue(ON_SPD_NDX_1) = settingValue(ON_SPD_NDX_1) * 16
settingValue(OFF_SPD_NDX_2) = settingValue(OFF_SPD_NDX_2) * 16
settingValue(ON_SPD_NDX_2) = settingValue(ON_SPD_NDX_2) * 16
settingValue(OFF_SPD_NDX_3) = settingValue(OFF_SPD_NDX_3) * 16
settingValue(ON_SPD_NDX_3) = settingValue(ON_SPD_NDX_3) * 16
settingValue(OFF_SPD_NDX_4) = settingValue(OFF_SPD_NDX_4) * 16
settingValue(ON_SPD_NDX_4) = settingValue(ON_SPD_NDX_4) * 16

End Sub

Private Sub Form_Load()

' Initialisation when form is first loaded

' Set up File Dialog filter to match Sema4 settings files
settingsFileDialog.Filter = "Sema4Set Files (*.sm4)|*.sm4" _
                            + "|Text Files (*.txt)|*.txt" _
                            + "!All Files (*.*)|*.*"
settingsFileDialog.FilterIndex = 1

settingsFilename = ""

setServo4Compatabillity

newSettings

selectComPort 1

Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

checkIfSaveNeeded "exiting"

End

End Sub

Private Sub fileNewMenuItem_Click()

newSettings

End Sub

Private Sub fileOpenMenuItem_Click()

loadSettings

End Sub

Private Sub fileSaveMenuItem_Click()

saveSettings

End Sub

Private Sub fileSaveAsMenuItem_Click()

settingsFilename = ""
saveSettings

End Sub

Private Sub fileExitMenuItem_Click()

Unload Me

End Sub

Private Sub helpAboutMenuItem_Click()

sema4About.Show
  
End Sub

Private Sub optSerPortMenuItem_Click()

changeComPort

End Sub

Private Sub optCompatServo4MenuItem_Click()

convertSpeedToServo4
setServo4Compatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4MenuItem_Click()

If servo4CompatabilityText = compatabilityText.Caption Then
    convertSpeedFromServo4
End If
setSema4Compatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4bMenuItem_Click()

If servo4CompatabilityText = compatabilityText.Caption Then
    convertSpeedFromServo4
End If
setSema4bCompatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4cMenuItem_Click()

If servo4CompatabilityText = compatabilityText.Caption Then
    convertSpeedFromServo4
End If
setSema4cCompatabillity
settingsChanged = True

End Sub

Private Sub runButton_Click()

' Allow module to run normally
setRunningMode

End Sub

Private Sub setallButton_Click()

' Download all current setting values to module
sendCurrentSettings

End Sub

Private Sub setButton_Click()

' Change setting value interactively, module tracks current setting value
setSettingMode
streamCurrentSetting

End Sub

Private Sub storeButton_Click()

' Command module to store setting values into non-volatile memory
sema4SetForm.MousePointer = vbHourglass
sendCommand (STORE_COMMAND)
sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub resetButton_Click()

' Command module to reset setting values to defaults
sema4SetForm.MousePointer = vbHourglass
sendCommand (RESET_COMMAND)
sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub centerButton_Click()

If settingValue(settingIndex) <> _
       ((valueScroller.Max - valueScroller.Min) / 2) Then
    changeSettingValue ((valueScroller.Max - valueScroller.Min) / 2)
End If

End Sub

Private Sub servoSettingOption_Click(optionIndex As Integer)

If settingIndex <> optionIndex Then
    selectSetting optionIndex
End If

End Sub

Private Sub valueScroller_Change()

If settingValue(settingIndex) <> valueScroller.Value Then
    changeSettingValue valueScroller.Value
End If

End Sub

Private Sub valuetext_KeyUp(keyCode As Integer, shift As Integer)

If RTN_KEYCODE = keyCode And _
   (settingValue(settingIndex) <> CInt(Val(valuetext.Text))) Then
    changeSettingValue CInt(Val(valuetext.Text))
End If

End Sub

Private Sub valuetext_LostFocus()

If settingValue(settingIndex) <> CInt(Val(valuetext.Text)) Then
    changeSettingValue CInt(Val(valuetext.Text))
End If

End Sub

Private Sub xtndTravelSelection_Click(Index As Integer)

sendSetting TRVL_SETTING, getExtendedTravelSelections
settingsChanged = True

End Sub
