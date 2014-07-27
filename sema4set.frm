VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sema4SetForm 
   BorderStyle     =   0  'None
   Caption         =   "Servo4Sem4Set"
   ClientHeight    =   5175
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4245
   Icon            =   "sema4set.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox valueText 
      Height          =   285
      Left            =   720
      MaxLength       =   3
      TabIndex        =   68
      Top             =   4800
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
      Caption         =   "S&et"
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
      Top             =   4800
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
      InBufferSize    =   5
      OutBufferSize   =   2256
   End
   Begin VB.Frame bounceSelectionGroup 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1080
      TabIndex        =   75
      Top             =   4080
      Width           =   1455
      Begin VB.CheckBox bounceSelection 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   76
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox bounceSelection 
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   77
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox bounceSelection 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   78
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox bounceSelection 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   79
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame xtndTravelSelectionGroup 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1080
      TabIndex        =   70
      Top             =   4440
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
      Caption         =   "Sema4d"
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
   Begin VB.Label bounceLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Bounce"
      Height          =   195
      Left            =   510
      TabIndex        =   80
      Top             =   4080
      Width           =   555
   End
   Begin VB.Label xtndTravelLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Long Travel"
      Height          =   195
      Left            =   210
      TabIndex        =   69
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label valueLabel 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   450
   End
   Begin VB.Menu fileMenu 
      Caption         =   "&Settings"
      Begin VB.Menu fileNewMenuItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu fileOpenMenuItem 
         Caption         =   "&Load"
         Shortcut        =   ^L
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
         Begin VB.Menu optCompatSema4dMenuItem 
            Caption         =   "Sema4&d"
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

' Key code to indicate completion of direct value input to valueText TextBox
Private Const RTN_KEYCODE As Integer = 13

' Compatability names
Private Const servo4Text As String = "Servo4"
Private Const sema4Text  As String = "Sema4"
Private Const sema4bText As String = "Sema4b"
Private Const sema4cText As String = "Sema4c"
Private Const sema4dText As String = "Sema4d"

Private Sub setExtendedTravelSelections(newOptions As Integer)

If (0 <> (SRV1_XTND_MASK And newOptions)) Then
    xtndTravelSelection(0).Value = vbChecked
Else
    xtndTravelSelection(0).Value = vbUnchecked
End If

If (0 <> (SRV2_XTND_MASK And newOptions)) Then
    xtndTravelSelection(1).Value = vbChecked
Else
    xtndTravelSelection(1).Value = vbUnchecked
End If

If (0 <> (SRV3_XTND_MASK And newOptions)) Then
    xtndTravelSelection(2).Value = vbChecked
Else
    xtndTravelSelection(2).Value = vbUnchecked
End If

If (0 <> (SRV4_XTND_MASK And newOptions)) Then
    xtndTravelSelection(3).Value = vbChecked
Else
    xtndTravelSelection(3).Value = vbUnchecked
End If

End Sub

Private Function getExtendedTravelSelections() As Integer

getExtendedTravelSelections = 0

If (vbChecked = xtndTravelSelection(0).Value) Then
    getExtendedTravelSelections = (SRV1_XTND_MASK Or getExtendedTravelSelections)
End If
    
If (vbChecked = xtndTravelSelection(1).Value) Then
    getExtendedTravelSelections = (SRV2_XTND_MASK Or getExtendedTravelSelections)
End If
    
If (vbChecked = xtndTravelSelection(2).Value) Then
    getExtendedTravelSelections = (SRV3_XTND_MASK Or getExtendedTravelSelections)
End If
    
If (vbChecked = xtndTravelSelection(3).Value) Then
    getExtendedTravelSelections = (SRV4_XTND_MASK Or getExtendedTravelSelections)
End If

End Function

Private Sub checkIfSaveNeeded(Optional beforeAction As String = "overwriting")
' Check if any settings have been changed and if so offer a chance to save
' these before proceeding

If (settingsChanged) Then
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

' Force into offline mode to prevent values being sent as they are loaded
setOffline

checkIfSaveNeeded

sema4SetForm.MousePointer = vbHourglass

' Walk the array of setting values restoring all to default value
For settingIndex = LBound(settingValue) To UBound(settingValue)
    settingValue(settingIndex) = DEFAULT_SETTING
    selectSetting settingIndex
Next

setExtendedTravelSelections 0

If (servo4Text = compatabilityText.Caption) Then
    convertSpeedToServo4
Else
    selectAllBounces
End If

' Select first setting option control and display corresponding value
selectSetting 0

settingsChanged = False

' Default into running mode
setRunningMode

' Clear display of any filename in the window title bar
sema4SetForm.Caption = ""

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub loadSettings()
' After checking if current settings need saving load all settings from file

' Force into offline mode to prevent values being sent as they are loaded
setOffline

checkIfSaveNeeded

On Error GoTo errorCancelLoad

' Get name of file to read settings from
settingsFileDialog.ShowOpen
settingsFilename = settingsFileDialog.FileName

If ("" = settingsFilename) Then
    GoTo errorCancelLoad
End If

sema4SetForm.MousePointer = vbHourglass

' Open the settings file
Open settingsFilename For Input As #1

' Load compatability mode from file and set to same
Dim loadedCompatabilityText As String

Input #1, loadedCompatabilityText

Select Case loadedCompatabilityText
    Case sema4dText
        setSema4dCompatabillity
    Case sema4cText
        setSema4cCompatabillity
    Case sema4bText
        setSema4bCompatabillity
    Case sema4Text
        setSema4Compatabillity
    Case Else
        setServo4Compatabillity
End Select

If (servo4Text <> compatabilityText.Caption) Then
    ' For any Sema4 initially select all bounces by default
    selectAllBounces
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

If (servo4Text <> compatabilityText.Caption) Then
    Dim selection As Integer
    For settingIndex = 0 To NUM_SERVOS - 1
        If (Not EOF(1)) Then
            Input #1, selection
            bounceSelection(settingIndex).Value = selection
        End If
    Next

    checkBounceSelections
Else
    limitServo4Speeds
End If

' Close the settings file
Close #1

' Select first setting option control and display corresponding value
selectSetting 0

settingsChanged = False

' Display the filename in the window title bar
sema4SetForm.Caption = _
    Right(settingsFilename, _
          (Len(settingsFilename) - InStrRev(settingsFilename, "\")))

errorCancelLoad:

' Default into running mode
setRunningMode

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub saveSettings()
' Save the current setting values to file

On Error GoTo errorCancel

' If not already set then get name of file to write settings to
If ("" = settingsFilename) Then
    settingsFileDialog.ShowSave
    settingsFilename = settingsFileDialog.FileName
End If

If ("" = settingsFilename) Then
    MsgBox "Filename blank, settings not saved", vbOKOnly, "No filename"
    GoTo errorCancel
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

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    Print #1, bounceSelection(servoIndex).Value
Next

' Close the settings file
Close #1

settingsChanged = False

' Display the filename in the window title bar
sema4SetForm.Caption = _
    Right(settingsFilename, _
          (Len(settingsFilename) - InStrRev(settingsFilename, "\")))

errorCancel:

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub setServo4Compatabillity()

If (sema4dText = compatabilityText.Caption) Then
    ' Sema4d speeds run 0 - fastest to 255 - slowest, previous versions were
    ' 0 - fastest then 1 - slowest to 255 - fastest
    flipSpeeds
End If

' Must have previously been in one of the Sema4 modes
convertSpeedToServo4

' Update compatability mode display
compatabilityText.Caption = servo4Text
disableCurrentCompatabilitySelection

initialiseServo4SettingCommands

deselectAllBounces
disableBounce
disableExtendedTravel

' Select first setting option control and display corresponding value
selectSetting 0

End Sub

Private Sub setSema4Compatabillity()

' Check the current compatability mode
Select Case compatabilityText.Caption
    Case sema4dText
        ' Sema4d speeds run 0 - fastest to 255 - slowest, previous versions
        ' were 0 - fastest then 1 - slowest to 255 - fastest
        flipSpeeds
    Case servo4Text
        convertSpeedFromServo4
End Select

' Update compatability mode display
compatabilityText.Caption = sema4Text
disableCurrentCompatabilitySelection

initialiseSema4SettingCommands

enableBounce

' Commands are initialised for Sema4 with bounce so use the appropriate Servo4
' command for any deslected bounces
Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    If (vbUnchecked = bounceSelection(servoIndex).Value) Then
        servoBounceOff servoIndex
    End If
Next

' Extended travel only supported from Sema4c onwards
disableExtendedTravel

' Select first setting option control and display corresponding value
selectSetting 0

End Sub

Private Sub setSema4bCompatabillity()

' Sema4b is a derivative of Sema4
setSema4Compatabillity

' Update compatability mode display
compatabilityText.Caption = sema4bText
disableCurrentCompatabilitySelection

initialiseSema4bSettingCommands

' Select first setting option control and display corresponding value
selectSetting 0

End Sub

Private Sub setSema4cCompatabillity()

' Sema4c is a derivative of Sema4b
setSema4bCompatabillity

' Update compatability mode display
compatabilityText.Caption = sema4cText
disableCurrentCompatabilitySelection

' Enable the selection controls to select extended servo travel supported
' from Sema4c onwards
enableExtendedTravel

' Select first setting option control and display corresponding value
selectSetting 0

End Sub

Private Sub setSema4dCompatabillity()

' Sema4d is a derivative of Sema4c
setSema4cCompatabillity

' Update compatability mode display
compatabilityText.Caption = sema4dText
disableCurrentCompatabilitySelection

' Sema4d speeds run 0 - fastest to 255 - slowest, previous versions were
' 0 - fastest then 1 - slowest to 255 - fastest
flipSpeeds

' Select first setting option control and display corresponding value
selectSetting 0

End Sub

Private Sub setOffline()
' Allow editing of settings whilst not connected to any hardware

runMode = OFFLINE

' Disable selection of Run or Set mode
runButton.Enabled = False
setButton.Enabled = False

disableSendStoreReset
enableChangingSettingValue
disableCurrentCompatabilitySelection
enableExtendedTravelSelections
enableBounceSelection

End Sub

Private Sub setRunningMode()
' Module is allowed to run normally without any settings being sent unless
' the Set All button is pressed

runMode = RUNNING

If (comPort.PortOpen) Then
    If (servo4Text <> compatabilityText.Caption) Then
        ' Ensure module is not in Set mode
        sendCommand (RUN_COMMAND)
    End If

    ' Disable selection of Run mode, enable selection of Set mode
    runButton.Enabled = False
    setButton.Enabled = True

    enableSendStoreReset
    disableChangingSettingValue
    disableCurrentCompatabilitySelection
    disableExtendedTravelSelections
    disableBounceSelection

Else
    ' COM port not available so act just as an offline settings editor
    setOffline
End If

End Sub

Private Sub setSettingMode()
' Currently selected option value is continually streamed to module allowing
' effect of change to be seen immediately

runMode = SETTING

runButton.Enabled = True
setButton.Enabled = False

disableSendStoreReset
enableChangingSettingValue
disableAllCompatabilitySelections
enableExtendedTravelSelections
enableBounceSelection

End Sub

Private Sub disableAllCompatabilitySelections()

optCompatServo4MenuItem.Enabled = False
optCompatSema4MenuItem.Enabled = False
optCompatSema4bMenuItem.Enabled = False
optCompatSema4cMenuItem.Enabled = False
optCompatSema4dMenuItem.Enabled = False

End Sub

Private Sub enableAllCompatabilitySelections()

optCompatServo4MenuItem.Enabled = True
optCompatSema4MenuItem.Enabled = True
optCompatSema4bMenuItem.Enabled = True
optCompatSema4cMenuItem.Enabled = True
optCompatSema4dMenuItem.Enabled = True

End Sub

Private Sub disableCurrentCompatabilitySelection()

enableAllCompatabilitySelections

' Disable selection of current compatability
Select Case compatabilityText.Caption
    Case sema4Text
        optCompatSema4MenuItem.Enabled = False
    Case sema4bText
        optCompatSema4bMenuItem.Enabled = False
    Case sema4cText
        optCompatSema4cMenuItem.Enabled = False
    Case sema4dText
        optCompatSema4dMenuItem.Enabled = False
    Case Else
        optCompatServo4MenuItem.Enabled = False
End Select

End Sub

Private Sub disableSendStoreReset()

setallButton.Enabled = False
storeButton.Enabled = False
resetButton.Enabled = False

End Sub

Private Sub enableSendStoreReset()

setallButton.Enabled = True
storeButton.Enabled = True
resetButton.Enabled = True

End Sub

Private Sub disableChangingSettingValue()

centerButton.Enabled = False
valueScroller.Enabled = False
valueText.Enabled = False

End Sub

Private Sub enableChangingSettingValue()

centerButton.Enabled = True
valueScroller.Enabled = True
valueText.Enabled = True

End Sub

Private Sub disableExtendedTravel()

xtndTravelSelectionGroup.Visible = False
xtndTravelLabel.Enabled = False

End Sub

Private Sub enableExtendedTravel()

If ((sema4cText = compatabilityText.Caption) Or _
    (sema4dText = compatabilityText.Caption)) Then
    xtndTravelSelectionGroup.Visible = True
    xtndTravelLabel.Enabled = True
End If

End Sub

Private Sub disableExtendedTravelSelections()

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    xtndTravelSelection(servoIndex).Enabled = False
Next

xtndTravelSelectionGroup.Enabled = False

End Sub

Private Sub enableExtendedTravelSelections()

If ((sema4cText = compatabilityText.Caption) Or _
    (sema4dText = compatabilityText.Caption)) Then
    Dim servoIndex As Integer
    For servoIndex = 0 To NUM_SERVOS - 1
        xtndTravelSelection(servoIndex).Enabled = True
    Next

    xtndTravelSelectionGroup.Enabled = True
End If

End Sub

Private Sub servoBounceAvailable(servoIndex As Integer, _
                                 available As Boolean)

Dim offPositionIndex As Integer
Dim onPositionIndex As Integer

offPositionIndex = SERVO_SETTINGS * servoIndex
onPositionIndex = offPositionIndex + 1

Dim offBounceIndex As Integer
Dim onBounceIndex As Integer

offBounceIndex = SERVO4_SETTINGS + (SEMA_SETTINGS * servoIndex)
onBounceIndex = offBounceIndex + (SEMA_SETTINGS / 2)

Dim bounceOffset As Integer
For bounceOffset = 0 To (SEMA_SETTINGS / 2) - 1
    servoSettingOption(offBounceIndex + bounceOffset).Visible = available
    servoSettingOption(offBounceIndex + bounceOffset).Enabled = available
    servoSettingOption(onBounceIndex + bounceOffset).Visible = available
    servoSettingOption(onBounceIndex + bounceOffset).Enabled = available

    If (Not available) Then
        If (settingIndex = (offBounceIndex + bounceOffset)) Then
            selectSetting offPositionIndex
        ElseIf (settingIndex = (onBounceIndex + bounceOffset)) Then
            selectSetting onPositionIndex
        End If
    End If
Next

End Sub

Private Sub disableBounce()

deselectAllBounces

offBounce3Label.Enabled = False
offBounce2Label.Enabled = False
offBounce1Label.Enabled = False
onBounce1Label.Enabled = False
onBounce2Label.Enabled = False
onBounce3Label.Enabled = False
bounceLabel.Enabled = False

bounceSelectionGroup.Visible = False

End Sub

Private Sub enableBounce()

Dim servoIndex As Integer
If (servo4Text <> compatabilityText.Caption) Then
    checkBounceSelections

    offBounce3Label.Enabled = True
    offBounce2Label.Enabled = True
    offBounce1Label.Enabled = True
    onBounce1Label.Enabled = True
    onBounce2Label.Enabled = True
    onBounce3Label.Enabled = True
    bounceLabel.Enabled = True

    bounceSelectionGroup.Visible = True
End If

End Sub

Private Sub selectBounce(servoIndex As Integer)

servoBounceOn servoIndex
servoBounceAvailable servoIndex, True
bounceSelection(servoIndex).Value = vbChecked

End Sub

Private Sub deselectBounce(servoIndex As Integer)

servoBounceOff servoIndex
servoBounceAvailable servoIndex, False
bounceSelection(servoIndex).Value = vbUnchecked

End Sub

Private Sub selectAllBounces()

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    selectBounce servoIndex
Next

End Sub

Private Sub deselectAllBounces()

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    deselectBounce servoIndex
Next

End Sub

Private Sub checkBounceSelections()

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    servoBounceAvailable servoIndex, _
                         (vbChecked = bounceSelection(servoIndex).Value)
Next

End Sub

Private Sub disableBounceSelection()

Dim servoIndex As Integer
For servoIndex = 0 To NUM_SERVOS - 1
    bounceSelection(servoIndex).Enabled = False
Next

bounceSelectionGroup.Enabled = False

End Sub

Private Sub enableBounceSelection()

If (servo4Text <> compatabilityText.Caption) Then
    Dim servoIndex As Integer
    For servoIndex = 0 To NUM_SERVOS - 1
        bounceSelection(servoIndex).Enabled = True
    Next

    bounceSelectionGroup.Enabled = True
End If

End Sub

Private Sub setComPort(Optional oldComPortNumber As Integer = 1)

Dim newComPortName As String
newComPortName = selectComPort(oldComPortNumber)
If ("Offline" = newComPortName) Then
    setOffline
Else
    setRunningMode
End If
connectionText.Caption = newComPortName

End Sub

Private Sub sendAllSettings()
' Download all the current settings

sema4SetForm.MousePointer = vbHourglass

Dim sendIndex As Integer

' Walk the array of setting values
For sendIndex = LBound(settingValue) To UBound(settingValue)
    ' Test if option button for setting is enabled
    If (servoSettingOption(sendIndex).Enabled) Then
        ' Send setting command and value
        sendSetting sendIndex
    End If
Next

If ((sema4cText = compatabilityText.Caption) Or _
    (sema4dText = compatabilityText.Caption)) Then
    sendSettingValue TRVL_SETTING, getExtendedTravelSelections
End If

If (servo4Text <> compatabilityText.Caption) Then
    ' Ensure module leaves Set mode after download
    sendCommand RUN_COMMAND
End If

' Select first setting option control and display corresponding value
selectSetting 0

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub limitSettingValue(testValue As Integer)

If (testValue > valueScroller.Max) Then
    testValue = valueScroller.Max
End If
If (testValue < valueScroller.Min) Then
    testValue = valueScroller.Min
End If

End Sub

Private Sub changeSettingValue(newValue As Integer)

limitSettingValue newValue

If (newValue <> settingValue(settingIndex)) Then
    settingValue(settingIndex) = newValue
    valueScroller.Value = newValue
    valueText.Text = newValue
    sendSetting settingIndex
    settingsChanged = True
End If

End Sub

Private Sub selectSetting(newSettingIndex As Integer)

settingIndex = newSettingIndex

valueScroller.Max = 255
valueScroller.LargeChange = 18

If (servo4Text = compatabilityText.Caption) Then
    Select Case settingIndex

    Case ON_SPD_NDX_1, OFF_SPD_NDX_1, ON_SPD_NDX_2, OFF_SPD_NDX_2, _
         ON_SPD_NDX_3, OFF_SPD_NDX_3, ON_SPD_NDX_4, OFF_SPD_NDX_4
    
        ' Selected setting is a speed limit for Servo4
        valueScroller.Value = settingValue(settingIndex)
        valueScroller.Max = SERVO4_MAX_SPEED
        valueScroller.LargeChange = 1
    End Select
End If
    
limitSettingValue settingValue(settingIndex)

valueScroller.Value = settingValue(settingIndex)
valueText.Text = settingValue(settingIndex)

servoSettingOption(settingIndex).Value = True

End Sub

Private Sub Form_Load()
' Initialisation when form is first loaded

' Export references to control referenced by other modules
Set sema4Port = comPort

' Set up File Dialog filter to match Sema4 settings files
settingsFileDialog.Filter = "Sema4Set Files (*.sm4)|*.sm4" _
                            + "|Text Files (*.txt)|*.txt" _
                            + "!All Files (*.*)|*.*"
settingsFileDialog.FilterIndex = 1

setRunningMode

settingsFilename = ""

setSema4dCompatabillity

' Prevent newSettings from prompting to save current settings
settingsChanged = False

newSettings

setComPort

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

setComPort comPort.commport

End Sub

Private Sub optCompatServo4MenuItem_Click()

setServo4Compatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4MenuItem_Click()

setSema4Compatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4bMenuItem_Click()

setSema4bCompatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4cMenuItem_Click()

setSema4cCompatabillity
settingsChanged = True

End Sub

Private Sub optCompatSema4dMenuItem_Click()

setSema4dCompatabillity
settingsChanged = True

End Sub

Private Sub runButton_Click()

' Allow module to run normally
setRunningMode

End Sub

Private Sub setallButton_Click()

' Download all current setting values to module
sendAllSettings

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

If (settingValue(settingIndex) <> _
       ((valueScroller.Max - valueScroller.Min) / 2)) Then
    changeSettingValue ((valueScroller.Max - valueScroller.Min) / 2)
End If

End Sub

Private Sub servoSettingOption_Click(optionIndex As Integer)

If (settingIndex <> optionIndex) Then
    selectSetting optionIndex
End If

End Sub

Private Sub valueScroller_Change()

If (settingValue(settingIndex) <> valueScroller.Value) Then
    changeSettingValue valueScroller.Value
End If

End Sub

Private Sub valueText_KeyUp(keyCode As Integer, shift As Integer)

If ((RTN_KEYCODE = keyCode) And _
    (settingValue(settingIndex) <> CInt(Val(valueText.Text)))) Then
    changeSettingValue CInt(Val(valueText.Text))
End If

End Sub

Private Sub valueText_LostFocus()

If (settingValue(settingIndex) <> CInt(Val(valueText.Text))) Then
    changeSettingValue CInt(Val(valueText.Text))
End If

End Sub

Private Sub xtndTravelSelection_Click(Index As Integer)

sendSettingValue TRVL_SETTING, getExtendedTravelSelections
settingsChanged = True

End Sub

Private Sub bounceSelection_Click(Index As Integer)

If (vbChecked = bounceSelection(Index).Value) Then
    selectBounce Index
Else
    deselectBounce Index
End If

settingsChanged = True

End Sub
