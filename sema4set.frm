VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sema4SetForm 
   BorderStyle     =   0  'None
   Caption         =   "Sem4Set"
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
   Begin VB.CommandButton setallButton 
      Caption         =   "Set &All"
      Height          =   375
      Left            =   2880
      TabIndex        =   65
      Top             =   1560
      Width           =   975
   End
   Begin MSComDlg.CommonDialog settingsFileDialog 
      Left            =   3600
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame servoSettingOptionGroup 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1455
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   41
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   37
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   27
         Left            =   480
         TabIndex        =   36
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   35
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   26
         Left            =   480
         TabIndex        =   34
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   33
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   31
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   23
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   28
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   27
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   26
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   25
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   32
         Left            =   840
         TabIndex        =   24
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   30
         Left            =   840
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   22
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   31
         Left            =   840
         TabIndex        =   21
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   29
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   28
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   18
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   15
         Left            =   1200
         TabIndex        =   17
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   39
         Left            =   1200
         TabIndex        =   16
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   15
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   38
         Left            =   1200
         TabIndex        =   14
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   36
         Left            =   1200
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   12
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   37
         Left            =   1200
         TabIndex        =   11
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   35
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton servoSettingOption 
         Height          =   255
         Index           =   13
         Left            =   1200
         TabIndex        =   8
         Top             =   2160
         Width           =   255
      End
   End
   Begin VB.CommandButton setButton 
      Caption         =   "&Set"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton resetButton 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3480
      Width           =   975
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
   Begin VB.CommandButton storeButton 
      Caption         =   "S&tore"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton runButton 
      Caption         =   "&Run"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin MSCommLib.MSComm comPort 
      Left            =   3000
      Top             =   -480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   128
      OutBufferSize   =   128
   End
   Begin VB.Label connectionText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "Com1"
      Height          =   255
      Left            =   3360
      TabIndex        =   67
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label connectionLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Connection:"
      Height          =   195
      Left            =   3000
      TabIndex        =   66
      Top             =   960
      Width           =   855
   End
   Begin VB.Label compatabilityLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Compatability:"
      Height          =   195
      Left            =   2880
      TabIndex        =   64
      Top             =   360
      Width           =   975
   End
   Begin VB.Label compatabilityText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "Sema4"
      Height          =   255
      Left            =   3360
      TabIndex        =   63
      Top             =   600
      Width           =   495
   End
   Begin VB.Label servo4Label 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   195
      Left            =   2190
      TabIndex        =   62
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo1Label 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   195
      Left            =   1110
      TabIndex        =   61
      Top             =   120
      Width           =   285
   End
   Begin VB.Label onBounce3Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 3"
      Height          =   195
      Left            =   0
      TabIndex        =   60
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label onBounce2Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 2"
      Height          =   195
      Left            =   0
      TabIndex        =   59
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label onBounce1Label 
      AutoSize        =   -1  'True
      Caption         =   "On Bounce 1"
      Height          =   195
      Left            =   0
      TabIndex        =   58
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label offBounce3Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 3"
      Height          =   195
      Left            =   0
      TabIndex        =   57
      Top             =   480
      Width           =   945
   End
   Begin VB.Label offBounce2Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 2"
      Height          =   195
      Left            =   0
      TabIndex        =   56
      Top             =   840
      Width           =   945
   End
   Begin VB.Label offBounce1Label 
      AutoSize        =   -1  'True
      Caption         =   "Off Bounce 1"
      Height          =   195
      Left            =   0
      TabIndex        =   55
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label onSpeedLabel 
      AutoSize        =   -1  'True
      Caption         =   "OnSpeed"
      Height          =   195
      Left            =   270
      TabIndex        =   54
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label offSpeedLabel 
      AutoSize        =   -1  'True
      Caption         =   "Off Speed"
      Height          =   195
      Left            =   225
      TabIndex        =   53
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label onPositionLabel 
      AutoSize        =   -1  'True
      Caption         =   "On Position"
      Height          =   195
      Left            =   135
      TabIndex        =   52
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label offPositionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Off Position"
      Height          =   195
      Left            =   135
      TabIndex        =   51
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label servo2Label 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   195
      Left            =   1470
      TabIndex        =   50
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servo3Label 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   195
      Left            =   1830
      TabIndex        =   49
      Top             =   120
      Width           =   285
   End
   Begin VB.Label servosLabel 
      AutoSize        =   -1  'True
      Caption         =   "Servo:"
      Height          =   195
      Left            =   480
      TabIndex        =   48
      Top             =   120
      Width           =   465
   End
   Begin VB.Label valueText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label valueLabel 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4200
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
            Caption         =   "Se&ma4"
         End
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
    running
    setting
End Enum

' Settings file format version
Const SETTINGS_FILE_FORMAT_VERSION As Integer = 0

' Number of setting values for Servo4
Const SERVO4_SETTINGS As Integer = 16

' Number of setting values for Sema4
Const SEMA4_SETTINGS  As Integer = 40

' Number of times to send a non streaming command or setting string
Const SEND_ITTERATIONS As Integer = 20

' Default value to assign to new setting value
Const DEFAULT_SETTING As Integer = 127

' Transmitted command characters
Const SYNCH_BYTE    As Integer = 0  ' ASCII null
Const SETTING_BASE  As Integer = 65 ' ASCII A
Const STORE_COMMAND As String = "@"
Const RESET_COMMAND As String = "#"

' Compatability names
Const servo4CompatabilityText As String = "Servo4"
Const sema4CompatabilityText  As String = "Sema4"

Dim settingValue(0 To (SEMA4_SETTINGS - 1))  As Integer
Dim settingLookup(0 To (SEMA4_SETTINGS - 1)) As Integer

Dim settingIndex    As Integer
Dim settingsChanged As Boolean

Dim settingsFilename As String

Dim currentMode     As OperatingMode

Private Sub openComPort(newComPortNumber As Integer)

If True = comPort.PortOpen Then
    comPort.PortOpen = False
End If

comPort.CommPort = newComPortNumber
comPort.PortOpen = True

End Sub

Public Sub selectComPort(oldComPortNumber As Integer)

Dim newComPortName As String
Dim newComPortNumber As Integer

' Prompt user to select COM port connected to Servo4
newComPortName = InputBox("Select COM Port", _
                          "COM Port number", _
                          oldComPortNumber)

' Convert entered COM Port string to an integer value
newComPortNumber = Val(newComPortName)

If (1 > newComPortNumber) Then
    End
End If

connectionText.Caption = "Com" + newComPortName

openComPort (newComPortNumber)

End Sub

Private Sub changeComPort()

selectComPort comPort.CommPort

End Sub

Private Sub comPortFailed()

MsgBox "Error accessing COM port, " + Error, vbOKOnly, "COM Port Error"

changeComPort

End Sub

Private Sub setRunningMode()

currentMode = running

If compatabilityText.Caption = servo4CompatabilityText Then
    optCompatSema4MenuItem.Enabled = True
    optCompatServo4MenuItem.Enabled = False
Else
    optCompatSema4MenuItem.Enabled = False
    optCompatServo4MenuItem.Enabled = True
End If

runButton.Enabled = False
setButton.Enabled = True

valueScroller.Enabled = False

End Sub

Private Sub setSettingMode()

currentMode = setting

optCompatSema4MenuItem.Enabled = False
optCompatServo4MenuItem.Enabled = False

runButton.Enabled = True
setButton.Enabled = False

valueScroller.Enabled = True

End Sub

Private Sub streamCurrentSetting()

On Error GoTo comPortFailure

While (setting = currentMode)
    ' Perform event dispatch to keep GUI alive, allows currentMode to be changed
    DoEvents

    ' Send setting message for currently selected setting and value
    comPort.Output = Chr(SYNCH_BYTE) _
                     + Chr(SETTING_BASE + settingLookup(settingIndex)) _
                     + Format(settingValue(settingIndex), "000")
Wend

Exit Sub

comPortFailure:
    comPortFailed

End Sub

Private Sub sendCommand(commandCharacter As String, _
                        Optional commandValue As Integer = 0)

On Error GoTo comPortFailure

Dim n As Integer

For n = 1 To SEND_ITTERATIONS
    ' Perform event dispatch to keep GUI alive, allows currentMode to be changed
    DoEvents

    ' Send command message
    comPort.Output = Chr(SYNCH_BYTE) _
                     + commandCharacter _
                     + Format(commandValue, "000")
Next

Exit Sub

comPortFailure:
    comPortFailed

End Sub

Private Sub sendCurrentSettings()

sema4SetForm.MousePointer = vbHourglass

Dim sendIndex As Integer

For sendIndex = LBound(settingValue) To UBound(settingValue)
    If servoSettingOption(sendIndex).Enabled Then
        sendCommand Chr(SETTING_BASE + settingLookup(sendIndex)), _
                    settingValue(sendIndex)
    End If
Next

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub checkIfSaveNeeded()

If settingsChanged Then
    If vbYes = MsgBox("Settings have changed, save before overwriting?", vbYesNo) Then
        saveSettings
    End If
End If

End Sub

Private Sub newSettings()

setRunningMode

checkIfSaveNeeded

For settingIndex = LBound(settingValue) To UBound(settingValue)
    settingValue(settingIndex) = DEFAULT_SETTING
Next

settingIndex = 0
servoSettingOption(settingIndex).Value = True
valueScroller.Value = settingValue(settingIndex)

settingsChanged = False

End Sub

Private Sub loadSettings()

setRunningMode

checkIfSaveNeeded

On Error GoTo errorCancel

settingsFileDialog.ShowOpen
settingsFilename = settingsFileDialog.FileName

If "" = settingsFilename Then
    GoTo errorCancel
End If

Open settingsFilename For Input As #1

Dim loadedCompatabilityText As String

Input #1, loadedCompatabilityText

If loadedCompatabilityText <> compatabilityText.Caption Then
    If loadedCompatabilityText = servo4CompatabilityText Then
        setServo4Compatabillity
    Else
        setSema4Compatabillity
    End If
End If

Dim loadedFileFormatVersion As Integer

Input #1, loadedFileFormatVersion

settingIndex = LBound(settingValue)

Do Until (EOF(1) Or (UBound(settingValue) < settingIndex))
    Input #1, settingValue(settingIndex)
    settingIndex = 1 + settingIndex
Loop

Close #1

settingIndex = 0
servoSettingOption(settingIndex).Value = True
valueScroller.Value = settingValue(settingIndex)

settingsChanged = False

errorCancel:

End Sub

Private Sub saveSettings()

On Error GoTo errorCancel

If "" = settingsFilename Then
    settingsFileDialog.ShowSave
    settingsFilename = settingsFileDialog.FileName
End If

If "" = settingsFilename Then
    MsgBox "Filename blank, settings not saved", vbOKOnly, "No filename"
End If

Open settingsFilename For Output As #1

Print #1, compatabilityText.Caption
Print #1, SETTINGS_FILE_FORMAT_VERSION

Dim outputIndex As Integer

For outputIndex = LBound(settingValue) To UBound(settingValue)
    Print #1, settingValue(outputIndex)
Next

Close #1

settingsChanged = False

errorCancel:

End Sub

Private Sub setSema4Compatabillity()

For settingIndex = LBound(settingLookup) To UBound(settingLookup)
    settingLookup(settingIndex) = settingIndex
Next

settingLookup(0) = 40
settingLookup(1) = 41
settingLookup(4) = 42
settingLookup(5) = 43
settingLookup(8) = 44
settingLookup(9) = 45
settingLookup(12) = 46
settingLookup(13) = 47

For settingIndex = SERVO4_SETTINGS To (SEMA4_SETTINGS - 1)
    servoSettingOption(settingIndex).Visible = True
    servoSettingOption(settingIndex).Enabled = True
Next

settingIndex = 0
servoSettingOption(settingIndex).Value = True
valueScroller.Value = settingValue(settingIndex)

compatabilityText.Caption = sema4CompatabilityText
optCompatSema4MenuItem.Enabled = False
optCompatServo4MenuItem.Enabled = True

End Sub

Private Sub setServo4Compatabillity()

For settingIndex = LBound(settingLookup) To UBound(settingLookup)
    settingLookup(settingIndex) = settingIndex
Next

For settingIndex = SERVO4_SETTINGS To (SEMA4_SETTINGS - 1)
    servoSettingOption(settingIndex).Visible = False
    servoSettingOption(settingIndex).Enabled = False
Next

settingIndex = 0
servoSettingOption(settingIndex).Value = True
valueScroller.Value = settingValue(settingIndex)

compatabilityText.Caption = servo4CompatabilityText
optCompatSema4MenuItem.Enabled = True
optCompatServo4MenuItem.Enabled = False

End Sub

Private Sub Form_Load()

settingsFileDialog.Filter = "Sema4Set Files (*.sm4)|*.sm4" _
                            + "|Text Files (*.txt)|*.txt" _
                            + "!All Files (*.*)|*.*"
settingsFileDialog.FilterIndex = 1

settingsFilename = ""

selectComPort 1

settingsChanged = False
newSettings

setSema4Compatabillity

Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

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

End

End Sub

Private Sub optSerPortMenuItem_Click()

changeComPort

End Sub

Private Sub optCompatServo4MenuItem_Click()

setServo4Compatabillity

End Sub

Private Sub optCompatSema4MenuItem_Click()

setSema4Compatabillity

End Sub

Private Sub runButton_Click()

setRunningMode

End Sub

Private Sub setallButton_Click()

sendCurrentSettings

End Sub

Private Sub setButton_Click()

setSettingMode
streamCurrentSetting

End Sub

Private Sub storeButton_Click()

sema4SetForm.MousePointer = vbHourglass
setRunningMode
sendCommand (STORE_COMMAND)
sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub resetButton_Click()

sema4SetForm.MousePointer = vbHourglass
setRunningMode
sendCommand (RESET_COMMAND)
sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub servoSettingOption_Click(Index As Integer)

settingIndex = Index
valueScroller.Value = settingValue(settingIndex)

End Sub

Private Sub valueScroller_Change()

settingValue(settingIndex) = valueScroller.Value
valueText.Caption = settingValue(settingIndex)
settingsChanged = True

End Sub
