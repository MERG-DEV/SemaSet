Attribute VB_Name = "settings"
Option Explicit

' Default value to assign to new setting value
Public Const DEFAULT_SETTING As Integer = 127

' Maximum speed for Servo4
Public Const SERVO4_MAX_SPEED As Integer = 7

' Number of Servos
Public Const NUM_SERVOS As Integer = 4

' Number of setting values for Servo (on and off, speed and position)
Public Const SERVO_SETTINGS As Integer = 4

' Number of setting values for Servo4
Public Const SERVO4_SETTINGS As Integer = NUM_SERVOS * SERVO_SETTINGS

' Number of extra setting values for Sema (3 off and 3 on bounces)
Public Const SEMA_SETTINGS  As Integer = 6

' Number of setting values for Sema4
Public Const SEMA4_SETTINGS  As Integer = _
    NUM_SERVOS * (SEMA_SETTINGS + SERVO_SETTINGS)

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
Public settingValue(0 To (SEMA4_SETTINGS - 1))   As Integer
Public settingCommand(0 To (SEMA4_SETTINGS - 1)) As Integer

' Define array index constants for certain values
Public Const OFF_PSTN_NDX_1 As Integer = 0
Public Const ON_PSTN_NDX_1  As Integer = 1
Public Const OFF_SPD_NDX_1  As Integer = 2
Public Const ON_SPD_NDX_1   As Integer = 3
Public Const OFF_PSTN_NDX_2 As Integer = 4
Public Const ON_PSTN_NDX_2  As Integer = 5
Public Const OFF_SPD_NDX_2  As Integer = 6
Public Const ON_SPD_NDX_2   As Integer = 7
Public Const OFF_PSTN_NDX_3 As Integer = 8
Public Const ON_PSTN_NDX_3  As Integer = 9
Public Const OFF_SPD_NDX_3  As Integer = 10
Public Const ON_SPD_NDX_3   As Integer = 11
Public Const OFF_PSTN_NDX_4 As Integer = 12
Public Const ON_PSTN_NDX_4  As Integer = 13
Public Const OFF_SPD_NDX_4  As Integer = 14
Public Const ON_SPD_NDX_4   As Integer = 15

Public settingIndex As Integer
Public settingsChanged As Boolean

' Extended travel options byte bitmasks
Public Const SRV1_XTND_MASK As Integer = &H4
Public Const SRV2_XTND_MASK As Integer = &H8
Public Const SRV3_XTND_MASK As Integer = &H40
Public Const SRV4_XTND_MASK As Integer = &H80

' Settings file format version
Public Const SETTINGS_FILE_FORMAT_VERSION As Integer = 0

Public settingsFilename As String

' Transmitted command characters
Public Const COMMAND_BASE  As Integer = 65 ' ASCII A
Public Const TRVL_SETTING  As Integer = 56
Public Const STORE_COMMAND As String = "@"
Public Const RESET_COMMAND As String = "#"
Public Const RUN_COMMAND   As String = "$"

' Reference to current operating mode
Public runMode As OperatingMode

Public Sub sendSettingValue(settingCommand As Integer, _
                            Optional commandValue As Integer = 0)

sendCommand Chr(COMMAND_BASE + settingCommand), commandValue
                        
End Sub

Public Sub sendSetting(sendIndex As Integer)

sendSettingValue settingCommand(settingIndex), settingValue(settingIndex)
                        
End Sub

Public Sub streamCurrentSetting()
' Continuosly send the currently selected setting value so the module tracks
' changes interactively

While (SETTING = runMode)
    ' Send setting command and value for currently selected setting
    sendSetting settingIndex
Wend

End Sub

Private Function toServo4Speed(sema4Speed As Integer) As Integer

toServo4Speed = 0

If (0 < sema4Speed) Then
    toServo4Speed = sema4Speed / 16

    If (1 > toServo4Speed) Then
        toServo4Speed = 1
    End If

    If (SERVO4_MAX_SPEED < toServo4Speed) Then
        toServo4Speed = SERVO4_MAX_SPEED
    End If
End If

End Function

Public Sub convertSpeedToServo4()

settingValue(OFF_SPD_NDX_1) = toServo4Speed(settingValue(OFF_SPD_NDX_1))
settingValue(ON_SPD_NDX_1) = toServo4Speed(settingValue(ON_SPD_NDX_1))
settingValue(OFF_SPD_NDX_2) = toServo4Speed(settingValue(OFF_SPD_NDX_2))
settingValue(ON_SPD_NDX_2) = toServo4Speed(settingValue(ON_SPD_NDX_2))
settingValue(OFF_SPD_NDX_3) = toServo4Speed(settingValue(OFF_SPD_NDX_3))
settingValue(ON_SPD_NDX_3) = toServo4Speed(settingValue(ON_SPD_NDX_3))
settingValue(OFF_SPD_NDX_4) = toServo4Speed(settingValue(OFF_SPD_NDX_4))
settingValue(ON_SPD_NDX_4) = toServo4Speed(settingValue(ON_SPD_NDX_4))

End Sub

Public Sub convertSpeedFromServo4()

settingValue(OFF_SPD_NDX_1) = settingValue(OFF_SPD_NDX_1) * 16
settingValue(ON_SPD_NDX_1) = settingValue(ON_SPD_NDX_1) * 16
settingValue(OFF_SPD_NDX_2) = settingValue(OFF_SPD_NDX_2) * 16
settingValue(ON_SPD_NDX_2) = settingValue(ON_SPD_NDX_2) * 16
settingValue(OFF_SPD_NDX_3) = settingValue(OFF_SPD_NDX_3) * 16
settingValue(ON_SPD_NDX_3) = settingValue(ON_SPD_NDX_3) * 16
settingValue(OFF_SPD_NDX_4) = settingValue(OFF_SPD_NDX_4) * 16
settingValue(ON_SPD_NDX_4) = settingValue(ON_SPD_NDX_4) * 16

End Sub

Private Function flipSpeed(speed As Integer) As Integer

flipSpeed = 0

If (0 < speed) Then
    flipSpeed = 256 - speed
End If

End Function

Public Sub flipSpeeds()

settingValue(OFF_SPD_NDX_1) = flipSpeed(settingValue(OFF_SPD_NDX_1))
settingValue(ON_SPD_NDX_1) = flipSpeed(settingValue(ON_SPD_NDX_1))
settingValue(OFF_SPD_NDX_2) = flipSpeed(settingValue(OFF_SPD_NDX_2))
settingValue(ON_SPD_NDX_2) = flipSpeed(settingValue(ON_SPD_NDX_2))
settingValue(OFF_SPD_NDX_3) = flipSpeed(settingValue(OFF_SPD_NDX_3))
settingValue(ON_SPD_NDX_3) = flipSpeed(settingValue(ON_SPD_NDX_3))
settingValue(OFF_SPD_NDX_4) = flipSpeed(settingValue(OFF_SPD_NDX_4))
settingValue(ON_SPD_NDX_4) = flipSpeed(settingValue(ON_SPD_NDX_4))

End Sub

Public Sub initialiseServo4SettingCommands()

' Initialise command for each setting
For settingIndex = LBound(settingCommand) To UBound(settingCommand)
    settingCommand(settingIndex) = settingIndex
Next

End Sub

Public Sub initialiseSema4SettingCommands()

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

Public Sub initialiseSema4bSettingCommands()

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

Public Sub servoBounceOff(servoIndex As Integer)

Dim offPositionIndex As Integer
Dim onPositionIndex As Integer

offPositionIndex = SERVO_SETTINGS * servoIndex
onPositionIndex = offPositionIndex + 1

Dim offBounceIndex As Integer
Dim onBounceIndex As Integer

offBounceIndex = SERVO4_SETTINGS + (SEMA_SETTINGS * servoIndex)
onBounceIndex = offBounceIndex + (SEMA_SETTINGS / 2)

Dim bounceNumber As Integer
For bounceNumber = 0 To (SEMA_SETTINGS / 2) - 1
    settingValue(offBounceIndex + bounceNumber) = _
        settingValue(offPositionIndex)
    settingValue(onBounceIndex + bounceNumber) = _
        settingValue(onPositionIndex)
Next

settingCommand(offPositionIndex) = offPositionIndex
settingCommand(onPositionIndex) = onPositionIndex

End Sub

Public Sub servoBounceOn(servoIndex As Integer)

servoBounceOff (servoIndex)

Dim offPositionIndex As Integer
Dim command As Integer

offPositionIndex = SERVO_SETTINGS * servoIndex

command = 40 + (servoIndex * 2)

settingCommand(offPositionIndex) = command
settingCommand(offPositionIndex + 1) = command + 1

End Sub
