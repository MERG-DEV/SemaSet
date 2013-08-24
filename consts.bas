Attribute VB_Name = "consts"
Option Explicit

Public Enum OperatingMode
    RUNNING
    SETTING
    OFFLINE
End Enum

' Key code to indicate completion of direct value input to valueText TextBox
Public Const RTN_KEYCODE As Integer = 13

' Settings file format version
Public Const SETTINGS_FILE_FORMAT_VERSION As Integer = 0

' Maximum speed for Servo4
Public Const SERVO4_MAX_SPEED As Integer = 7

' Number of setting values for Servo (on and off, speed and position)
Public Const SERVO_SETTINGS As Integer = 4

' Number of setting values for Servo4
Public Const SERVO4_SETTINGS As Integer = 4 * SERVO_SETTINGS

' Number of extra setting values for Sema (3 off and 3 on bounces)
Public Const SEMA_SETTINGS  As Integer = 6

' Number of setting values for Sema4
Public Const SEMA4_SETTINGS  As Integer = 4 * (SEMA_SETTINGS + SERVO_SETTINGS)

' Number of times to send a non streaming command or setting string
Public Const SEND_ITTERATIONS As Integer = 5

' Default value to assign to new setting value
Public Const DEFAULT_SETTING As Integer = 127

' Extended travel options byte bitmasks
Public Const SRV1_XTND_MASK As Integer = &H4
Public Const SRV2_XTND_MASK As Integer = &H8
Public Const SRV3_XTND_MASK As Integer = &H40
Public Const SRV4_XTND_MASK As Integer = &H80

' Transmitted command characters
Public Const SYNCH_BYTE    As Integer = 0  ' ASCII null
Public Const SETTING_BASE  As Integer = 65 ' ASCII A
Public Const TRVL_SETTING  As Integer = 56
Public Const STORE_COMMAND As String = "@"
Public Const RESET_COMMAND As String = "#"
Public Const RUN_COMMAND   As String = "$"

' Compatability names
Public Const servo4CompatabilityText As String = "Servo4"
Public Const sema4CompatabilityText  As String = "Sema4"
Public Const sema4bCompatabilityText As String = "Sema4b"
Public Const sema4cCompatabilityText As String = "Sema4c"
Public Const sema4dCompatabilityText As String = "Sema4d"


