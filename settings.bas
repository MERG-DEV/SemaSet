Attribute VB_Name = "settings"
Option Explicit

' Arrays for setting values and commands, layout:
'  Servo 1
'   Off Position, On Position, Off Speed, On Speed
'   [0]           [1]          [2]        [3]
'  Servo 2
'   Off Position, On Position, Off Speed, On Speed
'   [4]           [5]          [6]        [7]
'  Servo 3
'   Off Position, On Position, Off Speed, On Speed
'   [8]           [9]          [10]       [11]
'  Servo 4
'   Off Position, On Position, Off Speed, On Speed
'   [12]          [13]         [14]       [15]
'  Servo 1
'   Off Bounce 1, Off Bounce 2, Off Bounce 3
'   [16]          [17]          [18]
'   On Bounce 1,  On Bounce 2,  On Bounce 3
'   [19]          [20]          [21]
'  Servo 2
'   Off Bounce 1, Off Bounce 2, Off Bounce 3
'   [22]          [23]          [24]
'   On Bounce 1,  On Bounce 2,  On Bounce 3
'   [25]          [26]          [27]
'  Servo 3
'   Off Bounce 1, Off Bounce 2, Off Bounce 3
'   [28]          [29]          [30]
'   On Bounce 1,  On Bounce 2,  On Bounce 3
'   [31]          [32]          [33]
'  Servo 4
'   Off Bounce 1, Off Bounce 2, Off Bounce 3
'   [434          [35]          [36]
'   On Bounce 1,  On Bounce 2,  On Bounce 3
'   [37]          [38]          [39]
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

Public settingIndex    As Integer
Public settingsChanged As Boolean

Public settingsFilename As String

