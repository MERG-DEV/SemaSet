Attribute VB_Name = "commport"
Option Explicit

Private sema4Port As MSComm

Public Sub setComPort(newComPort As MSComm)

Set sema4Port = newComPort

End Sub

Public Sub openComPort(newComPortNumber As Integer)

On Error GoTo commerror

' If COM port currently open close it
If True = sema4Port.PortOpen Then
    sema4Port.PortOpen = False
End If

' Set new COM port number and open COM port
sema4Port.commport = newComPortNumber

sema4Port.PortOpen = True

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
    If True = sema4Port.PortOpen Then
        sema4Port.PortOpen = False
    End If
    sema4SetForm.connectionText.Caption = "Offline"
    sema4SetForm.setOffline
Else
    sema4SetForm.connectionText.Caption = "COM" + newComPortName
    openComPort (newComPortNumber)

    sema4SetForm.setRunningMode
End If

End Sub

Public Sub changeComPort()

selectComPort sema4Port.commport

End Sub

Public Sub streamCurrentSetting()
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

Public Sub sendCurrentSetting()

Dim n As Integer

' Send the currently selected setting value

If comPort.PortOpen Then
    On Error GoTo comPortFailure

    If (SETTING = currentMode) Then
        For n = 1 To SEND_ITTERATIONS
            ' Send setting command and value for currently selected setting
            comPort.Output = Chr(SYNCH_BYTE) _
                             + Chr(SETTING_BASE + settingCommand(settingIndex)) _
                             + Format(settingValue(settingIndex), "000")
        Next
    End If
End If

Exit Sub

comPortFailure:
    comPortFailed

End Sub

Public Sub sendCommand(commandCharacter As String, _
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

Public Sub sendSetting(settingCommand As Integer, _
                        Optional commandValue As Integer = 0, _
                        Optional sendItterations As Integer = SEND_ITTERATIONS)
' Send the command for a setting, and optionally a value for the command, repeatedly
' a set number of times to allow for garbled reception as link has no handshake

If (SETTING = currentMode) Then
    sendCommand Chr(SETTING_BASE + settingCommand), commandValue, sendItterations
End If
                        
End Sub

Public Sub sendCurrentSettings()
' Download all the current settings

sema4SetForm.MousePointer = vbHourglass

Dim sendIndex As Integer

' Walk the array of setting values
For sendIndex = LBound(settingCommand) To UBound(settingCommand)
    ' Test if option button for setting is enabled
    If servoSettingOption(sendIndex).Enabled Then
        ' Send setting command and value
        sendCommand Chr(SETTING_BASE + settingCommand(sendIndex)), _
                    settingValue(sendIndex)
    End If
Next

If sema4SetForm.compatabilityText.Caption = sema4cCompatabilityText Or _
   sema4SetForm.compatabilityText.Caption = sema4dCompatabilityText Then
    sendCommand Chr(SETTING_BASE + TRVL_SETTING), getExtendedTravelSelections
End If

If sema4SetForm.compatabilityText.Caption <> servo4CompatabilityText Then
    ' Ensure module leaves Set mode after download
    sendCommand (RUN_COMMAND)
End If

sema4SetForm.MousePointer = vbDefault

End Sub

Private Sub comPortFailed()

MsgBox "Error accessing COM port, " + Error, vbOKOnly, "COM Port Error"

changeComPort

End Sub
