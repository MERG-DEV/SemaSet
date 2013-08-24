Attribute VB_Name = "commport"
Option Explicit

Private comPort As MSComm

Public Sub setComPort(newComPort As MSComm)

Set comPort = newComPort

End Sub

Public Sub openComPort(newComPortNumber As Integer)

On Error GoTo commerror

' If COM port currently open close it
If True = comPort.PortOpen Then
    comPort.PortOpen = False
End If

' Set new COM port number and open COM port
comPort.commport = newComPortNumber

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
    sema4SetForm.connectionText.Caption = "Offline"
    sema4SetForm.setOffline
Else
    sema4SetForm.connectionText.Caption = "COM" + newComPortName
    openComPort (newComPortNumber)

    sema4SetForm.setRunningMode
End If

End Sub

Public Sub changeComPort()

selectComPort comPort.commport

End Sub

Public Sub comPortFailed()

MsgBox "Error accessing COM port, " + Error, vbOKOnly, "COM Port Error"

changeComPort

End Sub
