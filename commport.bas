Attribute VB_Name = "commport"
Option Explicit

' Number of times to send a non streaming command or setting string
Public Const SEND_ITTERATIONS As Integer = 7

Public Const SYNCH_BYTE    As Integer = 0  ' ASCII null

' Reference to COM port control
Public sema4Port As MSComm

Public Function selectComPort(oldComPortNumber As Integer) As String

selectComPort = "Offline"

On Error GoTo commError

' If COM port is currently open close it
If (sema4Port.PortOpen) Then
    sema4Port.PortOpen = False
End If

' Prompt user to select COM port for connection
selectComPort = InputBox("Select COM Port (0 for Offline editing)", _
                         "COM Port number", _
                         oldComPortNumber)

' Convert entered COM Port string to an integer value
Dim newComPortNumber As Integer
newComPortNumber = CInt(Val(selectComPort))

' Ensure COM port number is greater than 0
If (1 > newComPortNumber) Then
    GoTo commOffline
Else
    ' Set new COM port number and open COM port
    sema4Port.commport = newComPortNumber
    sema4Port.PortOpen = True

    selectComPort = "COM" + selectComPort
End If

Exit Function

commError:

MsgBox "Unable to open the selected COM Port. Please choose another port.", _
       vbExclamation, _
       "COM Port Error"

commOffline:

selectComPort = "Offline"

End Function

Public Sub sema4PortFailed()

MsgBox "Error accessing COM port, " + Error, vbOKOnly, "COM Port Error"

selectComPort sema4Port.commport

End Sub

Public Sub sendCommand(commandCharacter As String, _
                       Optional commandValue As Integer = 0, _
                       Optional sendItterations As Integer = SEND_ITTERATIONS)
' Send the given command, and optionally a value for the command, repeatedly
' a set number of times to allow for garbled reception as link has no handshake

' Limit command value to an unsigned 8 bit integer (byte)
If (0 > commandValue) Then
    commandValue = 0
ElseIf (255 < commandValue) Then
    commandValue = 255
End If

Dim n As Integer

On Error GoTo comPortFailure

For n = 1 To sendItterations
    ' Perform event dispatch to keep GUI alive
    DoEvents

    If (sema4Port.PortOpen) Then
        ' Send command and value
        sema4Port.Output = Chr(SYNCH_BYTE) _
                           + commandCharacter _
                           + Format(commandValue, "000")
    End If
Next

If (sema4Port.PortOpen) Then
    While (sema4Port.InBufferCount > 0)
        ' Dump any received characters to avoid buffer overrun
        sema4Port.InBufferCount = 0
    
        ' Perform event dispatch to keep GUI alive
        DoEvents
    Wend
End If

Exit Sub

comPortFailure:
    sema4PortFailed

End Sub
