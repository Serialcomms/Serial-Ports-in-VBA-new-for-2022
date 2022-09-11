Attribute VB_Name = "SERIAL_PORT_RIBBON"
Option Explicit

Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim MESSAGE_BOX_TEXT As String
Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TITLE As String
Dim PORT_SETTINGS As String
'

Sub COM_PORT_CONTROL_1(control As IRibbonControl)       'Callback for COM_PORT_START onAction

MESSAGE_BOX_TITLE = "Start COM Port 1"

Dim START_RESULT As Boolean

START_RESULT = START_COM_PORT(1)
   
PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf
MESSAGE_BOX_TEXT = PORT_SETTINGS & "Start Result = " & START_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_CONTROL_2(control As IRibbonControl)       'Callback for COM_PORT_STOP onAction

MESSAGE_BOX_TITLE = "Stop COM Port 1"

Dim STOP_RESULT As Boolean

STOP_RESULT = STOP_COM_PORT(1)

MESSAGE_BOX_TEXT = "Stop Result = " & STOP_RESULT
   
MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_1(control As IRibbonControl)          'Callback for COM_PORT_CHECK onAction

MESSAGE_BOX_TITLE = "Check for data waiting"

Dim CHARACTERS_WAITING As Long

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

CHARACTERS_WAITING = CHECK_COM_PORT(1)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Waiting = " & CHARACTERS_WAITING

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_2(control As IRibbonControl)          'Callback for COM_PORT_READ onAction

MESSAGE_BOX_TITLE = "COM Port Data Read Test"

Dim CHARACTERS_WAITING As Long
Dim CHARACTERS_READ As String

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

CHARACTERS_READ = READ_COM_PORT(1, 20)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Read = " & CHARACTERS_READ

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_3(control As IRibbonControl)          'Callback for COM_PORT_WRITE onAction

MESSAGE_BOX_TITLE = "COM Port Data Send Test"

Dim CHARACTERS_SENT As Long
Dim SEND_MESSAGE As String

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

SEND_MESSAGE = Application.Name & " " & Application.Version & " @ " & Time & vbCrLf

CHARACTERS_SENT = Len(SEND_MESSAGE)

SEND_COM_PORT 1, SEND_MESSAGE

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Sent = " & CHARACTERS_SENT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_SIGNAL_1(control As IRibbonControl)        'Callback for COM_PORT_RTS_ON onAction

MESSAGE_BOX_TITLE = "COM Port 1 - Request To Send ON"

Dim SIGNAL_RESULT As Boolean

SIGNAL_RESULT = REQUEST_TO_SEND(1, 1)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Set RTS Result = " & SIGNAL_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_SIGNAL_2(control As IRibbonControl)        'Callback for COM_PORT_RTS_OFF onAction

MESSAGE_BOX_TITLE = "COM Port 1 - Request To Send OFF"

Dim SIGNAL_RESULT As Boolean

SIGNAL_RESULT = REQUEST_TO_SEND(1, 0)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Set RTS Result = " & SIGNAL_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

