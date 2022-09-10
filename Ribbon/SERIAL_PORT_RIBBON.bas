Attribute VB_Name = "SERIAL_PORT_RIBBON"
Option Explicit

Sub COM_PORT_CONTROL_1(control As IRibbonControl)       'Callback for COM_PORT_START onAction

Const MESSAGE_BOX_TITLE As String = "Start COM Port 1"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim START_RESULT As Boolean
Dim PORT_SETTINGS As String
Dim MESSAGE_BOX_TEXT As String
Dim MESSAGE_BOX_RESULT As Long

START_RESULT = START_COM_PORT(1)
   
PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf
MESSAGE_BOX_TEXT = PORT_SETTINGS & "Start Result = " & START_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_CONTROL_2(control As IRibbonControl)       'Callback for COM_PORT_STOP onAction

Const MESSAGE_BOX_TITLE As String = "Stop COM Port 1"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim STOP_RESULT As Boolean
Dim MESSAGE_BOX_TEXT As String
Dim MESSAGE_BOX_RESULT As Long

STOP_RESULT = STOP_COM_PORT(1)

MESSAGE_BOX_TEXT = "Stop Result = " & STOP_RESULT
   
MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_1(control As IRibbonControl)          'Callback for COM_PORT_CHECK onAction

Const MESSAGE_BOX_TITLE As String = "Check for data waiting"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim CHARACTERS_WAITING As Long
Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TEXT As String
Dim PORT_SETTINGS As String

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

CHARACTERS_WAITING = CHECK_COM_PORT(1)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Waiting = " & CHARACTERS_WAITING

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_2(control As IRibbonControl)          'Callback for COM_PORT_READ onAction

Const MESSAGE_BOX_TITLE As String = "COM Port Data Read Test"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim CHARACTERS_WAITING As Long
Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TEXT As String
Dim CHARACTERS_READ As String
Dim PORT_SETTINGS As String

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

CHARACTERS_READ = READ_COM_PORT(1, 20)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Read = " & CHARACTERS_READ

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_DATA_3(control As IRibbonControl)          'Callback for COM_PORT_WRITE onAction

Const MESSAGE_BOX_TITLE As String = "COM Port Data Send Test"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim CHARACTERS_WAITING As Long
Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TEXT As String
Dim CHARACTERS_SENT As Long
Dim PORT_SETTINGS As String
Dim SEND_MESSAGE As String

PORT_SETTINGS = "Port Settings = " & GET_PORT_SETTINGS(1) & vbCrLf & vbCrLf

SEND_MESSAGE = Application.Name & " " & Application.Version & " @ " & Time & vbCrLf

CHARACTERS_SENT = Len(SEND_MESSAGE)

SEND_COM_PORT 1, SEND_MESSAGE

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Characters Sent = " & CHARACTERS_SENT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_SIGNAL_1(control As IRibbonControl)        'Callback for COM_PORT_RTS_ON onAction

Const MESSAGE_BOX_TITLE As String = "COM Port 1 - Request To Send ON"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TEXT As String
Dim SIGNAL_RESULT As Boolean
Dim PORT_SETTINGS As String

SIGNAL_RESULT = REQUEST_TO_SEND(1, 1)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Set RTS Result = " & SIGNAL_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

Sub COM_PORT_SIGNAL_2(control As IRibbonControl)        'Callback for COM_PORT_RTS_OFF onAction

Const MESSAGE_BOX_TITLE As String = "COM Port 1 - Request To Send OFF"
Const MESSAGE_BOX_BUTTONS As Long = vbInformation + vbOKOnly

Dim MESSAGE_BOX_RESULT As Long
Dim MESSAGE_BOX_TEXT As String
Dim SIGNAL_RESULT As Boolean
Dim PORT_SETTINGS As String

SIGNAL_RESULT = REQUEST_TO_SEND(1, 0)

MESSAGE_BOX_TEXT = PORT_SETTINGS & "Set RTS Result = " & SIGNAL_RESULT

MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

End Sub

