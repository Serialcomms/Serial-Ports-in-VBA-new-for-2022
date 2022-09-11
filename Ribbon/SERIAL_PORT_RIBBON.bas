Attribute VB_Name = "SERIAL_PORT_RIBBON"
Option Explicit

Dim Stop_Result As Boolean
Dim Start_Result As Boolean
Dim Signal_Result As Boolean

Dim Characters_Sent As Long
Dim Characters_Waiting As Long

Dim Send_Message As String
Dim Port_Settings As String
Dim Characters_Read As String

Dim Message_Box_Text As String
Dim Message_Box_Title As String
Dim Message_Box_Result As Long

Const Number As Long = 1

Const Message_Box_Buttons As Long = vbInformation + vbOKOnly
'
Sub COM_PORT_CONTROL_1(control As IRibbonControl)   'Callback for COM_PORT_START onAction

Message_Box_Title = "Start COM Port " & Number

Start_Result = START_COM_PORT(Number)
   
Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Number) & vbCrLf & vbCrLf

Message_Box_Text = Port_Settings & "Start Result = " & Start_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_CONTROL_2(control As IRibbonControl)   'Callback for COM_PORT_STOP onAction

Message_Box_Title = "Stop COM Port " & Number

Stop_Result = STOP_COM_PORT(Number)

Message_Box_Text = "Stop Result = " & Stop_Result
   
Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_1(control As IRibbonControl)  'Callback for COM_PORT_CHECK onAction

Message_Box_Title = "Check for data waiting"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Number) & vbCrLf & vbCrLf

Characters_Waiting = CHECK_COM_PORT(Number)

Message_Box_Text = Port_Settings & "Characters Waiting = " & Characters_Waiting

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_2(control As IRibbonControl)  'Callback for COM_PORT_READ onAction

Message_Box_Title = "COM Port Data Read Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Number) & vbCrLf & vbCrLf

Characters_Read = READ_COM_PORT(Number, 20)

Message_Box_Text = Port_Settings & "Characters Read = " & Characters_Read

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_3(control As IRibbonControl)  'Callback for COM_PORT_WRITE onAction

Message_Box_Title = "COM Port Data Send Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Number) & vbCrLf & vbCrLf

Send_Message = Application.Name & " " & Application.Version & " @ " & Time & vbCrLf

Characters_Sent = Len(Send_Message)

SEND_COM_PORT Number, Send_Message

Message_Box_Text = Port_Settings & "Characters Sent = " & Characters_Sent

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_1(control As IRibbonControl)    'Callback for COM_PORT_RTS_ON onAction

Message_Box_Title = "COM Port " & Number & " - Request To Send ON"

Signal_Result = REQUEST_TO_SEND(Number, 1)

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_2(control As IRibbonControl)    'Callback for COM_PORT_RTS_OFF onAction

Message_Box_Title = "COM Port " & Number & " - Request To Send OFF"

Signal_Result = REQUEST_TO_SEND(Number, 0)

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

