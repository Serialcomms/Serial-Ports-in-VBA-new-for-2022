Attribute VB_Name = "SERIAL_PORT_RIBBON"
Option Explicit

'----------------------------------------
' Change Com port number here if required
'
  Const Port_Number As Long = 1
'----------------------------------------

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

Const Message_Box_Buttons As Long = vbInformation + vbOKOnly
'
Sub COM_PORT_CONTROL_1(Optional control As Variant)   'Callback for COM_PORT_START onAction

Message_Box_Title = "Start COM Port " & Port_Number

Start_Result = START_COM_PORT(Port_Number)
   
Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Message_Box_Text = Port_Settings & "Start Result = " & Start_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_CONTROL_2(Optional control As Variant)   'Callback for COM_PORT_STOP onAction

Message_Box_Title = "Stop COM Port " & Port_Number

Stop_Result = STOP_COM_PORT(Port_Number)

Message_Box_Text = "Stop Result = " & Stop_Result
   
Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_1(Optional control As Variant)  'Callback for COM_PORT_CHECK onAction

Message_Box_Title = "Check for data waiting"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Characters_Waiting = CHECK_COM_PORT(Port_Number)

Message_Box_Text = Port_Settings & "Characters Waiting = " & Characters_Waiting

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_2(Optional control As Variant)  'Callback for COM_PORT_READ onAction

Message_Box_Title = "COM Port Data Read Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Characters_Read = READ_COM_PORT(Port_Number, 20)

Message_Box_Text = Port_Settings & "Characters Read = " & Characters_Read

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_3(Optional control As Variant)  'Callback for COM_PORT_WRITE onAction

Message_Box_Title = "COM Port Data Send Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Send_Message = Application.Name & " " & Application.Version & " @ " & Time & vbCrLf

Characters_Sent = Len(Send_Message)

SEND_COM_PORT Port_Number, Send_Message

Message_Box_Text = Port_Settings & "Characters Sent = " & Characters_Sent

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_1(Optional control As Variant)    'Callback for COM_PORT_RTS_ON onAction

Message_Box_Title = "COM Port " & Port_Number & " - Request To Send ON"

Signal_Result = REQUEST_TO_SEND(Port_Number, 1)

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_2(Optional control As Variant)    'Callback for COM_PORT_RTS_OFF onAction

Message_Box_Title = "COM Port " & Port_Number & " - Request To Send OFF"

Signal_Result = REQUEST_TO_SEND(Port_Number, 0)

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

