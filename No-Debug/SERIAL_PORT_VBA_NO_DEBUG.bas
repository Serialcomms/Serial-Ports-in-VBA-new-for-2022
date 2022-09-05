Attribute VB_Name = "SERIAL_PORT_VBA"
'
' https://github.com/Serialcomms/Serial-Ports-in-VBA-new-for-2022
'
  Option Explicit

' Option Private Module
'
' Change Com Port min/max values below to match your installed hardware and intended usage

Private Const COM_PORT_MIN As Integer = 1               ' = COM1
Private Const COM_PORT_MAX As Integer = 2               ' = COM2

'------------------------------------------------------------------------------------------
' Optional - can define port settings for your devices here.
' Use constant to start com port instead of settings string.
'
' Public Const BARCODE As String = "Baud=9600 Data=8 Parity=N Stop=1"
' Public Const GPS_SET As String = "Baud=1200 Data=7 Parity=E Stop=1"
'------------------------------------------------------------------------------------------

Private Const HANDLE_INVALID As LongPtr = -1
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const VBA_TIMEOUT As Long = 5200                ' VBA "Not Responding" time in MilliSeconds (approximate)
Private Const LONG_NEG_1 As Long = -1

Private Const LONG_0  As Long = 0                       ' some predefined constants for minor performance gain.
Private Const LONG_1  As Long = 1
Private Const LONG_2  As Long = 2
Private Const LONG_3  As Long = 3
Private Const LONG_4  As Long = 4
Private Const LONG_5  As Long = 5

Private Const LONG_10 As Long = 10
Private Const LONG_21 As Long = 21
Private Const LONG_50 As Long = 50

Private Const LONG_100 As Long = 100
Private Const LONG_333 As Long = 333
Private Const LONG_1000 As Long = 1000
Private Const LONG_3000 As Long = 3000
Private Const LONG_1E6  As Long = 1000000

Private Const HEX_10 As Byte = &H10                      ' some hexadecimal constants for minor readability gain.
Private Const HEX_20 As Byte = &H20
Private Const HEX_40 As Byte = &H40
Private Const HEX_80 As Byte = &H80

Private Type DEVICE_CONTROL_BLOCK                         ' DCB  - Check latest Microsoft documentation

             LENGTH_DCB As Long
             BAUD_RATE  As Long
             BIT_FIELD  As Long
             RESERVED   As Integer
             LIMIT_XON  As Integer
             LIMIT_XOFF As Integer
             BYTE_SIZE  As Byte
             PARITY     As Byte
             STOP_BITS  As Byte
             CHAR_XON   As Byte
             CHAR_XOFF  As Byte
             CHAR_ERROR As Byte
             CHAR_EOF   As Byte
             CHAR_EVENT As Byte
             RESERVED_1 As Integer
End Type

Private Type COM_PORT_STATUS                              ' COMSTAT Structure - Check latest Microsoft documentation

             BIT_FIELD As Long                            ' 32 bits = waiting for CTS, DRS etc. Top 25 bits not used.
             QUEUE_IN  As Long
             QUEUE_OUT As Long
End Type

Private Type COM_PORT_TIMEOUTS                            ' Check latest Microsoft documentation before changing

             Read_Interval_Timeout          As Long
             Read_Total_Timeout_Multiplier  As Long
             Read_Total_Timeout_Constant    As Long
             Write_Total_Timeout_Multiplier As Long
             Write_Total_Timeout_Constant   As Long
End Type

Private Type COM_PORT_TIMERS
            
             Char_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming single characters)
             Data_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming multiple characters)
             Line_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming lines)
             Exit_Loop_Wait As Long                        ' Arbitrary loop wait time before read exit (allow minimum 1 character time)
             Read_Timeout As Boolean
             Timeslice_Bytes As Long                       ' Approximate bytes per second for timesliced synchronous read/write
             Bytes_Per_Second As Long
             Port_Data_Time As Currency                    ' Currency-scaled time in QPC MicroSeconds of > 0 bytes read
             Last_Data_Time As Currency                    ' Currency-scaled time in QPC MicroSeconds since Port_Data_Time
             Read_Wait_Time As Currency                    ' Currency-scaled time in QPC MicroSeconds of read wait before timeout
             Timing_QPC_Now As Currency                    ' Currency-scaled time in QPC MicroSeconds for timing data start
             Timing_QPC_End As Currency                    ' Currency-scaled time in QPC MicroSeconds for timing data end
             Frame_MilliSeconds As Single                  ' Approximate time in MilliSeconds required to send or receive a character
             Frame_MicroSeconds As Single                  ' Approximate time in MicroSeconds required to send or receive a character
End Type

Private Type COM_PORT_BUFFERS
            
             Read_Result As String
             Read_Buffer As String * 4096                  ' fixed size buffer for synchronous port read (maximum timeslice bytes)
             Write_Result As String
             Write_Buffer As String
             Receive_Result As String
             Receive_Buffer As String
             Receive_Length As Long
             Transmit_Length As Long
             Transmit_Result As String
             Transmit_Buffer As String
             Read_Buffer_Empty As Boolean
             Read_Buffer_Length As Long
             Synchronous_Bytes_Read As Long
             Synchronous_Bytes_Sent As Long
End Type

Private Type COM_PORT_PROFILE                              ' Not Microsoft - check/change locally if required

             Handle As LongPtr
             Errors As Long
             Signals As Long
             Status As COM_PORT_STATUS
             Timers As COM_PORT_TIMERS
             Buffers As COM_PORT_BUFFERS
             Timeouts As COM_PORT_TIMEOUTS
             DCB As DEVICE_CONTROL_BLOCK
End Type

Private COM_PORT(COM_PORT_MIN To COM_PORT_MAX) As COM_PORT_PROFILE

Private Declare PtrSafe Sub Kernel_Sleep_MilliSeconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_MilliSeconds As Long)
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Boolean
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_Frequency As Currency) As Boolean

Private Declare PtrSafe Function Query_Port_DCB Lib "Kernel32.dll" Alias "GetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Apply_Port_DCB Lib "Kernel32.dll" Alias "SetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Build_Port_DCB Lib "Kernel32.dll" Alias "BuildCommDCBA" (ByVal Config_Text As String, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Get_Com_Timers Lib "Kernel32.dll" Alias "GetCommTimeouts" (ByVal Port_Handle As LongPtr, ByRef TIMEOUT As COM_PORT_TIMEOUTS) As Boolean
Private Declare PtrSafe Function Set_Com_Timers Lib "Kernel32.dll" Alias "SetCommTimeouts" (ByVal Port_Handle As LongPtr, ByRef TIMEOUT As COM_PORT_TIMEOUTS) As Boolean
Private Declare PtrSafe Function Set_Com_Signal Lib "Kernel32.dll" Alias "EscapeCommFunction" (ByVal Port_Handle As LongPtr, ByVal Signal_Function As Long) As Boolean
Private Declare PtrSafe Function Get_Port_Modem Lib "Kernel32.dll" Alias "GetCommModemStatus" (ByVal Port_Handle As LongPtr, ByRef Modem_Status As Long) As Boolean
Private Declare PtrSafe Function Com_Port_Purge Lib "Kernel32.dll" Alias "PurgeComm" (ByVal Port_Handle As LongPtr, ByVal Port_Purge_Flags As Long) As Boolean
Private Declare PtrSafe Function Com_Port_Close Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As LongPtr) As Boolean

Private Declare PtrSafe Function Com_Port_Clear Lib "Kernel32.dll" Alias "ClearCommError" _
(ByVal Port_Handle As LongPtr, ByRef Port_Error_Mask As Long, ByRef Port_Status As COM_PORT_STATUS) As Boolean

Private Declare PtrSafe Function Com_Port_Open Lib "Kernel32.dll" Alias "CreateFileA" _
(ByVal Port_Name As String, ByVal PORT_ACCESS As Long, ByVal SHARE_MODE As Long, ByVal SECURITY_ATTRIBUTES_NULL As Any, _
 ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, Optional TEMPLATE_FILE_HANDLE_NULL) As LongPtr

Private Declare PtrSafe Function Synchronous_Read Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean

Private Declare PtrSafe Function Synchronous_Write Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean
'
Public Function START_COM_PORT(Port_Number As Long, Optional Port_Setttings As String) As Boolean

' Port_Settings if supplied should have the same structure as the equivalent command-line Mode arguments for a COM Port:
' [baud=b][parity=p][data=d][stop=s][to={on|off}][xon={on|off}][odsr={on|off}][octs={on|off}][dtr={on|off|hs}][rts={on|off|hs|tg}][idsr={on|off}]
' For example, to configure a baud rate of 1200, no parity, 8 data bits, and 1 stop bit, Port_Settings text is "baud=1200 parity=N data=8 stop=1"

Dim Temp_Result As Boolean

If Port_Valid(Port_Number) Then
If COM_PORT_CLOSED(Port_Number) Then
If COM_PORT_CREATE(Port_Number) Then
If COM_PORT_CONFIGURE(Port_Number, Port_Setttings) Then
        
        Temp_Result = True
        
        PURGE_BUFFERS Port_Number

Else

        STOP_COM_PORT Port_Number                           ' close com port if configure failed

End If
End If
End If
End If

START_COM_PORT = Temp_Result

End Function

Private Function COM_PORT_CREATE(Port_Number As Long) As Boolean

Dim Temp_Handle As LongPtr
Dim Temp_Result As Boolean
Dim Temp_Name As String, Device_Path As String

Const OPEN_EXISTING As Long = LONG_3
Const OPEN_EXCLUSIVE As Long = LONG_0
Const GENERIC_RW As Long = &HC0000000
Const SYNCHRONOUS_MODE As Long = LONG_0
Const DEVICE_PREFIX As String = "\\.\COM"
        
Device_Path = DEVICE_PREFIX & CStr(Port_Number)

Temp_Handle = Com_Port_Open(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, SYNCHRONOUS_MODE)

Select Case Temp_Handle

Case HANDLE_INVALID

    Temp_Result = False
    COM_PORT(Port_Number).Handle = LONG_0

Case Else

    Temp_Result = True
    COM_PORT(Port_Number).Handle = Temp_Handle

End Select

COM_PORT_CREATE = Temp_Result

End Function

Private Function COM_PORT_CONFIGURE(Port_Number As Long, Optional Port_Settings As String) As Boolean

Dim Temp_Result As Boolean
Dim Clean_Settings As String

Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)

If SET_PORT_CONFIG(Port_Number, Clean_Settings) Then
If SET_PORT_TIMERS(Port_Number) Then
If SET_PORT_VALUES(Port_Number) Then Temp_Result = True
      
End If
End If
     
COM_PORT_CONFIGURE = Temp_Result

End Function

Private Function SET_PORT_CONFIG(Port_Number As Long, Optional Port_Settings As String) As Boolean

Dim Temp_Build As Boolean
Dim Temp_Result As Boolean
Dim New_Settings As Boolean

With COM_PORT(Port_Number)

If Query_Port_DCB(.Handle, .DCB) Then

    New_Settings = IIf(Len(Port_Settings) > LONG_4, True, False)
     
    If New_Settings Then

        Temp_Build = Build_Port_DCB(Port_Settings, .DCB)
        
        If Temp_Build Then Temp_Result = Apply_Port_DCB(.Handle, .DCB)
                             
    Else

        Temp_Result = True
       
    End If

Else

    Temp_Result = False
   
End If

End With

SET_PORT_CONFIG = Temp_Result

End Function

Private Function SET_PORT_VALUES(Port_Number As Long) As Boolean

' ------------------------------------------------------------------------
'  optional - can change wait_characters_nnnn to suit local requirements.
'
Const WAIT_CHARACTERS_EXIT As Long = 2                  ' characters
Const WAIT_CHARACTERS_CHAR As Long = 5
Const WAIT_CHARACTERS_DATA As Long = 20
Const WAIT_CHARACTERS_LINE As Long = 100
'
'  optional - can change read exit wait timers to suit local requirements.
'
Const READ_EXIT_TIMER_FAST As Long = 100000             ' MicroSeconds
Const READ_EXIT_TIMER_SLOW As Long = 500000
Const READ_EXIT_TIMER_ELSE As Long = 125000
'
' ------------------------------------------------------------------------

Dim Frame_MicroSeconds As Single, Read_Buffer_Length As Long
Dim Effective_Byte_Count As Long, Bytes_Per_Second As Long
Dim Timeslice_Byte_Count As Boolean, Temp_Result As Boolean

Frame_MicroSeconds = GET_FRAME_TIME(Port_Number)
Bytes_Per_Second = Int(LONG_1 / Frame_MicroSeconds * LONG_1E6)
Read_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Read_Buffer)

If Read_Buffer_Length > LONG_0 Then

Temp_Result = True

With COM_PORT(Port_Number)

.Timers.Port_Data_Time = LONG_0
.Timers.Last_Data_Time = LONG_0
.Timers.Bytes_Per_Second = Bytes_Per_Second
.Timers.Frame_MicroSeconds = Frame_MicroSeconds
.Timers.Frame_MilliSeconds = Frame_MicroSeconds / LONG_1000
.Buffers.Read_Buffer_Length = Read_Buffer_Length

.Timers.Exit_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_EXIT
.Timers.Char_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_CHAR
.Timers.Data_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_DATA
.Timers.Line_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_LINE

If .Timers.Exit_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Exit_Loop_Wait = LONG_1000
If .Timers.Char_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Char_Loop_Wait = LONG_1000
If .Timers.Data_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Data_Loop_Wait = LONG_1000
If .Timers.Line_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Line_Loop_Wait = LONG_1000

Timeslice_Byte_Count = IIf(Bytes_Per_Second < Read_Buffer_Length, True, False)
Effective_Byte_Count = IIf(Timeslice_Byte_Count, Bytes_Per_Second, Read_Buffer_Length)

.Timers.Timeslice_Bytes = Effective_Byte_Count

Select Case Bytes_Per_Second

    Case Is > LONG_1000: .Timers.Read_Wait_Time = READ_EXIT_TIMER_FAST
    Case Is < LONG_100:  .Timers.Read_Wait_Time = READ_EXIT_TIMER_SLOW
    Case Else:           .Timers.Read_Wait_Time = READ_EXIT_TIMER_ELSE

End Select

End With

Else   ' read buffer size not > 0

Temp_Result = False

End If

SET_PORT_VALUES = Temp_Result

End Function

Public Function STOP_COM_PORT(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

If Port_Ready(Port_Number) Then

    PURGE_BUFFERS Port_Number
    
    PURGE_COM_PORT Port_Number

    If Com_Port_Close(COM_PORT(Port_Number).Handle) Then

        Temp_Result = True
        
        COM_PORT(Port_Number).Handle = LONG_0
        
    End If
        
End If

STOP_COM_PORT = Temp_Result

End Function

Public Function WAIT_COM_PORT(Port_Number As Long, Optional Wait_MilliSeconds As Long = LONG_333) As Boolean

Dim Wait_Result As Boolean

If Port_Ready(Port_Number) Then Wait_Result = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)
 
WAIT_COM_PORT = Wait_Result

End Function

Private Function SYNCHRONOUS_WAIT_COM_PORT(Port_Number As Long, Wait_MilliSeconds As Long) As Boolean

Dim Data_Waiting As Boolean, Wait_Expired As Boolean, Clear_Result As Boolean
Dim Loop_Iteration As Long, Loop_Remainder As Long, Loop_Wait_Time As Long
Dim Wait_Remaining As Long, Queue_Length As Long, Sleep_Time As Long

Const Loop_Time As Long = LONG_100                        ' MilliSeconds

Wait_Remaining = IIf(Wait_MilliSeconds < LONG_1, LONG_1, Wait_MilliSeconds)
Loop_Wait_Time = IIf(Wait_MilliSeconds < Loop_Time, Wait_Remaining, Loop_Time)
Loop_Remainder = IIf(Wait_Remaining Mod Loop_Wait_Time > LONG_0, LONG_1, LONG_0)
Loop_Iteration = Int(Wait_Remaining / Loop_Wait_Time) + Loop_Remainder

With COM_PORT(Port_Number)

Do

Clear_Result = Com_Port_Clear(.Handle, .Errors, .Status)

If Clear_Result Then

    Queue_Length = .Status.QUEUE_IN
    Data_Waiting = IIf(Queue_Length > LONG_0, True, False)
    
    If Not Data_Waiting Then
    
        Wait_Expired = IIf(Wait_Remaining < LONG_1, True, False)
        
        If Not Wait_Expired Then
     
            Sleep_Time = IIf(Wait_Remaining < Loop_Wait_Time, Wait_Remaining, Loop_Wait_Time)
            
            Kernel_Sleep_MilliSeconds Sleep_Time
            Loop_Iteration = Loop_Iteration - LONG_1
            Wait_Remaining = Wait_Remaining - Sleep_Time
      
        End If
    
    End If
    
Else

    Wait_Expired = True
    Data_Waiting = False

End If

DoEvents

Loop Until Data_Waiting Or Wait_Expired Or Not Clear_Result

End With

SYNCHRONOUS_WAIT_COM_PORT = Data_Waiting

End Function

Public Function READ_COM_PORT(Port_Number As Long, Optional Number_Characters As Long) As String

Dim Temp_Result As Boolean
Dim Read_Timeslice_Bytes As Long
Dim Read_Character_Count As Long
Dim Read_Character_String As String

If Port_Ready(Port_Number) Then

With COM_PORT(Port_Number)
        
    Read_Character_String = vbNullString
    Read_Character_Count = Number_Characters
    Read_Timeslice_Bytes = .Timers.Timeslice_Bytes
    
    If Number_Characters < LONG_1 Or Number_Characters > Read_Timeslice_Bytes Then Read_Character_Count = Read_Timeslice_Bytes
    
    Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, Read_Character_Count)
            
    If Temp_Result And Not .Buffers.Read_Buffer_Empty Then Read_Character_String = .Buffers.Read_Result
                   
End With

End If

READ_COM_PORT = Read_Character_String

End Function

Public Function RECEIVE_COM_PORT(Port_Number As Long) As String

Dim Temp_Result As Boolean, Full_Read As Long
If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

    .Buffers.Receive_Result = vbNullString
    
    Do
        Do
            Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, .Timers.Timeslice_Bytes)
            
            If Temp_Result And Not .Buffers.Read_Buffer_Empty Then
        
                .Buffers.Receive_Result = .Buffers.Receive_Result & .Buffers.Read_Result
            
                Full_Read = .Timers.Timeslice_Bytes
            
                Select Case .Buffers.Synchronous_Bytes_Read
                
                    Case Is < LONG_4:       Kernel_Sleep_MilliSeconds .Timers.Char_Loop_Wait
                    Case Is < LONG_21:      Kernel_Sleep_MilliSeconds .Timers.Data_Loop_Wait
                    Case Is = Full_Read:  ' Timeslice full, no delay, more data anticipated
                    Case Else:              Kernel_Sleep_MilliSeconds .Timers.Line_Loop_Wait
                
                End Select
                                            
                DoEvents
            
            End If
                        
        Loop Until .Buffers.Read_Buffer_Empty Or Not Temp_Result
        
        If Not .Timers.Read_Timeout Then Kernel_Sleep_MilliSeconds .Timers.Exit_Loop_Wait
                
     Loop Until .Timers.Read_Timeout Or Not Temp_Result
     
     End With
     
End If
      
RECEIVE_COM_PORT = COM_PORT(Port_Number).Buffers.Receive_Result

End Function

Public Function TRANSMIT_COM_PORT(Port_Number As Long, Transmit_Text As String) As Boolean

Dim Loop_Closing As Boolean, Temp_Result As Boolean
Dim Temp_Pointer As Long, Transmit_Length As Long
Dim Byte_Pointer As Long, Timeslice_Bytes As Long
Dim Byte_Count As Long, Loop_Counter As Long

If Port_Ready(Port_Number) Then

With COM_PORT(Port_Number)

    Transmit_Length = Len(Transmit_Text)
    Timeslice_Bytes = .Timers.Timeslice_Bytes

    For Loop_Counter = LONG_1 To Transmit_Length Step Timeslice_Bytes

        Byte_Pointer = (Loop_Counter + Timeslice_Bytes) - LONG_1
        Loop_Closing = IIf(Transmit_Length - Loop_Counter < Timeslice_Bytes, True, False)
        Temp_Pointer = IIf(Loop_Closing, Transmit_Length, Byte_Pointer)
        Byte_Count = Temp_Pointer - Loop_Counter + LONG_1
        
        .Buffers.Write_Buffer = Mid$(Transmit_Text, Loop_Counter, Timeslice_Bytes)

        Temp_Result = SYNCHRONOUS_WRITE_COM_PORT(Port_Number)

        DoEvents

    Next Loop_Counter
    
    End With

End If

DoEvents

TRANSMIT_COM_PORT = Temp_Result

End Function

Private Function GET_FRAME_TIME(Port_Number As Long) As Single

Dim Frame_Duration As Single
Dim Frame_Length As Long, Length_Stop As Long, BAUD_RATE As Long
Dim Length_Start As Long, Length_Data As Long, Length_Parity As Long

With COM_PORT(Port_Number)

BAUD_RATE = .DCB.BAUD_RATE

Length_Start = LONG_1
Length_Data = .DCB.BYTE_SIZE
Length_Stop = IIf(.DCB.STOP_BITS = LONG_0, LONG_1, LONG_2)
Length_Parity = IIf(.DCB.PARITY = LONG_0, LONG_0, LONG_1)

Frame_Length = Length_Start + Length_Data + Length_Parity + Length_Stop
Frame_Duration = Frame_Length / BAUD_RATE * LONG_1E6

End With

GET_FRAME_TIME = Frame_Duration

End Function

Public Function GET_PORT_SETTINGS(Port_Number As Long) As String

Dim Port_Settings As String

Const TEXT_DASH As String = "-"
Const TEXT_PORT_INVALID As String = "INVALID-PORT"
Const TEXT_NOT_STARTED As String = "PORT-NOT-STARTED"

If Port_Valid(Port_Number) Then


With COM_PORT(Port_Number)

    If .Handle > LONG_0 Then

        Port_Settings = vbNullString
        Port_Settings = Port_Settings & .DCB.BAUD_RATE & TEXT_DASH
        Port_Settings = Port_Settings & .DCB.BYTE_SIZE & TEXT_DASH
        Port_Settings = Port_Settings & CONVERT_PARITY(.DCB.PARITY) & TEXT_DASH
        Port_Settings = Port_Settings & CONVERT_STOPBITS(.DCB.STOP_BITS)

    Else

        Port_Settings = TEXT_NOT_STARTED

    End If
    
End With

Else

    Port_Settings = TEXT_PORT_INVALID

End If

GET_PORT_SETTINGS = Port_Settings

End Function

Private Function SYNCHRONOUS_READ_COM_PORT(Port_Number As Long, Read_Bytes_Requested As Long) As Boolean

Dim Temp_Result As Boolean

With COM_PORT(Port_Number)

Temp_Result = Synchronous_Read(.Handle, .Buffers.Read_Buffer, Read_Bytes_Requested, .Buffers.Synchronous_Bytes_Read)

If Temp_Result Then

    If .Buffers.Synchronous_Bytes_Read = LONG_0 Then
     
        .Timers.Last_Data_Time = GET_HOST_MICROSECONDS - .Timers.Port_Data_Time
        .Timers.Read_Timeout = IIf(.Timers.Last_Data_Time > .Timers.Read_Wait_Time, True, False)
        .Buffers.Read_Result = vbNullString
        .Buffers.Read_Buffer_Empty = True
    
    Else
        
        .Timers.Port_Data_Time = GET_HOST_MICROSECONDS
        .Timers.Last_Data_Time = LONG_0
        .Timers.Read_Timeout = False
        .Buffers.Read_Result = Left$(.Buffers.Read_Buffer, .Buffers.Synchronous_Bytes_Read)
        .Buffers.Read_Buffer_Empty = False
        
    End If

Else

     Temp_Result = False
    .Timers.Read_Timeout = True
    .Buffers.Read_Buffer_Empty = True
    .Buffers.Read_Result = vbNullString
      
End If

End With

DoEvents

SYNCHRONOUS_READ_COM_PORT = Temp_Result

End Function

Private Function SYNCHRONOUS_WRITE_COM_PORT(Port_Number As Long) As Boolean

Dim Write_Buffer_Length As Long
Dim Write_Complete As Boolean, Temp_Result As Boolean

With COM_PORT(Port_Number)

Write_Buffer_Length = Len(.Buffers.Write_Buffer)

Temp_Result = Synchronous_Write(.Handle, .Buffers.Write_Buffer, Write_Buffer_Length, .Buffers.Synchronous_Bytes_Sent)

If .Buffers.Synchronous_Bytes_Sent = Write_Buffer_Length Then Write_Complete = True

End With

DoEvents

SYNCHRONOUS_WRITE_COM_PORT = Write_Complete

End Function

Public Function SEND_COM_PORT(Port_Number As Long, Send_Variable As Variant) As Boolean

Dim Send_Result As Boolean

If Port_Ready(Port_Number) Then Send_Result = TRANSMIT_COM_PORT(Port_Number, CStr(Send_Variable))

SEND_COM_PORT = Send_Result

End Function

Public Function PUT_COM_PORT(Port_Number As Long, Put_String As String) As Boolean

Dim Write_Result As Boolean
Dim Write_Byte_Count As Long
    
If Port_Ready(Port_Number) Then Write_Result = Synchronous_Write(COM_PORT(Port_Number).Handle, Left$(Put_String, LONG_1), LONG_1, Write_Byte_Count)

PUT_COM_PORT = Write_Result

End Function

Public Function GET_COM_PORT(Port_Number As Long) As String

Dim Read_Byte_Count As Long
Dim Get_Character As String * LONG_1               ' must be fixed length 1

Get_Character = vbNullString

If Port_Ready(Port_Number) Then Synchronous_Read COM_PORT(Port_Number).Handle, Get_Character, LONG_1, Read_Byte_Count
            
GET_COM_PORT = Get_Character

End Function

Private Function PURGE_COM_PORT(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Const HEX_0F As Byte = &HF
Const PURGE_ALL As Long = HEX_0F

Temp_Result = Com_Port_Purge(COM_PORT(Port_Number).Handle, PURGE_ALL)

DoEvents

PURGE_COM_PORT = Temp_Result

End Function

Private Sub PURGE_BUFFERS(Port_Number As Long)

With COM_PORT(Port_Number).Buffers

    .Read_Result = vbNullString
    .Read_Buffer = vbNullString
    .Write_Result = vbNullString
    .Write_Buffer = vbNullString
    .Receive_Result = vbNullString
    .Receive_Buffer = vbNullString
    .Receive_Length = LONG_0
    .Transmit_Length = LONG_0
    .Transmit_Result = vbNullString
    .Transmit_Buffer = vbNullString
    .Read_Buffer_Empty = True
    .Synchronous_Bytes_Read = LONG_0
    .Synchronous_Bytes_Sent = LONG_0
   
End With

End Sub

Private Function SET_PORT_TIMERS(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Const NO_TIMEOUT As Long = MAXDWORD
Const WRITE_CONSTANT As Long = LONG_3000

With COM_PORT(Port_Number)

.Timeouts.Read_Interval_Timeout = NO_TIMEOUT              ' Timeouts not used for file reads.
.Timeouts.Read_Total_Timeout_Constant = LONG_0            '
.Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

.Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT
.Timeouts.Write_Total_Timeout_Multiplier = LONG_0

Temp_Result = Set_Com_Timers(.Handle, .Timeouts)

End With

SET_PORT_TIMERS = Temp_Result

End Function

Public Function CHECK_COM_PORT(Port_Number As Long) As Long
' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_Result As Boolean, Temp_Queue As Long, Error_Text As String

Temp_Queue = LONG_NEG_1

If Port_Ready(Port_Number) Then

        With COM_PORT(Port_Number)

        If Com_Port_Clear(.Handle, .Errors, .Status) Then Temp_Queue = .Status.QUEUE_IN
             
        End With
        
End If

DoEvents

CHECK_COM_PORT = Temp_Queue

End Function

Public Function CLEAR_TO_SEND(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port CTS signal is asserted.
' CTS = Clear To Send, from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean
Dim Signal_State As Boolean

Const CTS_ON As Long = HEX_10

If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

    Temp_Result = Get_Port_Modem(.Handle, .Signals)
    
    If Temp_Result Then Signal_State = IIf(.Signals And CTS_ON, True, False)
    
    End With

End If

CLEAR_TO_SEND = Signal_State

End Function

Public Function DEVICE_READY(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port DSR signal is asserted.
' DSR = Data Set Ready,from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean
Dim Signal_State As Boolean

Const DSR_ON As Long = HEX_20

If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

    Temp_Result = Get_Port_Modem(.Handle, .Signals)
    
    If Temp_Result Then Signal_State = IIf(.Signals And DSR_ON, True, False)

    End With

End If

DEVICE_READY = Signal_State

End Function

Public Function DEVICE_CALLING(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port RI signal is asserted.
' Ring Indicator, from attached modem, serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean
Dim Signal_State As Boolean

Const RING_ON As Long = HEX_40

If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

    Temp_Result = Get_Port_Modem(.Handle, .Signals)
    
    If Temp_Result Then Signal_State = IIf(.Signals And RING_ON, True, False)

    End With

End If

DEVICE_CALLING = Signal_State

End Function

Public Function SIGNAL_COM_PORT(Port_Number As Long, Signal_Function As Long) As Boolean

Dim Signal_Valid As Boolean
Dim Signal_Result As Boolean

Signal_Valid = IIf(Signal_Function < LONG_10 And Signal_Function > LONG_0, True, False)

If Port_Ready(Port_Number) And Signal_Valid Then Signal_Result = Set_Com_Signal(COM_PORT(Port_Number).Handle, Signal_Function)
    
SIGNAL_COM_PORT = Signal_Result

End Function

Public Function REQUEST_TO_SEND_COM_PORT(Port_Number As Long, RTS_State As Boolean) As Boolean

Dim RTS_Signal As Long
Dim RTS_Result As Boolean

Const SIGNAL_RTS_1 As Long = LONG_3
Const SIGNAL_RTS_0 As Long = LONG_4

RTS_Signal = IIf(RTS_State, SIGNAL_RTS_1, SIGNAL_RTS_0)

If Port_Ready(Port_Number) Then
                
    RTS_Result = Set_Com_Signal(COM_PORT(Port_Number).Handle, RTS_Signal)
                        
    If RTS_Result Then Kernel_Sleep_MilliSeconds LONG_50                 ' optional - allow local and remote hardware devices to settle.

End If

REQUEST_TO_SEND_COM_PORT = RTS_Result

End Function

Private Function CLEAN_PORT_SETTINGS(Port_Settings As String) As String

Dim New_Settings As String

Const TEXT_COMMA As String = ","
Const TEXT_SPACE As String = " "
Const TEXT_EQUALS As String = "="
Const TEXT_DOUBLE_SPACE As String = "  "
Const TEXT_EQUALS_SPACE As String = "= "
Const TEXT_SPACE_EQUALS As String = " ="

New_Settings = Trim(Port_Settings)
New_Settings = UCase(New_Settings)

New_Settings = Replace(New_Settings, TEXT_COMMA, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_SPACE_EQUALS, TEXT_EQUALS, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_EQUALS_SPACE, TEXT_EQUALS, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)

CLEAN_PORT_SETTINGS = New_Settings

End Function

Public Function GET_HOST_MICROSECONDS() As Currency

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Const QPF As Long = LONG_1000

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MICROSECONDS = Int(Temp_QPC * QPF)

End Function

Public Function Port_Ready(Port_Number As Long) As Boolean

Port_Ready = IIf(COM_PORT(Port_Number).Handle > LONG_0 And Port_Valid(Port_Number), True, False)

End Function

Public Function Port_Valid(Port_Number As Long) As Boolean

Port_Valid = IIf((Port_Number < COM_PORT_MIN) Or (Port_Number > COM_PORT_MAX), False, True)

End Function

Public Function COM_PORT_CLOSED(Port_Number As Long) As Boolean

COM_PORT_CLOSED = IIf(COM_PORT(Port_Number).Handle < LONG_1, True, False)

End Function

Public Function CONVERT_PARITY(DCB_PARITY As Byte) As String

Dim Parity_Text As String

Select Case DCB_PARITY

Case LONG_0:    Parity_Text = "N"
Case LONG_1:    Parity_Text = "O"
Case LONG_2:    Parity_Text = "E"
Case LONG_3:    Parity_Text = "M"
Case LONG_4:    Parity_Text = "S"

Case Else:                          Parity_Text = "?"

End Select

CONVERT_PARITY = Parity_Text

End Function

Public Function CONVERT_STOPBITS(DCB_STOPBITS As Byte) As String

Dim Stop_Text As String

Select Case DCB_STOPBITS

Case LONG_0:    Stop_Text = "1"
Case LONG_1:    Stop_Text = "1.5"
Case LONG_2:    Stop_Text = "2"

Case Else:                          Stop_Text = "?"

End Select

CONVERT_STOPBITS = Stop_Text

End Function

Public Function DEBUG_COM_PORT(Optional Port_Number As Long, Optional Debug_State As Variant) As Boolean

DEBUG_COM_PORT = False

End Function

