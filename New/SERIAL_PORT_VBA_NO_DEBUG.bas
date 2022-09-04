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
Private Const LONG_125000 As Long = 125000

Private Const HEX_00 As Byte = &H0                      ' some hexadecimal constants for minor readability gain.
Private Const HEX_01 As Byte = &H1
Private Const HEX_02 As Byte = &H2
Private Const HEX_03 As Byte = &H3
Private Const HEX_04 As Byte = &H4
Private Const HEX_08 As Byte = &H8
Private Const HEX_0F As Byte = &HF

Private Const HEX_10 As Byte = &H10
Private Const HEX_20 As Byte = &H20
Private Const HEX_30 As Byte = &H30
Private Const HEX_40 As Byte = &H40
Private Const HEX_7F As Byte = &H7F
Private Const HEX_80 As Byte = &H80
Private Const HEX_C0 As Byte = &HC0

Private Const HEX_100 As Long = &H100
Private Const HEX_102 As Long = &H102
Private Const HEX_103 As Long = &H103
Private Const HEX_200 As Long = &H200
Private Const HEX_400 As Long = &H400
Private Const HEX_800 As Long = &H800

Private Const HEX_1000 As Long = &H1000
Private Const HEX_2000 As Long = &H2000
Private Const HEX_3000 As Long = &H3000
Private Const HEX_4000 As Long = &H4000
Private Const HEX_8000 As Long = &H8000
Private Const HEX_C000 As Long = &HC000

Private Const TEXT_DOT As String = "."
Private Const TEXT_DASH As String = "-"
Private Const TEXT_COMMA As String = ","
Private Const TEXT_SPACE As String = " "
Private Const TEXT_EQUALS As String = "="
Private Const TEXT_DOUBLE_SPACE As String = "  "
Private Const TEXT_EQUALS_SPACE As String = "= "
Private Const TEXT_SPACE_EQUALS As String = " ="
Private Const TEXT_COM_PORT As String = "COM Port "

Private Type DEVICE_CONTROL_BLOCK                         ' DCB  - Check latest Microsoft documentation

             LENGTH_DCB As Long
             Baud_Rate  As Long
             BIT_FIELD  As Long
             Reserved   As Integer
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
             READ_TIMEOUT As Boolean
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

             Name As String
             Debug As Boolean
             Handle As LongPtr
             Settings As String
             Port_Errors As Long
             Port_Signals As Long
             Status As COM_PORT_STATUS
             Timers As COM_PORT_TIMERS
             Buffers As COM_PORT_BUFFERS
             Timeouts As COM_PORT_TIMEOUTS
             DCB As DEVICE_CONTROL_BLOCK
End Type

Private Enum PORT_FILE_MODES

             GENERIC_RW = &HC0000000
             GENERIC_READ = &H80000000
             GENERIC_WRITE = &H40000000
             OPEN_EXISTING = 3
             OPEN_EXCLUSIVE = 0
End Enum

Private Enum PORT_FILE_FLAGS

             SYNCHRONOUS_MODE = 0
             ATTRIBUTE_NORMAL = &H80
             NO_BUFFERING = &H20000000
             WRITE_THROUGH = &H80000000
             Overlapped_Mode = &H40000000
End Enum

Private Enum PORT_BAUD_RATE

             CBR_110 = 110
             CBR_300 = 300
             CBR_600 = 600
             CBR_1200 = 1200
             CBR_2400 = 2400
             CBR_4800 = 4800
             CBR_9600 = 9600
             CBR_19200 = 19200                            ' add further baud rates if required.
End Enum

Private Enum PORT_DATA_BITS
             
             BITS_5 = 5
             BITS_6 = 6
             BITS_7 = 7
             BITS_8 = 8
End Enum

Private Enum PORT_FRAMING

             PARITY_NONE = 0
             PARITY_ODD = 1
             PARITY_EVEN = 2
             PARITY_MARK = 3
             PARITY_SPACE = 4
             STOP_BITS_ONE = 0
             STOP_BITS_1P5 = 1
             STOP_BITS_TWO = 2
End Enum

Private Enum PORT_EVENT
                
             RX_CHAR = HEX_01                               ' Normal Character Received Event
             RX_FLAG = HEX_02                               ' Escaped or Interrupt Character - e.g. Control-C Event
             TX_EMPTY = HEX_04                              ' Transmit Buffer Empty
             CTS = HEX_08                                   ' Clear To Send
             DSR = HEX_10                                   ' Data Set (modem or equivalent comms device) Ready
             RLSD = HEX_20                                  ' Receive Line Signal Detect
             RING = HEX_100
             BREAK = HEX_40
             EVENT_ERROR = HEX_00
             LINE_ERROR = HEX_80                            ' Line Error (Parity/Frame/Overrun)
             RX_80_FULL = HEX_400                           ' Receive Buffer 80% full
             EVENT_1 = HEX_800
             EVENT_2 = HEX_1000
             PRINTER_ERROR = HEX_200
End Enum

Private Enum PORT_CONTROL

             DTR_CONTROL_ENABLE = 1
             DTR_CONTROL_DISABLE = 0
             DTR_CONTROL_HANDSHAKE = 2
             RTS_CONTROL_TOGGLE = 3
             RTS_CONTROL_ENABLE = 1
             RTS_CONTROL_DISABLE = 0
             RTS_CONTROL_HANDSHAKE = 2
             CTS_ON = HEX_10
             DSR_ON = HEX_20
             RING_ON = HEX_40
             RLSD_ON = HEX_80
             PURGE_ALL = HEX_0F
             PURGE_ABORT_TX = HEX_01
             PURGE_ABORT_RX = HEX_02
             PURGE_CLEAR_TX = HEX_04
             PURGE_CLEAR_RX = HEX_08
End Enum

Private Enum FLOW_CONTROL

             DTR_ON = 5
             DTR_OFF = 6
             RTS_ON = 3
             RTS_OFF = 4
             XOFF_ON = 1
             XOFF_OFF = 2
             BREAK_ON = 8
             BREAK_OFF = 9
End Enum


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

If Not (Port_Number < COM_PORT_MIN Or Port_Number > COM_PORT_MAX) Then

If COM_PORT_CLOSED(Port_Number) Then
If COM_PORT_CREATE(Port_Number) Then
If COM_PORT_CONFIGURE(Port_Number, Port_Setttings) Then
        
        Temp_Result = PURGE_BUFFERS(Port_Number)

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
Dim CREATE_FILE_FLAGS As Long
Dim Temp_Name As String, Device_Path As String

Const DEVICE_PREFIX As String = "\\.\COM"

Device_Path = DEVICE_PREFIX & CStr(Port_Number)

CREATE_FILE_FLAGS = PORT_FILE_FLAGS.SYNCHRONOUS_MODE

Temp_Name = TEXT_COM_PORT & CStr(Port_Number) & TEXT_COMMA

Temp_Handle = Com_Port_Open(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, CREATE_FILE_FLAGS)

Select Case Temp_Handle

Case HANDLE_INVALID

    Temp_Result = False
    COM_PORT(Port_Number).Name = vbNullString
    COM_PORT(Port_Number).Handle = LONG_0

Case Else

    Temp_Result = True
    COM_PORT(Port_Number).Name = Temp_Name
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

If Query_Port_DCB(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB) Then

    New_Settings = IIf(Len(Port_Settings) > LONG_4, True, False)
     
    If New_Settings Then

        Temp_Build = Build_Port_DCB(Port_Settings, COM_PORT(Port_Number).DCB)
        
        If Temp_Build Then Temp_Result = Apply_Port_DCB(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
                             
    Else

        Temp_Result = True
       
    End If

Else

    Temp_Result = False
   
End If

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

COM_PORT(Port_Number).Timers.Port_Data_Time = LONG_0
COM_PORT(Port_Number).Timers.Last_Data_Time = LONG_0
COM_PORT(Port_Number).Timers.Bytes_Per_Second = Bytes_Per_Second
COM_PORT(Port_Number).Timers.Frame_MicroSeconds = Frame_MicroSeconds
COM_PORT(Port_Number).Timers.Frame_MilliSeconds = Frame_MicroSeconds / LONG_1000
COM_PORT(Port_Number).Buffers.Read_Buffer_Length = Read_Buffer_Length

COM_PORT(Port_Number).Settings = GET_PORT_SETTINGS(Port_Number)

COM_PORT(Port_Number).Timers.Exit_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_EXIT
COM_PORT(Port_Number).Timers.Char_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_CHAR
COM_PORT(Port_Number).Timers.Data_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_DATA
COM_PORT(Port_Number).Timers.Line_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_LINE

If COM_PORT(Port_Number).Timers.Exit_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Exit_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Char_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Char_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Data_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Data_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Line_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Line_Loop_Wait = LONG_1000

Timeslice_Byte_Count = IIf(Bytes_Per_Second < Read_Buffer_Length, True, False)
Effective_Byte_Count = IIf(Timeslice_Byte_Count, Bytes_Per_Second, Read_Buffer_Length)

COM_PORT(Port_Number).Timers.Timeslice_Bytes = Effective_Byte_Count

Select Case Bytes_Per_Second

    Case Is > LONG_1000: COM_PORT(Port_Number).Timers.Read_Wait_Time = READ_EXIT_TIMER_FAST
    Case Is < LONG_100:  COM_PORT(Port_Number).Timers.Read_Wait_Time = READ_EXIT_TIMER_SLOW
    Case Else:           COM_PORT(Port_Number).Timers.Read_Wait_Time = READ_EXIT_TIMER_ELSE

End Select

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

        COM_PORT(Port_Number).Name = vbNullString
        COM_PORT(Port_Number).Handle = LONG_0
        Temp_Result = True
     
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
Dim Loop_Iteration As Long, Wait_Remaining As Long, Loop_Wait_Time As Long
Dim Queue_Length As Long, Sleep_Time As Long

Const Loop_Time As Long = LONG_100                        ' MilliSeconds

Wait_Remaining = IIf(Wait_MilliSeconds < LONG_1, LONG_1, Wait_MilliSeconds)
Loop_Wait_Time = IIf(Wait_MilliSeconds < Loop_Time, Wait_Remaining, Loop_Time)
Loop_Iteration = Int(Wait_Remaining / Loop_Wait_Time) + IIf(Wait_Remaining Mod Loop_Wait_Time > LONG_0, LONG_1, LONG_0)

Do

Clear_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)

If Clear_Result Then

    Queue_Length = COM_PORT(Port_Number).Status.QUEUE_IN
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

SYNCHRONOUS_WAIT_COM_PORT = Data_Waiting

End Function

Public Function READ_COM_PORT(Port_Number As Long, Optional Number_Characters As Long) As String

Dim Temp_Result As Boolean
Dim Read_Timeslice_Bytes As Long
Dim Read_Character_Count As Long
Dim Read_Character_String As String

If Port_Ready(Port_Number) Then
        
    Read_Character_String = vbNullString
    Read_Character_Count = Number_Characters
    Read_Timeslice_Bytes = COM_PORT(Port_Number).Timers.Timeslice_Bytes
    
    If Number_Characters < LONG_1 Or Number_Characters > Read_Timeslice_Bytes Then Read_Character_Count = Read_Timeslice_Bytes
    
    Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, Read_Character_Count)
            
    If Temp_Result And Not COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Then Read_Character_String = COM_PORT(Port_Number).Buffers.Read_Result
                   
End If
    
READ_COM_PORT = Read_Character_String

End Function

Public Function RECEIVE_COM_PORT(Port_Number As Long) As String

Dim Temp_Result As Boolean

If Port_Ready(Port_Number) Then

    COM_PORT(Port_Number).Buffers.Receive_Result = vbNullString

    Do
        Do
            Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, COM_PORT(Port_Number).Timers.Timeslice_Bytes)
            
            If Temp_Result And Not COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Then
            
                COM_PORT(Port_Number).Buffers.Receive_Result = COM_PORT(Port_Number).Buffers.Receive_Result & COM_PORT(Port_Number).Buffers.Read_Result
                
                Select Case COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read
                
                    Case Is < LONG_4:   Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Char_Loop_Wait
                    Case Is < LONG_21:  Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Data_Loop_Wait
                    Case Is = COM_PORT(Port_Number).Timers.Timeslice_Bytes
                    Case Else:          Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Line_Loop_Wait
                
                End Select
                                            
                DoEvents
            
            End If
                        
        Loop Until COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Or Not Temp_Result
        
        If Not COM_PORT(Port_Number).Timers.READ_TIMEOUT Then Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Exit_Loop_Wait
                
     Loop Until COM_PORT(Port_Number).Timers.READ_TIMEOUT Or Not Temp_Result
     
End If
      
RECEIVE_COM_PORT = COM_PORT(Port_Number).Buffers.Receive_Result

End Function

Public Function TRANSMIT_COM_PORT(Port_Number As Long, Transmit_Text As String) As Boolean

Dim Loop_Closing As Boolean, Temp_Result As Boolean
Dim Temp_Pointer As Long, Transmit_Length As Long
Dim Byte_Pointer As Long, Timeslice_Bytes As Long
Dim Byte_Count As Long, Loop_Counter As Long

If Port_Ready(Port_Number) Then

    Transmit_Length = Len(Transmit_Text)
    Timeslice_Bytes = COM_PORT(Port_Number).Timers.Timeslice_Bytes

    For Loop_Counter = LONG_1 To Transmit_Length Step Timeslice_Bytes

        Byte_Pointer = (Loop_Counter + Timeslice_Bytes) - LONG_1
        Loop_Closing = IIf(Transmit_Length - Loop_Counter < Timeslice_Bytes, True, False)
        Temp_Pointer = IIf(Loop_Closing, Transmit_Length, Byte_Pointer)
        Byte_Count = Temp_Pointer - Loop_Counter + LONG_1
        
        COM_PORT(Port_Number).Buffers.Write_Buffer = Mid$(Transmit_Text, Loop_Counter, Timeslice_Bytes)

        Temp_Result = SYNCHRONOUS_WRITE_COM_PORT(Port_Number)

        DoEvents

    Next Loop_Counter

End If

DoEvents

TRANSMIT_COM_PORT = Temp_Result

End Function

Private Function GET_FRAME_TIME(Port_Number As Long) As Single

Dim Frame_Duration As Single
Dim Frame_Length As Long, Length_Stop As Long, Baud_Rate As Long
Dim Length_Start As Long, Length_Data As Long, Length_Parity As Long

Baud_Rate = COM_PORT(Port_Number).DCB.Baud_Rate

Length_Start = LONG_1
Length_Data = COM_PORT(Port_Number).DCB.BYTE_SIZE
Length_Stop = IIf(COM_PORT(Port_Number).DCB.STOP_BITS = LONG_0, LONG_1, LONG_2)
Length_Parity = IIf(COM_PORT(Port_Number).DCB.PARITY = LONG_0, LONG_0, LONG_1)

Frame_Length = Length_Start + Length_Data + Length_Parity + Length_Stop
Frame_Duration = Frame_Length / Baud_Rate * LONG_1E6   ' frame (character) duration in MicroSeconds

GET_FRAME_TIME = Frame_Duration

End Function

Public Function GET_PORT_SETTINGS(Port_Number As Long) As String

Dim Port_Settings As String

Const TEXT_PORT_INVALID As String = "INVALID-PORT"
Const TEXT_NOT_STARTED As String = "PORT-NOT-STARTED"

If Not (Port_Number < COM_PORT_MIN) Or (Port_Number > COM_PORT_MAX) Then

    If COM_PORT(Port_Number).Handle > LONG_0 Then

        Port_Settings = vbNullString
        Port_Settings = Port_Settings & COM_PORT(Port_Number).DCB.Baud_Rate & TEXT_DASH
        Port_Settings = Port_Settings & COM_PORT(Port_Number).DCB.BYTE_SIZE & TEXT_DASH
        Port_Settings = Port_Settings & CONVERT_PARITY(COM_PORT(Port_Number).DCB.PARITY) & TEXT_DASH
        Port_Settings = Port_Settings & CONVERT_STOPBITS(COM_PORT(Port_Number).DCB.STOP_BITS)

    Else

        Port_Settings = TEXT_NOT_STARTED

    End If

Else

    Port_Settings = TEXT_PORT_INVALID

End If

GET_PORT_SETTINGS = Port_Settings

End Function

Private Function SYNCHRONOUS_READ_COM_PORT(Port_Number As Long, Read_Bytes_Requested As Long) As Boolean

Dim Temp_Result As Boolean

Temp_Result = Synchronous_Read(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Read_Buffer, Read_Bytes_Requested, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)

If Temp_Result Then

    If COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read = LONG_0 Then
     
        COM_PORT(Port_Number).Timers.Last_Data_Time = GET_HOST_MICROSECONDS - COM_PORT(Port_Number).Timers.Port_Data_Time
        COM_PORT(Port_Number).Timers.READ_TIMEOUT = IIf(COM_PORT(Port_Number).Timers.Last_Data_Time > COM_PORT(Port_Number).Timers.Read_Wait_Time, True, False)
        COM_PORT(Port_Number).Buffers.Read_Result = vbNullString
        COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = True
    
    Else
        
        COM_PORT(Port_Number).Timers.Port_Data_Time = GET_HOST_MICROSECONDS
        COM_PORT(Port_Number).Timers.Last_Data_Time = LONG_0
        COM_PORT(Port_Number).Timers.READ_TIMEOUT = False
        COM_PORT(Port_Number).Buffers.Read_Result = Left$(COM_PORT(Port_Number).Buffers.Read_Buffer, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)
        COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = False
        
    End If

Else

    Temp_Result = False
    COM_PORT(Port_Number).Timers.READ_TIMEOUT = True
    COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = True
    COM_PORT(Port_Number).Buffers.Read_Result = vbNullString
      
End If

DoEvents

SYNCHRONOUS_READ_COM_PORT = Temp_Result

End Function

Private Function SYNCHRONOUS_WRITE_COM_PORT(Port_Number As Long) As Boolean

Dim Write_Buffer_Length As Long
Dim Write_Complete As Boolean, Temp_Result As Boolean

Write_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Write_Buffer)

Temp_Result = Synchronous_Write(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Write_Buffer, Write_Buffer_Length, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent)

If COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent = Write_Buffer_Length Then Write_Complete = True

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

If Port_Ready(Port_Number) Then
    
    Synchronous_Read COM_PORT(Port_Number).Handle, Get_Character, LONG_1, Read_Byte_Count
            
Else

    Get_Character = vbNullString
        
End If

GET_COM_PORT = Get_Character

End Function

Private Function PURGE_COM_PORT(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Temp_Result = Com_Port_Purge(COM_PORT(Port_Number).Handle, PORT_CONTROL.PURGE_ALL)

DoEvents

PURGE_COM_PORT = Temp_Result

End Function

Private Function PURGE_BUFFERS(Port_Number As Long) As Boolean

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
    .Read_Buffer_Length = Len(.Read_Buffer)
    .Synchronous_Bytes_Read = LONG_0
    .Synchronous_Bytes_Sent = LONG_0

End With

PURGE_BUFFERS = True

End Function

Private Function SET_PORT_TIMERS(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Const READ_TIMEOUT As Long = MAXDWORD
Const WRITE_CONSTANT As Long = LONG_3000

COM_PORT(Port_Number).Timeouts.Read_Interval_Timeout = READ_TIMEOUT            ' Timeouts not used for file reads.
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Constant = LONG_0            '
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT
COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Multiplier = LONG_0

Temp_Result = Set_Com_Timers(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Timeouts)

SET_PORT_TIMERS = Temp_Result

End Function

Public Function CHECK_COM_PORT(Port_Number As Long) As Long
Attribute CHECK_COM_PORT.VB_Description = "Count of characters waiting to be read"

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_Result As Boolean, Temp_Queue As Long, Error_Text As String

Temp_Queue = LONG_NEG_1

If Port_Ready(Port_Number) Then
    
        Temp_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
             
        If Temp_Result Then Temp_Queue = COM_PORT(Port_Number).Status.QUEUE_IN
        
End If

DoEvents

CHECK_COM_PORT = Temp_Queue

End Function

Private Function CLEAR_PORT_ERROR(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Temp_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)

CLEAR_PORT_ERROR = Temp_Result

End Function

Public Function DEVICE_READY(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port DSR signal is asserted.
' DSR = Data Set Ready,from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean, Signal_State As Boolean

If Port_Ready(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    If Temp_Result Then Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.DSR_ON, True, False)

End If

DEVICE_READY = Signal_State

End Function

Public Function DEVICE_CALLING(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port RI signal is asserted.
' Ring Indicator, from attached modem, serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean, Signal_State As Boolean

If Port_Ready(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    If Temp_Result Then Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.RING_ON, True, False)

End If

DEVICE_CALLING = Signal_State

End Function

Public Function CLEAR_TO_SEND(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port CTS signal is asserted.
' CTS = Clear To Send, from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_Result As Boolean
Dim Signal_State As Boolean

If Port_Ready(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    If Temp_Result Then Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.CTS_ON, True, False)

End If

CLEAR_TO_SEND = Signal_State

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
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile
' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter

Const QPF As Long = LONG_1000

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MICROSECONDS = Int(Temp_QPC * QPF)

End Function

Public Function Port_Ready(Port_Number As Long) As Boolean

Port_Ready = IIf(COM_PORT(Port_Number).Handle > LONG_0 And Not (Port_Number < COM_PORT_MIN) Or (Port_Number > COM_PORT_MAX), True, False)

End Function

Public Function COM_PORT_CLOSED(Port_Number As Long) As Boolean

COM_PORT_CLOSED = IIf(COM_PORT(Port_Number).Handle < LONG_1, True, False)

End Function

Public Function CONVERT_PARITY(DCB_PARITY As Byte) As String

Dim Parity_Text As String

Select Case DCB_PARITY

Case PORT_FRAMING.PARITY_ODD:       Parity_Text = "O"
Case PORT_FRAMING.PARITY_NONE:      Parity_Text = "N"
Case PORT_FRAMING.PARITY_EVEN:      Parity_Text = "E"
Case PORT_FRAMING.PARITY_MARK:      Parity_Text = "M"
Case PORT_FRAMING.PARITY_SPACE:     Parity_Text = "S"

Case Else:                          Parity_Text = "?"

End Select

CONVERT_PARITY = Parity_Text

End Function

Public Function CONVERT_STOPBITS(DCB_STOPBITS As Byte) As String

Dim Stop_Text As String

Select Case DCB_STOPBITS

Case PORT_FRAMING.STOP_BITS_ONE:    Stop_Text = "1"
Case PORT_FRAMING.STOP_BITS_TWO:    Stop_Text = "2"
Case PORT_FRAMING.STOP_BITS_1P5:    Stop_Text = "1.5"

Case Else:                          Stop_Text = "?"

End Select

CONVERT_STOPBITS = Stop_Text

End Function

