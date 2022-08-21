Attribute VB_Name = "SERIAL_PORT_VBA"
Option Explicit

' Option Private Module
' Change Com Port min/max values below to match your installed hardware and intended usage

Private Const COM_PORT_MIN As Integer = 1               ' = COM1
Private Const COM_PORT_MAX As Integer = 2               ' = COM2

Private Const MAXDWORD As Long = &HFFFFFFFF

Private Const VBA_TIMEOUT As Long = 5200                ' VBA "Not Responding" time in Milliseconds (approximate)

Private Const LONG_NEG_1 As Long = -1

Private Const LONG_0 As Long = 0                        ' some predefined constants for minor performance gain.
Private Const LONG_1 As Long = 1
Private Const LONG_2 As Long = 2
Private Const LONG_3 As Long = 3
Private Const LONG_4 As Long = 4
Private Const LONG_5 As Long = 5
Private Const LONG_6 As Long = 6
Private Const LONG_7 As Long = 7
Private Const LONG_8 As Long = 8
Private Const LONG_9 As Long = 9
Private Const LONG_10 As Long = 10
Private Const LONG_14 As Long = 14
Private Const LONG_20 As Long = 20
Private Const LONG_21 As Long = 21
Private Const LONG_30 As Long = 30
Private Const LONG_36 As Long = 36
Private Const LONG_40 As Long = 40
Private Const LONG_50 As Long = 50
Private Const LONG_60 As Long = 60

Private Const LONG_100 As Long = 100
Private Const LONG_333 As Long = 333
Private Const LONG_1000 As Long = 1000
Private Const LONG_3000 As Long = 3000
Private Const LONG_1E5 As Long = 100000
Private Const LONG_1E6 As Long = 1000000
Private Const LONG_50000 As Long = 50000
Private Const LONG_100000 As Long = 100000
Private Const LONG_120000 As Long = 120000
Private Const LONG_125000 As Long = 125000
Private Const LONG_250000 As Long = 250000
Private Const LONG_333333 As Long = 333333
Private Const LONG_500000 As Long = 500000

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

Private Const TEXT_CR As String = vbCr
Private Const TEXT_LF As String = vbLf
Private Const TEXT_MS As String = " mS"                 ' some text string constants for minor gains.
Private Const TEXT_ON As String = "On"
Private Const TEXT_OFF As String = "Off"
Private Const TEXT_TRUE As String = "True"
Private Const TEXT_FALSE As String = "False"

Private Const TEXT_DOT As String = "."
Private Const TEXT_DASH As String = "-"
Private Const TEXT_COMMA As String = ","
Private Const TEXT_SPACE As String = " "
Private Const TEXT_EQUALS As String = "="
Private Const TEXT_DOUBLE_SPACE As String = "  "
Private Const TEXT_EQUALS_SPACE As String = "= "
Private Const TEXT_SPACE_EQUALS As String = " ="

Private Const TEXT_CONFIG As String = "CONFIG"
Private Const TEXT_RESULT As String = "RESULT"
Private Const TEXT_FAILED As String = "FAILED"
Private Const TEXT_SINGLE As String = "SINGLE"
Private Const TEXT_TIMING As String = "TIMING"

Private Const TEXT_SUCCESS As String = "SUCCESS"
Private Const TEXT_FAILURE As String = "FAILURE"
Private Const TEXT_WAITING As String = "WAITING"
Private Const TEXT_READING As String = "READING"
Private Const TEXT_WRITING As String = "WRITING"
Private Const TEXT_LOOPING As String = "LOOPING"
Private Const TEXT_TIMEOUT As String = "TIMEOUT"

Private Const TEXT_DURATION As String = "DURATION"
Private Const TEXT_RECEIVED As String = "RECEIVED"
Private Const TEXT_STARTING As String = "STARTING"
Private Const TEXT_FINISHED As String = "FINISHED"
Private Const TEXT_SETTINGS As String = "SETTINGS"

Private Const TEXT_COM_PORT As String = "COM Port "

Private Const COM_PORT_RANGE As String = COM_PORT_MIN & " to " & COM_PORT_MAX

Private Type SYSTEMTIME

             Year As Integer
             Month As Integer
             WeekDay As Integer
             Day As Integer
             Hour As Integer
             Minute As Integer
             Second As Integer
             Milliseconds As Integer                      ' used for debug timestamp
End Type

Private Type DEVICE_CONTROL_BLOCK                         ' DCB  - Check latest Microsoft documentation

             LENGTH_DCB As Long
             Baud_Rate As Long
             BIT_FIELD As Long
             Reserved As Integer
             LIMIT_XON As Integer
             LIMIT_XOFF As Integer
             BYTE_SIZE As Byte
             PARITY As Byte
             STOP_BITS As Byte
             CHAR_XON As Byte
             CHAR_XOFF As Byte
             CHAR_ERROR As Byte
             CHAR_EOF As Byte
             CHAR_EVENT As Byte
             RESERVED_1 As Integer
End Type

Private Type COM_PORT_STATUS                              ' COMSTAT Structure - Check latest Microsoft documentation

             BIT_FIELD As Long                            ' 32 bits = waiting for CTS, DRS etc. Top 25 bits not used.
             QUEUE_IN As Long
             QUEUE_OUT As Long
End Type

Private Type COM_PORT_TIMEOUTS                            ' Check latest Microsoft documentation before changing

             Read_Interval_Timeout As Long
             Read_Total_Timeout_Multiplier As Long
             Read_Total_Timeout_Constant As Long
             Write_Total_Timeout_Multiplier As Long
             Write_Total_Timeout_Constant As Long
End Type

Private Type COM_PORT_CONFIG                              ' Check latest Microsoft documentation before changing

             Size As Long
             Version As Integer
             Reserved As Integer
             DCB As DEVICE_CONTROL_BLOCK
             Provider_SubType As Long
             Provider_Offset As Long
             Provider_Size As Long
             Provider_Data As Byte
End Type

Private Type COM_PORT_TIMERS
            
             Char_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming single characters)
             Data_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming multiple characters)
             Line_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming lines)
             Exit_Loop_Wait As Long                        ' Arbitrary loop wait time before read exit (allow minimum 1 char)
             Read_Timeout As Boolean
             Timeslice_Bytes As Long                       ' Approximate bytes per second for timesliced synchronous read/write
             Bytes_Per_Second As Long
             Port_Data_Time As Currency                    ' Time in QPC Microseconds of > 0 bytes read
             Last_Data_Time As Currency                    ' Time in QPC Microseconds since Port_Data_Time
             Read_Wait_Time As Currency                    ' Time in QPC Microseconds of read wait before timeout
             Timing_QPC_Now As Currency                    ' Win32 Query Performance Counter for microsecond timing data
             Timing_QPC_End As Currency                    ' Win32 Query Performance Counter for microsecond timing data
             Frame_MilliSeconds As Single                  ' Approximate time in Milliseconds required to send or receive a character
             Frame_MicroSeconds As Single                  ' Approximate time in Microseconds required to send or receive a character
End Type

Private Type COM_PORT_BUFFERS
            
             Read_Result As String
             Read_Buffer As String * 4096                  ' fixed size buffer for synchronous port read (maximum timeslice bytes)
             Write_Result As String
             Write_Buffer As String
             Receive_Length As Long
             Receive_Result As String
             Receive_Buffer As String
             Transmit_Length As Long
             Transmit_Result As String
             Transmit_Buffer As String
             Read_Buffer_Length As Long
             Read_Buffer_Empty As Boolean
             Synchronous_Bytes_Read As Long
             Synchronous_Bytes_Sent As Long
End Type

Private Type COM_PORT_PROFILE                              ' Not Microsoft - check/change locally if required

             Name As String
             Handle As Long
             Debug As Boolean
             DLL_Error As Long
             Settings As String
             Port_Errors As Long
             Port_Signals As Long
             Timers As COM_PORT_TIMERS
             Config As COM_PORT_CONFIG
             Status As COM_PORT_STATUS
             Buffers As COM_PORT_BUFFERS
             DCB As DEVICE_CONTROL_BLOCK
             Timeouts As COM_PORT_TIMEOUTS
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

Private Enum Port_Errors

             OVERFLOW = HEX_01       ' Input buffer overflow, buffer full or data after EOF
             OVERRUN = HEX_02        ' Character-buffer overrun. The next character is lost
             PARITY = HEX_04         ' Port hardware detected a parity error
             FRAME = HEX_08          ' Port hardware detected a framing error
             BREAK = HEX_10          ' Port hardware detected a break signal
End Enum

Private Enum SYSTEM_ERRORS
              
             SUCCESS = 0
             INVALID_FUNCTION = 1
             FILE_NOT_FOUND = 2
             PATH_NOT_FOUND = 3
             TOO_MANY_OPEN_FILES = 4
             ACCESS_DENIED = 5
             INVALID_HANDLE = 6
             INVALID_DATA = 13
             DEVICE_NOT_READY = 15
             INVALID_PARAMETER = 87
             INSUFFICIENT_BUFFER = 122
             OPERATION_ABORTED = 995
             IO_INCOMPLETE = 996
             IO_PENDING = 997
             NO_ACCESS = 998
End Enum

Private COM_PORT(COM_PORT_MIN To COM_PORT_MAX) As COM_PORT_PROFILE

Private Declare PtrSafe Sub Kernel_Sleep_Milliseconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_Milliseconds As Long)
Private Declare PtrSafe Function Get_System_Time Lib "Kernel32.dll" Alias "GetSystemTime" (ByRef System_Time As SYSTEMTIME) As Boolean
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_Frequency As Currency) As Boolean
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Boolean

' https://docs.microsoft.com/en-us/windows/win32/devio/communications-functions

Private Declare PtrSafe Function Get_Com_State Lib "Kernel32.dll" Alias "GetCommState" (ByVal Port_Handle As Long, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Set_Com_State Lib "Kernel32.dll" Alias "SetCommState" (ByVal Port_Handle As Long, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Set_Com_Queue Lib "Kernel32.dll" Alias "SetupComm" (ByVal Port_Handle As Long, ByVal QUEUE_IN As Long, ByVal QUEUE_OUT As Long) As Boolean
Private Declare PtrSafe Function Get_Com_Config Lib "Kernel32.dll" Alias "GetCommConfig" (ByVal Port_Handle As Long, ByRef Port_CC As COM_PORT_CONFIG, ByVal CC_LENGTH As Long) As Boolean
Private Declare PtrSafe Function Set_Com_Config Lib "Kernel32.dll" Alias "SetCommConfig" (ByVal Port_Handle As Long, ByRef Port_CC As COM_PORT_CONFIG, ByVal CC_LENGTH As Long) As Boolean
Private Declare PtrSafe Function Get_Com_Timers Lib "Kernel32.dll" Alias "GetCommTimeouts" (ByVal Port_Handle As Long, ByRef COM_Timeouts As COM_PORT_TIMEOUTS) As Boolean
Private Declare PtrSafe Function Set_Com_Timers Lib "Kernel32.dll" Alias "SetCommTimeouts" (ByVal Port_Handle As Long, ByRef COM_Timeouts As COM_PORT_TIMEOUTS) As Boolean
Private Declare PtrSafe Function Com_Port_Purge Lib "Kernel32.dll" Alias "PurgeComm" (ByVal Port_Handle As Long, ByVal Port_Purge_Flags As Long) As Boolean
Private Declare PtrSafe Function Build_Port_DCB Lib "Kernel32.dll" Alias "BuildCommDCBA" (ByVal Config_Text As String, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Clear_Com_Break Lib "Kernel32.dll" Alias "ClearCommBreak" (ByVal Port_Handle As Long) As Boolean
Private Declare PtrSafe Function Clear_Com_Error Lib "Kernel32.dll" Alias "ClearCommError" (ByVal Port_Handle As Long, ByRef Error_Mask As Long, ByRef Port_Comstat As COM_PORT_STATUS) As Boolean
Private Declare PtrSafe Function Set_Port_Control Lib "Kernel32.dll" Alias "EscapeCommFunction" (ByVal Port_Handle As Long, ByVal Port_Function As Long) As Boolean
Private Declare PtrSafe Function Get_Modem_Status Lib "Kernel32.dll" Alias "GetCommModemStatus" (ByVal Port_Handle As Long, ByRef Modem_Status As Long) As Boolean
Private Declare PtrSafe Function Com_Port_Release Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As Long) As Boolean

Private Declare PtrSafe Function Com_Port_Seize Lib "Kernel32.dll" Alias "CreateFileA" _
  (ByVal Port_Name As String, ByVal PORT_ACCESS As Long, _
   ByVal SHARE_MODE As Long, ByRef PORT_SECURITY_ATTRIBUTES As Any, _
   ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, _
   ByVal TEMPLATE_FILE_Handle As Any) As Long

Private Declare PtrSafe Function Synchronous_Read Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As Long, ByVal BUFFER_DATA As String, ByVal Bytes_Requested As Long, ByRef BYTES_RECEIVED As Long, ByRef OVERLAPPED_NULL As Any) As Boolean

Private Declare PtrSafe Function Synchronous_Write Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As Long, ByVal BUFFER_DATA As String, ByVal Bytes_Requested As Long, ByRef BYTES_WRITTEN As Long, ByRef OVERLAPPED_NULL As Any) As Boolean
'
Public Function START_COM_PORT(Port_Number As Long, Optional Port_Setttings As String) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "START_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Port_Result As Boolean
Dim Temp_String(LONG_1 To LONG_5) As String, Result_Text As String, Detail_Text As String

Port_Result = False

If Port_Valid Then

Temp_String(LONG_1) = " Attempting to Start and Configure COM Port " & Port_Number
Temp_String(LONG_2) = " Started and Configured COM Port " & Port_Number & " with Handle "
Temp_String(LONG_3) = " Failed to Configure COM Port " & Port_Number
Temp_String(LONG_4) = " Failed to Start, Create and Configure COM Port " & Port_Number
Temp_String(LONG_5) = " Failed to Start COM Port, Existing Port Handle = "

COM_PORT(Port_Number).Name = TEXT_COM_PORT & CStr(Port_Number) & TEXT_COMMA

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTING, COM_PORT(Port_Number).Name & Temp_String(LONG_1))

If COM_PORT(Port_Number).Handle < LONG_1 Then
If COM_PORT_CREATE(Port_Number) Then
If COM_PORT_CONFIGURE(Port_Number, Port_Setttings) Then

Port_Result = True
Result_Text = TEXT_SUCCESS
Detail_Text = Temp_String(LONG_2) & COM_PORT(Port_Number).Handle

Else
Result_Text = TEXT_FAILURE
Detail_Text = Temp_String(LONG_3)
Call STOP_COM_PORT(Port_Number) ' close com port if configure failed
End If
    
Else
Result_Text = TEXT_FAILURE
Detail_Text = Temp_String(LONG_4)
End If

Else
Result_Text = TEXT_FAILURE
Detail_Text = Temp_String(LONG_5) & COM_PORT(Port_Number).Handle
End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

START_COM_PORT = Port_Result

End Function

Private Function COM_PORT_CREATE(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "COM_PORT_CREATE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Temp_String(LONG_1 To LONG_6) As String
Dim Temp_Handle As Long, CREATE_FILE_FLAGS As Long
Dim Device_Path As String, Error_Text As String, Result_Text As String, Detail_Text As String

Const DEVICE_PREFIX As String = "\\.\COM"

Device_Path = DEVICE_PREFIX & CStr(Port_Number)

CREATE_FILE_FLAGS = PORT_FILE_FLAGS.SYNCHRONOUS_MODE

Temp_String(LONG_1) = "CREATING"
Temp_String(LONG_2) = "PORT_MODE"
Temp_String(LONG_3) = " Attempting to Open Port with Device Path " & Device_Path
Temp_String(LONG_4) = " Open for Exclusive Access, Handle = "
Temp_String(LONG_5) = " Failed to Open COM Port, Last Error "
Temp_String(LONG_6) = " Creating Synchronous (non-overlapped) mode Port "

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Temp_String(LONG_1), COM_PORT(Port_Number).Name & Temp_String(LONG_3))
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Temp_String(LONG_2), COM_PORT(Port_Number).Name & Temp_String(LONG_6))

Temp_Handle = Com_Port_Seize(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, ByVal vbNullString, OPEN_EXISTING, CREATE_FILE_FLAGS, ByVal vbNullString)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Select Case COM_PORT(Port_Number).DLL_Error

Case SYSTEM_ERRORS.SUCCESS

    Temp_Bool = True
    Result_Text = TEXT_SUCCESS
    Detail_Text = Temp_String(LONG_4) & Temp_Handle
    COM_PORT(Port_Number).Handle = Temp_Handle
    
Case Else

    Temp_Bool = False
    Result_Text = TEXT_FAILURE
    Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
    Detail_Text = Temp_String(LONG_5) & Error_Text
    COM_PORT(Port_Number).Handle = LONG_0
    
End Select

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

COM_PORT_CREATE = Temp_Bool

End Function

Private Function COM_PORT_CONFIGURE(Port_Number As Long, Optional Port_Settings As String) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "COM_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Length As Long
Dim Temp_Bool As Boolean
Dim Temp_String(LONG_1 To LONG_6) As String
Dim Temp_Settings As String, Temp_Text As String, Result_Text As String, Detail_Text As String

Temp_Bool = False

Temp_Settings = CLEAN_PORT_SETTINGS(Port_Settings)

Temp_Length = Len(Temp_Settings)

Temp_Text = "Settings " & Temp_Settings

Temp_String(LONG_1) = " Attempting to Configure Port With "
Temp_String(LONG_2) = IIf(Temp_Length > LONG_0, Temp_Settings, "Default Settings")
Temp_String(LONG_3) = " Configured COM Port with Settings = "
Temp_String(LONG_4) = " Failed to Set Port Values "
Temp_String(LONG_5) = " Failed to Set Port Timers, Last Error "
Temp_String(LONG_6) = " Failed to Set Port Config, Last Error "

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "STARTUP", COM_PORT(Port_Number).Name & Temp_String(LONG_1) & Temp_String(LONG_2))

If SET_PORT_CONFIG(Port_Number, Temp_Settings) Then
If SET_PORT_TIMERS(Port_Number) Then
If SET_PORT_VALUES(Port_Number) Then
      Temp_Bool = True
      Result_Text = TEXT_SUCCESS: Detail_Text = Temp_String(LONG_3) & GET_PORT_SETTINGS(Port_Number)
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_4): End If
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_5): End If
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_6): End If
     
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

COM_PORT_CONFIGURE = Temp_Bool

End Function

Private Function SET_PORT_CONFIG(Port_Number As Long, Optional Port_Settings As String) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-buildcommdcba
'
' Port_Settings should have the same structure as the equivalent command-line Mode arguments for a COM Port:
' COMx[:][baud=b][parity=p][data=d][stop=s][to={on|off}][xon={on|off}][odsr={on|off}][octs={on|off}][dtr={on|off|hs}][rts={on|off|hs|tg}][idsr={on|off}]
' For example, to configure a baud rate of 1200, no parity, 8 data bits, and 1 stop bit, Port_Settings text is "baud=1200 parity=N data=8 stop=1"

Dim Temp_Bool As Boolean
Dim Temp_String(LONG_1 To LONG_7) As String
Dim Config_Length As Long, Temp_Handle As Long
Dim Temp_Text As String, Error_Text As String, Result_Text As String, Detail_Text As String, Success_Text As String, Failure_Text As String

Config_Length = Len(Port_Settings)

Temp_Text = IIf(Config_Length > LONG_4, Port_Settings, "= Default")

Temp_String(LONG_1) = " Attempting to Set Port  Mode With"
Temp_String(LONG_2) = " Settings " & Temp_Text
Temp_String(LONG_3) = " Build COM Port DCB result = "
Temp_String(LONG_4) = " Settings applied to Port"
Temp_String(LONG_5) = " Failed to apply configuration settings, Last Error "
Temp_String(LONG_6) = " Failed to get Existing Settings, Last Error "
Temp_String(LONG_7) = " Using Existing Port Settings,"

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Temp_String(LONG_1) & Temp_String(LONG_2))

Temp_Bool = GET_PORT_CONFIG(Port_Number)           ' get existing com port config (baud, parity etc.) into a device control block

If Config_Length > LONG_4 Then
   
   Temp_Text = Trim(Port_Settings)
   Temp_Handle = COM_PORT(Port_Number).Handle
 
   Temp_Bool = Build_Port_DCB(Port_Settings, COM_PORT(Port_Number).DCB)
   COM_PORT(Port_Number).DLL_Error = Err.LastDllError
 
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Temp_String(LONG_3) & Temp_Bool)

If Temp_Bool Then

   Temp_Bool = Set_Com_State(Temp_Handle, COM_PORT(Port_Number).DCB)
   COM_PORT(Port_Number).DLL_Error = Err.LastDllError
   
   Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
   Success_Text = Temp_String(LONG_4)
   Failure_Text = Temp_String(LONG_5) & Error_Text
   Result_Text = IIf(Temp_Bool, TEXT_SUCCESS, TEXT_FAILURE)
   Detail_Text = IIf(Temp_Bool, Success_Text, Failure_Text)
   
Else

   Temp_Bool = False
   Detail_Text = Temp_String(LONG_6) & Error_Text

End If

Else

Temp_Bool = True
Result_Text = TEXT_SUCCESS
Detail_Text = Temp_String(LONG_7) & GET_PORT_SETTINGS_FROM_DCB(Port_Number)

End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

SET_PORT_CONFIG = Temp_Bool

End Function

Private Function SET_PORT_VALUES(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_VALUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Effective_Byte_Count As Long
Dim Temp_Bool As Boolean, Timeslice_Byte_Count As Boolean

Const Temp_Text_1 As String = " Insufficient Read Buffer Size, Buffer Length = "
Const Temp_Text_2 As String = " Setting Timeslice Bytes per Synchronous Read / Write = "
Const Temp_Text_3 As String = " Synchronous Read Buffer Length (Max Timeslice Bytes) = "

If Len(COM_PORT(Port_Number).Buffers.Read_Buffer) > LONG_0 Then

Temp_Bool = True

COM_PORT(Port_Number).Timers.Port_Data_Time = LONG_0
COM_PORT(Port_Number).Timers.Last_Data_Time = LONG_0

COM_PORT(Port_Number).Settings = GET_PORT_SETTINGS(Port_Number)
COM_PORT(Port_Number).Timers.Frame_MicroSeconds = GET_FRAME_TIME(Port_Number)
COM_PORT(Port_Number).Timers.Frame_MilliSeconds = COM_PORT(Port_Number).Timers.Frame_MicroSeconds / LONG_1000
COM_PORT(Port_Number).Timers.Bytes_Per_Second = Int(LONG_1 / COM_PORT(Port_Number).Timers.Frame_MicroSeconds * LONG_1E6)
COM_PORT(Port_Number).Buffers.Read_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Read_Buffer)

' can change loop wait timers to suit local requirements.

COM_PORT(Port_Number).Timers.Exit_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * LONG_2
COM_PORT(Port_Number).Timers.Char_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * LONG_5
COM_PORT(Port_Number).Timers.Data_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * LONG_20
COM_PORT(Port_Number).Timers.Line_Loop_Wait = Int(LONG_1 + COM_PORT(Port_Number).Timers.Frame_MilliSeconds) * LONG_100

If COM_PORT(Port_Number).Timers.Exit_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Exit_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Char_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Char_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Data_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Data_Loop_Wait = LONG_1000
If COM_PORT(Port_Number).Timers.Line_Loop_Wait > VBA_TIMEOUT / LONG_5 Then COM_PORT(Port_Number).Timers.Line_Loop_Wait = LONG_1000

Timeslice_Byte_Count = IIf(COM_PORT(Port_Number).Timers.Bytes_Per_Second < COM_PORT(Port_Number).Buffers.Read_Buffer_Length, True, False)
Effective_Byte_Count = IIf(Timeslice_Byte_Count, COM_PORT(Port_Number).Timers.Bytes_Per_Second, COM_PORT(Port_Number).Buffers.Read_Buffer_Length)

COM_PORT(Port_Number).Timers.Timeslice_Bytes = Effective_Byte_Count

' can change read wait timers to suit local requirements.

Select Case COM_PORT(Port_Number).Timers.Bytes_Per_Second

    Case Is > LONG_1000: COM_PORT(Port_Number).Timers.Read_Wait_Time = LONG_100000
    Case Is < LONG_100:  COM_PORT(Port_Number).Timers.Read_Wait_Time = LONG_500000
    Case Else:           COM_PORT(Port_Number).Timers.Read_Wait_Time = LONG_125000

End Select


If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "READ_BYTES", COM_PORT(Port_Number).Name & Temp_Text_2 & COM_PORT(Port_Number).Timers.Timeslice_Bytes)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "BUFFER_SIZE", COM_PORT(Port_Number).Name & Temp_Text_3 & COM_PORT(Port_Number).Buffers.Read_Buffer_Length)

Else   ' read buffer too small or not fixed size

Temp_Bool = False

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "BUFFER_SIZE", COM_PORT(Port_Number).Name & Temp_Text_1 & COM_PORT(Port_Number).Buffers.Read_Buffer_Length)

End If

SET_PORT_VALUES = Temp_Bool

End Function

Public Function SHOW_PORT_VALUES(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_VALUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Const Temp_Text_1 As String = "PORT_VALUES"
Const Temp_Text_2 As String = "FRAME_TIME"
Const Temp_Text_3 As String = "SPEED"
Const Temp_Text_4 As String = "BUFFER"
Const Temp_Text_5 As String = "TIMING"
Const Temp_Text_6 As String = "SETTINGS"

Dim Temp_Bool As Boolean, Port_Name As String

Temp_Bool = False

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
    Temp_Bool = True

    Port_Name = COM_PORT(Port_Number).Name
    
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_6, Port_Name & " Standard Port Settings                = ", COM_PORT(Port_Number).Settings)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, Port_Name & " MilliSeconds per Read/Write character = ", COM_PORT(Port_Number).Timers.Frame_MilliSeconds)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, Port_Name & " MicroSeconds per Read/Write character = ", COM_PORT(Port_Number).Timers.Frame_MicroSeconds)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, Port_Name & " Read/Write speed in Bytes per Second  = ", COM_PORT(Port_Number).Timers.Bytes_Per_Second)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Exit Loop Wait Time Milliseconds      = ", COM_PORT(Port_Number).Timers.Exit_Loop_Wait)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Char Loop Wait Time Milliseconds      = ", COM_PORT(Port_Number).Timers.Char_Loop_Wait)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Data Loop Wait Time Milliseconds      = ", COM_PORT(Port_Number).Timers.Data_Loop_Wait)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Line Loop Wait Time Milliseconds      = ", COM_PORT(Port_Number).Timers.Line_Loop_Wait)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Synch. Read Timeout Microseconds      = ", COM_PORT(Port_Number).Timers.Read_Wait_Time)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, Port_Name & " Read/Write 1-Second Timeslice Bytes   = ", COM_PORT(Port_Number).Timers.Timeslice_Bytes)
    Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_4, Port_Name & " Maximum Synchronous Read Buffer Size  = ", COM_PORT(Port_Number).Buffers.Read_Buffer_Length)
    
    Else
        
        Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If

Else
    
    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_VALUES = Temp_Bool

End Function

Private Function CLOSE_PORT_HANDLE(Port_Number As Long) As Boolean
'----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "RELEASE_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'----------------------------------------------------------------------

Dim Port_Handle As Long
Dim Temp_Close As Boolean
Dim Temp_String(LONG_1 To LONG_3) As String
Dim Result_Text As String, Detail_Text As String

Port_Handle = COM_PORT(Port_Number).Handle

Temp_String(LONG_1) = " Attempting to Close Synchronous Port Handle " & Port_Handle
Temp_String(LONG_2) = " Closed Synchronous Port Handle " & Port_Handle
Temp_String(LONG_3) = " Error Closing Port, Last Error "

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "CLOSING", COM_PORT(Port_Number).Name & Temp_String(LONG_1))

Temp_Close = Com_Port_Release(Port_Handle)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Close Then

Result_Text = TEXT_SUCCESS: Detail_Text = Temp_String(LONG_2)

Else

Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_3) & Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

End If

CLOSE_PORT_HANDLE = Temp_Close

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

End Function

Public Function STOP_COM_PORT(Port_Number As Long) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "STOP_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Temp_Handle As Long
Dim Temp_Bool As Boolean
Dim Temp_String(LONG_1 To LONG_5) As String
Dim Result_Text As String, Detail_Text As String

Temp_Bool = False

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Handle = COM_PORT(Port_Number).Handle

Temp_String(LONG_1) = " Attempting to Stop COM Port " & Port_Number & " with Handle " & Temp_Handle
Temp_String(LONG_2) = " Stopped COM Port with Handle " & Temp_Handle
Temp_String(LONG_3) = " Error Closing Port with Handle " & Temp_Handle
Temp_String(LONG_4) = " Error Purging Port with Handle " & Temp_Handle
Temp_String(LONG_5) = " Failed to Stop COM Port " & Port_Number & ", Invalid Handle " & Temp_Handle

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "STOPPING", COM_PORT(Port_Number).Name & Temp_String(LONG_1))

If Temp_Handle > LONG_0 Then
If PURGE_COM_PORT(Port_Number) Then
If CLOSE_PORT_HANDLE(Port_Number) Then

     COM_PORT(Port_Number).Handle = LONG_0
     Detail_Text = Temp_String(LONG_2)
     Result_Text = TEXT_SUCCESS
     Temp_Bool = True
     
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_3): End If
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_4): End If
Else: Result_Text = TEXT_FAILURE: Detail_Text = Temp_String(LONG_5): End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

STOP_COM_PORT = Temp_Bool

End Function

Public Function WAIT_COM_PORT(Port_Number As Long, Optional Wait_MilliSeconds As Long = LONG_333) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "WAIT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Wait_Text As String
Dim Wait_Result As Boolean

Const Temp_String_1 As String = " Waiting for Receive Data, Wait Time = "
Const Temp_String_2 As String = " mS for Receive Data, Result = "
Const Temp_String_3 As String = "WAIT_START"
Const Temp_String_4 As String = "WAIT_RESULT"
Const Temp_String_5 As String = " Waited "

Wait_Result = False

If Port_Valid Then
 
     If Port_Started(Port_Number) Then
    
        If Port_Debug Then
        
        Wait_Text = Temp_String_1 & Wait_MilliSeconds & TEXT_MS
        
        Call PRINT_DEBUG_TEXT(Module_Name, Temp_String_3, COM_PORT(Port_Number).Name & Wait_Text)
      
        PORT_MICROSECONDS_NOW Port_Number

        Wait_Result = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)
        
        PORT_MICROSECONDS_END Port_Number
        
        Wait_Text = Temp_String_5 & PORT_MILLISECONDS(Port_Number) & Temp_String_2 & Wait_Result

        Call PRINT_DEBUG_TEXT(Module_Name, Temp_String_4, COM_PORT(Port_Number).Name & Wait_Text)

        Else
        
        Wait_Result = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)
        
        End If
        
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If

Else

If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

WAIT_COM_PORT = Wait_Result

End Function

Private Function SYNCHRONOUS_WAIT_COM_PORT(Port_Number As Long, Wait_MilliSeconds As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_WAIT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Data_Waiting As Boolean, Wait_Expired As Boolean, Clear_Result As Boolean, Clear_Success As Boolean
Dim Queue_Length As Long, Loop_Iteration As Long, Loop_Wait_Time As Long, Wait_Remaining As Long, Sleep_Time As Long
Dim Temp_Text As String, Error_Text As String

Const Loop_Time As Long = LONG_100

Const Temp_Text_1 As String = " Approximate Wait Time "
Const Temp_Text_2 As String = " mS, Loop Count = "
Const Temp_Text_3 As String = " Loop Countdown "
Const Temp_Text_4 As String = " Wait Time Remaining = "
Const Temp_Text_5 As String = " Receive Data Queue Length = "
Const Temp_Text_6 As String = " Synchronous Wait, Wait Time Remaining = "
Const Temp_Text_7 As String = " Clear Comms Error Failed, Last Error = "
Const Temp_Text_8 As String = " Clear Comms Error Failed, Input Queue Data not available"

Wait_Remaining = IIf(Wait_MilliSeconds < LONG_1, LONG_1, Wait_MilliSeconds)
Loop_Wait_Time = IIf(Wait_MilliSeconds < Loop_Time, Wait_Remaining, Loop_Time)
Loop_Iteration = Int(Wait_Remaining / Loop_Wait_Time) + IIf(Wait_Remaining Mod Loop_Wait_Time > LONG_0, LONG_1, LONG_0)

If Port_Debug Then Temp_Text = Temp_Text_1 & Wait_Remaining & Temp_Text_2 & Loop_Iteration
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text)

Do

Clear_Result = Clear_Com_Error(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Clear_Success = IIf(COM_PORT(Port_Number).DLL_Error = SYSTEM_ERRORS.SUCCESS, True, False)

If Clear_Success Then

    Queue_Length = COM_PORT(Port_Number).Status.QUEUE_IN
    Data_Waiting = IIf(Queue_Length > LONG_0, True, False)
    
    If Not Data_Waiting Then
    
        Wait_Expired = IIf(Wait_Remaining < LONG_1, True, False)
        
        If Not Wait_Expired Then
    
            If Port_Debug Then Temp_Text = Temp_Text_3 & Loop_Iteration & TEXT_COMMA & Temp_Text_4 & Wait_Remaining & TEXT_MS
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text)
    
            Sleep_Time = IIf(Wait_Remaining < Loop_Wait_Time, Wait_Remaining, Loop_Wait_Time)
    
            Kernel_Sleep_Milliseconds Sleep_Time
        
            Loop_Iteration = Loop_Iteration - LONG_1

            Wait_Remaining = Wait_Remaining - Sleep_Time
        
        Else
      
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMEOUT, COM_PORT(Port_Number).Name & Temp_Text_6 & Wait_Remaining & TEXT_MS)
      
        End If
       
    Else

    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RESULT, COM_PORT(Port_Number).Name & Temp_Text_5 & Queue_Length)

    End If
    
Else

    Data_Waiting = False
    Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_7 & Error_Text)
    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_8)

End If

DoEvents

Loop Until Data_Waiting Or Wait_Expired Or Not Clear_Success

SYNCHRONOUS_WAIT_COM_PORT = Data_Waiting

End Function

Public Function DEBUG_COM_PORT(Port_Number As Long, Optional Debug_State As Variant) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_DEBUG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

Dim Com_Port_String As String

Const Debug_Text_1 As String = " New Debug State = "
Const Debug_Text_2 As String = "SET_DEBUG"

Port_Debug = False

If Port_Valid Then
   
    If IsMissing(Debug_State) Then
    
        Port_Debug = Not COM_PORT(Port_Number).Debug

    Else
    
        Port_Debug = CBool(Debug_State)
    
    End If

        COM_PORT(Port_Number).Debug = Port_Debug
        Com_Port_String = TEXT_COM_PORT & CStr(Port_Number) & TEXT_COMMA & Debug_Text_1
        Call PRINT_DEBUG_TEXT(Module_Name, Debug_Text_2, Com_Port_String & Port_Debug)
Else
    
    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DEBUG_COM_PORT = Port_Debug

End Function

Public Function RECEIVE_COM_PORT(Port_Number As Long) As String

'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "RECEIVE_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Read_Return As Boolean
Dim Receive_Microseconds As Currency
Dim Bytes_Per_Second As Long, Receive_Byte_Count As Long

Const Temp_Text_01 As String = " Port Settings = "
Const Temp_Text_02 As String = " Timeslice Bytes/Second = "
Const Temp_Text_03 As String = " Synchronous Read, Result = "
Const Temp_Text_04 As String = " Read Buffer Loop, Reading "
Const Temp_Text_05 As String = " Read Buffer Zero, Waiting "
Const Temp_Text_06 As String = " Read Buffer Char, Waiting "
Const Temp_Text_07 As String = " Read Buffer Data, Waiting "
Const Temp_Text_08 As String = " Read Buffer Line, Waiting "
Const Temp_Text_09 As String = " Read Buffer Full, Looping "
Const Temp_Text_10 As String = " Effective Bytes/Second = "
Const Temp_Text_11 As String = " Last Data Microseconds = "
Const Temp_Text_12 As String = " Read Wait Microseconds = "
Const Temp_Text_13 As String = " Receive   Microseconds = "
Const Temp_Text_14 As String = " Receive   Byte Count   = "
Const Temp_Text_15 As String = " Synchronous Read Failed "

If Port_Valid Then

If Port_Started(Port_Number) Then

COM_PORT(Port_Number).Buffers.Receive_Result = vbNullString

If Port_Debug Then

    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SETTINGS, COM_PORT(Port_Number).Name & Temp_Text_01 & COM_PORT(Port_Number).Settings)
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTING, COM_PORT(Port_Number).Name & Temp_Text_02 & COM_PORT(Port_Number).Timers.Timeslice_Bytes)
    
    PORT_MICROSECONDS_NOW Port_Number

End If

    Do
       If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READING, COM_PORT(Port_Number).Name & Temp_Text_04)
       
        Do
            Read_Return = SYNCHRONOUS_READ_COM_PORT(Port_Number, COM_PORT(Port_Number).Timers.Timeslice_Bytes)
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READING, COM_PORT(Port_Number).Name & Temp_Text_03 & Read_Return)
            
            If Read_Return Then
            
            If Not COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Then
            
                COM_PORT(Port_Number).Buffers.Receive_Result = COM_PORT(Port_Number).Buffers.Receive_Result & COM_PORT(Port_Number).Buffers.Read_Result
                
                Select Case COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read
                
                Case Is < LONG_4                                         ' assume manual data entry, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SINGLE, COM_PORT(Port_Number).Name & Temp_Text_06 & COM_PORT(Port_Number).Timers.Char_Loop_Wait & TEXT_MS)
                Kernel_Sleep_Milliseconds COM_PORT(Port_Number).Timers.Char_Loop_Wait
            
                Case Is < LONG_21                                        ' assume continuous data ending, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text_07 & COM_PORT(Port_Number).Timers.Data_Loop_Wait & TEXT_MS)
                Kernel_Sleep_Milliseconds COM_PORT(Port_Number).Timers.Data_Loop_Wait
                
                Case Is = COM_PORT(Port_Number).Timers.Timeslice_Bytes   ' assume more data available immediately, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_09)
            
                Case Else                                                ' assume more data from continuous source, allow buffer to refill.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text_08 & COM_PORT(Port_Number).Timers.Line_Loop_Wait & TEXT_MS)
                Kernel_Sleep_Milliseconds COM_PORT(Port_Number).Timers.Line_Loop_Wait
                
                End Select
                                            
                DoEvents
                
            End If
                    
            Else
            
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_15)
    
            End If
                        
        Loop Until COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Or Not Read_Return
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_11 & COM_PORT(Port_Number).Timers.Last_Data_Time)
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_12 & COM_PORT(Port_Number).Timers.Read_Wait_Time)
        
        If Not COM_PORT(Port_Number).Timers.Read_Timeout Then
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "NO_DATA", COM_PORT(Port_Number).Name & Temp_Text_05 & COM_PORT(Port_Number).Timers.Exit_Loop_Wait & TEXT_MS)
        Kernel_Sleep_Milliseconds COM_PORT(Port_Number).Timers.Exit_Loop_Wait
        
        End If
                
     Loop Until COM_PORT(Port_Number).Timers.Read_Timeout Or Not Read_Return
     
    If Port_Debug Then
    
        PORT_MICROSECONDS_END Port_Number
        Receive_Microseconds = PORT_MICROSECONDS(Port_Number)
        Receive_Byte_Count = Len(COM_PORT(Port_Number).Buffers.Receive_Result)
        Bytes_Per_Second = Receive_Byte_Count / (Receive_Microseconds / LONG_1E6)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_DURATION, COM_PORT(Port_Number).Name & Temp_Text_13 & Receive_Microseconds)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RECEIVED, COM_PORT(Port_Number).Name & Temp_Text_14 & Receive_Byte_Count)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FINISHED, COM_PORT(Port_Number).Name & Temp_Text_10 & Bytes_Per_Second)
    
    End If
    
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If
      
Else

    If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If
    
RECEIVE_COM_PORT = COM_PORT(Port_Number).Buffers.Receive_Result

End Function

Public Function TRANSMIT_COM_PORT(Port_Number As Long, Transmit_Text As String) As Boolean
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "TRANSMIT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Temp_String As String
Dim Temp_Bool As Boolean, Loop_Closing As Boolean
Dim Byte_Pointer As Long, Temp_Pointer As Long, Loop_Counter As Long
Dim Bytes_Per_Second As Long, Byte_Count As Long, Transmit_Length As Long
Dim Transmit_Time As Currency

Const Temp_Text_1 As String = " Port Settings = "
Const Temp_Text_2 As String = " Timeslice Bytes/Second = "
Const Temp_Text_3 As String = " Transmitting Bytes "
Const Temp_Text_4 As String = " Transmit Time for "
Const Temp_Text_5 As String = " Bytes = "
Const Temp_Text_6 As String = " Effective Bytes per Second = "

If Port_Valid Then

If Port_Started(Port_Number) Then

Transmit_Length = Len(Transmit_Text)

If Port_Debug Then

Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SETTINGS, COM_PORT(Port_Number).Name & Temp_Text_1 & COM_PORT(Port_Number).Settings)
Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, COM_PORT(Port_Number).Name & Temp_Text_2 & COM_PORT(Port_Number).Timers.Timeslice_Bytes)

PORT_MICROSECONDS_NOW Port_Number

End If

For Loop_Counter = LONG_1 To Transmit_Length Step COM_PORT(Port_Number).Timers.Timeslice_Bytes

    Byte_Pointer = (Loop_Counter + COM_PORT(Port_Number).Timers.Timeslice_Bytes) - LONG_1
    Loop_Closing = IIf(Transmit_Length - Loop_Counter < COM_PORT(Port_Number).Timers.Timeslice_Bytes, True, False)
    Temp_Pointer = IIf(Loop_Closing, Transmit_Length, Byte_Pointer)
    Byte_Count = Temp_Pointer - Loop_Counter + LONG_1
    
    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WRITING, COM_PORT(Port_Number).Name & Temp_Text_3 & Loop_Counter & " To " & Temp_Pointer & " (" & Byte_Count & " Bytes)")
    
    COM_PORT(Port_Number).Buffers.Write_Buffer = Mid$(Transmit_Text, Loop_Counter, COM_PORT(Port_Number).Timers.Timeslice_Bytes)

    Temp_Bool = SYNCHRONOUS_WRITE_COM_PORT(Port_Number)

    DoEvents

Next Loop_Counter

If Port_Debug Then
    
    PORT_MICROSECONDS_END Port_Number
    
    Transmit_Time = PORT_MICROSECONDS(Port_Number)
    Bytes_Per_Second = Transmit_Length / Transmit_Time * LONG_1E6
    Temp_String = Temp_Text_4 & Transmit_Length & Temp_Text_5 & Int(Transmit_Time / LONG_1000) & TEXT_MS

    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, COM_PORT(Port_Number).Name & Temp_String)
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, COM_PORT(Port_Number).Name & Temp_Text_6 & Bytes_Per_Second)

End If

Else

    If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
End If

Else

    If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DoEvents

TRANSMIT_COM_PORT = Temp_Bool

End Function

Private Function GET_PORT_CONFIG(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getcommstate
' get existing com port config and write to port's device control block

Dim Temp_Bool As Boolean
Dim Temp_Settings As String, Success_Text As String, Failure_Text As String, Progress_Text As String, Error_Text As String

Temp_Bool = Get_Com_State(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Temp_Settings = GET_PORT_SETTINGS_FROM_DCB(Port_Number)

Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
Failure_Text = " Port not started, no data available, Last DLL Error " & Error_Text
Success_Text = " Existing Com Port DCB Config," & Temp_Settings
Progress_Text = IIf(Temp_Bool, Success_Text, Failure_Text)

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Progress_Text)

GET_PORT_CONFIG = Temp_Bool

End Function

Public Function GET_PORT_SETTINGS(Port_Number As Long) As String
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_SETTINGS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Port_Settings As String

If Port_Valid Then

If Port_Started(Port_Number) Then

Port_Settings = vbNullString
Port_Settings = Port_Settings & COM_PORT(Port_Number).DCB.Baud_Rate & TEXT_DASH
Port_Settings = Port_Settings & COM_PORT(Port_Number).DCB.BYTE_SIZE & TEXT_DASH
Port_Settings = Port_Settings & CONVERT_PARITY(COM_PORT(Port_Number).DCB.PARITY) & TEXT_DASH
Port_Settings = Port_Settings & CONVERT_STOPBITS(COM_PORT(Port_Number).DCB.STOP_BITS)

Else

Port_Settings = "PORT-NOT-STARTED"

End If

Else

Port_Settings = "INVALID-PORT"

End If

GET_PORT_SETTINGS = Port_Settings

End Function

Private Function SYNCHRONOUS_READ_COM_PORT(Port_Number As Long, Read_Bytes_Requested As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_READ"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Error_Text As String

Const Temp_Text_1 As String = "SYNC_READ"
Const Temp_Text_2 As String = " Synchronous Read, Bytes = "
Const Temp_Text_3 As String = " Synchronous Read, Last Error "

Temp_Bool = Synchronous_Read(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Read_Buffer, Read_Bytes_Requested, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read, ByVal vbNullString)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If COM_PORT(Port_Number).DLL_Error = SYSTEM_ERRORS.SUCCESS Then

    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Temp_Text_1, COM_PORT(Port_Number).Name & Temp_Text_2 & COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)

    If COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read = LONG_0 Then
     
        COM_PORT(Port_Number).Timers.Last_Data_Time = GET_HOST_MICROSECONDS - COM_PORT(Port_Number).Timers.Port_Data_Time
        COM_PORT(Port_Number).Timers.Read_Timeout = IIf(COM_PORT(Port_Number).Timers.Last_Data_Time > COM_PORT(Port_Number).Timers.Read_Wait_Time, True, False)
        COM_PORT(Port_Number).Buffers.Read_Result = vbNullString
        COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = True
    
    Else
        
        COM_PORT(Port_Number).Timers.Port_Data_Time = GET_HOST_MICROSECONDS
        COM_PORT(Port_Number).Timers.Last_Data_Time = LONG_0
        COM_PORT(Port_Number).Timers.Read_Timeout = False
        COM_PORT(Port_Number).Buffers.Read_Result = Left$(COM_PORT(Port_Number).Buffers.Read_Buffer, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)
        COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = False
        
    End If

Else

    Temp_Bool = False
    COM_PORT(Port_Number).Timers.Read_Timeout = True
    COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = True
    COM_PORT(Port_Number).Buffers.Read_Result = vbNullString

    Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Temp_Text_1, COM_PORT(Port_Number).Name & Temp_Text_3 & COM_PORT(Port_Number).DLL_Error & TEXT_EQUALS & Error_Text)
    
End If

SYNCHRONOUS_READ_COM_PORT = Temp_Bool

End Function

Private Function SYNCHRONOUS_WRITE_COM_PORT(Port_Number As Long) As Boolean
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_WRITE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Temp_Bool As Boolean, Write_Buffer_Length As Long, Error_Text As String

Const Sync_Write As String = "SYNC_WRITE"
Const Temp_Text_1 As String = " Synchronous Write, Last Error "
Const Temp_Text_2 As String = " Synchronous Write, Write Length  = "
Const Temp_Text_3 As String = " Synchronous Write, Bytes Written = "

Write_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Write_Buffer)

Temp_Bool = Synchronous_Write(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Write_Buffer, Write_Buffer_Length, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent, ByVal vbNullString)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If COM_PORT(Port_Number).DLL_Error = SYSTEM_ERRORS.SUCCESS Then

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Sync_Write, COM_PORT(Port_Number).Name & Temp_Text_2 & Write_Buffer_Length)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Sync_Write, COM_PORT(Port_Number).Name & Temp_Text_3 & COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent)

Else

Temp_Bool = False
Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Sync_Write, COM_PORT(Port_Number).Name & Temp_Text_1 & COM_PORT(Port_Number).DLL_Error & " = " & Error_Text)

End If

SYNCHRONOUS_WRITE_COM_PORT = Temp_Bool

End Function

Public Function SEND_COM_PORT(Port_Number As Long, Send_Variable As Variant) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SEND_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/return-values-for-the-cstr-function

Dim Temp_Bool As Boolean

Const Temp_Text_1 As String = ", Transmit Result = "

Temp_Bool = False

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
        Temp_Bool = TRANSMIT_COM_PORT(Port_Number, CStr(Send_Variable))
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RESULT, TEXT_COM_PORT & Port_Number & Temp_Text_1 & Temp_Bool)

    Else
    
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
    End If

Else

    If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SEND_COM_PORT = Temp_Bool

End Function

Public Function PUT_COM_PORT(Port_Number As Long, Put_String As String) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "PUT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Write_Byte_Count As Long
Dim Error_Text As String, Put_Com_Character As String

If Port_Valid Then

        If Port_Started(Port_Number) Then

        Put_Com_Character = Left$(Put_String, LONG_1)
        Temp_Bool = Synchronous_Write(COM_PORT(Port_Number).Handle, Put_Com_Character, LONG_1, Write_Byte_Count, ByVal vbNullString)
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError: Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

        Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
        End If

Else

    If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

PUT_COM_PORT = Temp_Bool

End Function

Public Function GET_COM_PORT(Port_Number As Long) As String
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Error_Text As String
Dim Temp_Bool As Boolean
Dim Read_Byte_Count As Long
Dim Get_Com_Character As String * LONG_1               ' must be fixed length 1

If Port_Valid Then

        If Port_Started(Port_Number) Then

        Temp_Bool = Synchronous_Read(COM_PORT(Port_Number).Handle, Get_Com_Character, LONG_1, Read_Byte_Count, ByVal vbNullString)
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError: Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

        Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
        
        End If

Else

    If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

GET_COM_PORT = Get_Com_Character

End Function

Private Function PURGE_COM_PORT(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "PURGE_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Temp_Text As String, Error_Text As String

Temp_Bool = Com_Port_Purge(COM_PORT(Port_Number).Handle, PORT_CONTROL.PURGE_ALL)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError
Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

DoEvents

If Port_Debug Then Temp_Text = COM_PORT(Port_Number).Name & " Purge All Result = " & Temp_Bool & ", Last DLL Error " & Error_Text
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "PURGE", Temp_Text)

PURGE_COM_PORT = Temp_Bool

End Function

Private Function SET_PORT_TIMERS(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_TIMERS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Bool As Boolean, Read_Bytes_Maximum As Boolean
Dim Temp_String(LONG_1 To LONG_3) As String
Dim Error_Text As String, Result_Text As String, Detail_Text As String, Success_Text As String, Failure_Text As String

Const Read_Timeout As Long = MAXDWORD
Const WRITE_CONSTANT As Long = LONG_3000

Temp_String(LONG_1) = " Port Timers Not Set, Last Error = "
Temp_String(LONG_2) = " Port Timer Values Applied to Port,"
Temp_String(LONG_3) = " Read = " & Read_Timeout & TEXT_MS & ", Write = "

COM_PORT(Port_Number).Timeouts.Read_Interval_Timeout = MAXDWORD                ' Set values as shown to indicate that timeouts are not used for file reads.
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Constant = LONG_0            '
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT
COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Multiplier = LONG_0         ' Small total values here may cause unexpected transmit data loss.

Temp_Bool = Set_Com_Timers(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Timeouts)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
Failure_Text = Temp_String(LONG_1) & Error_Text
Success_Text = Temp_String(LONG_2) & Temp_String(LONG_3) & COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant & TEXT_MS
Result_Text = IIf(Temp_Bool, TEXT_SUCCESS, TEXT_FAILURE)
Detail_Text = IIf(Temp_Bool, Success_Text, Failure_Text)

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

SET_PORT_TIMERS = Temp_Bool

End Function

Public Function SHOW_PORT_TIMERS(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_TIMERS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Error_Text As String
Dim Temp_Timer(LONG_1 To LONG_5) As String

Const Temp_Text_1 As String = "COM_PORT_TIMERS"
Const Temp_Text_2 As String = "TIMER READ"
Const Temp_Text_3 As String = "TIMER WRITE"

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Bool = Get_Com_Timers(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Timeouts)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Bool Then

Temp_Timer(LONG_1) = COM_PORT(Port_Number).Timeouts.Read_Interval_Timeout
Temp_Timer(LONG_2) = COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Constant
Temp_Timer(LONG_3) = COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Multiplier
Temp_Timer(LONG_4) = COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant
Temp_Timer(LONG_5) = COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Multiplier

Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Read Interval     ", TEXT_EQUALS_SPACE & Temp_Timer(LONG_1))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Read Constant     ", TEXT_EQUALS_SPACE & Temp_Timer(LONG_2))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Read Multiplier   ", TEXT_EQUALS_SPACE & Temp_Timer(LONG_3))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, COM_PORT(Port_Number).Name & " Write Constant    ", TEXT_EQUALS_SPACE & Temp_Timer(LONG_4))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, COM_PORT(Port_Number).Name & " Write Multiplier  ", TEXT_EQUALS_SPACE & Temp_Timer(LONG_5))

Else

Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
Call PRINT_DEBUG_TEXT(Module_Name, "ERROR", "Error retrieving Timer Settings for " & COM_PORT(Port_Number).Name & " Last Error " & Error_Text)

End If

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_TIMERS = Temp_Bool

End Function

Public Function SHOW_PORT_QUEUES(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_QUEUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Bool As Boolean

Const Temp_Text_1 As String = "COM PORT QUEUE"
Const Temp_Text_2 As String = "QUEUE_IN  "
Const Temp_Text_3 As String = "QUEUE_OUT "
Const Temp_Text_4 As String = " Input  Queue "
Const Temp_Text_5 As String = " Output Queue "
Const Temp_Text_6 As String = " Clear Comms Error Failed, Queue Data not available"

Temp_Bool = False

If Port_Valid Then

    If Port_Started(Port_Number) Then

    If CLEAR_PORT_ERROR(Port_Number) Then

    Debug.Print
    
        Temp_Bool = True
        Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & Temp_Text_4, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Status.QUEUE_IN)
        Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, COM_PORT(Port_Number).Name & Temp_Text_5, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Status.QUEUE_OUT)

    Else
        
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_6)
    
    End If
    
    Else
    
        Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
    End If
     
Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_QUEUES = Temp_Bool

End Function

Public Function CHECK_COM_PORT(Port_Number As Long) As Long
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CHECK_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' Application.Volatile  ' - remove comment mark to allow this function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_Clear As Boolean, Temp_Queue As Long, Error_Text As String

Const Temp_Text_1 As String = " Receive characters waiting to be read = "
Const Temp_Text_2 As String = " Clear Comms Error Failed, Queue Data not available"
Const Temp_Text_3 As String = " Last Error = "

Temp_Queue = LONG_NEG_1

If Port_Valid Then

If Port_Started(Port_Number) Then
    
        Temp_Clear = Clear_Com_Error(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError
        
        If Temp_Clear Then
    
            Temp_Queue = COM_PORT(Port_Number).Status.QUEUE_IN
        
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SUCCESS, COM_PORT(Port_Number).Name & Temp_Text_1 & Temp_Queue)

        Else

            Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & Temp_Text_2)
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & Temp_Text_3 & Error_Text)

        End If
        
Else

   If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If
        

Else

   If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DoEvents

CHECK_COM_PORT = Temp_Queue

End Function

Public Function SHOW_PORT_ERRORS(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_ERRORS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Port_Error As Long
Dim Temp_Bool As Boolean
Dim Error_Text As String
Dim Temp_Error(LONG_1 To LONG_5) As String

Const Temp_Text_1 As String = "COM_PORT_ERRORS"

Temp_Bool = False

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Bool = Clear_Com_Error(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError: Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

Port_Error = COM_PORT(Port_Number).Port_Errors

Temp_Error(LONG_1) = IIf(Port_Error And Port_Errors.OVERFLOW, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_2) = IIf(Port_Error And Port_Errors.OVERRUN, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_3) = IIf(Port_Error And Port_Errors.PARITY, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_4) = IIf(Port_Error And Port_Errors.FRAME, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_5) = IIf(Port_Error And Port_Errors.BREAK, TEXT_TRUE, TEXT_FALSE)

Call PRINT_SHOW_TEXT(Temp_Text_1, "OVERFLOW ", COM_PORT(Port_Number).Name & " Input Buffer Overflow       ", TEXT_EQUALS_SPACE & Temp_Error(LONG_1))
Call PRINT_SHOW_TEXT(Temp_Text_1, "OVERRUN  ", COM_PORT(Port_Number).Name & " Character Buffer Over-Run   ", TEXT_EQUALS_SPACE & Temp_Error(LONG_2))
Call PRINT_SHOW_TEXT(Temp_Text_1, "PARITY   ", COM_PORT(Port_Number).Name & " Hardware Parity Error       ", TEXT_EQUALS_SPACE & Temp_Error(LONG_3))
Call PRINT_SHOW_TEXT(Temp_Text_1, "FRAMING  ", COM_PORT(Port_Number).Name & " Hardware Framing Error      ", TEXT_EQUALS_SPACE & Temp_Error(LONG_4))
Call PRINT_SHOW_TEXT(Temp_Text_1, "BREAK    ", COM_PORT(Port_Number).Name & " Hardware Break Signal       ", TEXT_EQUALS_SPACE & Temp_Error(LONG_5))

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_ERRORS = Temp_Bool

End Function

Public Function SHOW_PORT_ALL(Port_Number As Long)
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_ALL"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

If Port_Valid Then

If Port_Started(Port_Number) Then

SHOW_PORT_DCB Port_Number
SHOW_PORT_MODEM Port_Number
SHOW_PORT_QUEUES Port_Number
SHOW_PORT_ERRORS Port_Number
SHOW_PORT_STATUS Port_Number
SHOW_PORT_TIMERS Port_Number
SHOW_PORT_VALUES Port_Number

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

End Function

Public Function SHOW_PORT_MODEM(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_MODEM"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Error_Text As String
Dim Temp_Bool As Boolean
Dim SIGNAL_CTS As String, SIGNAL_DSR As String, SIGNAL_RNG As String, SIGNAL_RLS As String

Const Temp_Text_1 As String = "COM_PORT_MODEM"
Const Temp_Text_2 As String = "MODEM (In)  "

Temp_Bool = False

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Bool = Get_Modem_Status(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Bool Then

SIGNAL_CTS = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.CTS_ON, TEXT_ON, TEXT_OFF)
SIGNAL_DSR = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.DSR_ON, TEXT_ON, TEXT_OFF)
SIGNAL_RNG = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.RING_ON, TEXT_ON, TEXT_OFF)
SIGNAL_RLS = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.RLSD_ON, TEXT_ON, TEXT_OFF)

Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Clear to Send                  CTS  ", TEXT_EQUALS_SPACE & SIGNAL_CTS)
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Data Set (Modem) Ready         DSR  ", TEXT_EQUALS_SPACE & SIGNAL_DSR)
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Ring Signal                    RING ", TEXT_EQUALS_SPACE & SIGNAL_RNG)
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Receive Line Signal Detect     RLSD ", TEXT_EQUALS_SPACE & SIGNAL_RLS)

Else

Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
Debug.Print TIMESTAMP & "Error retrieving Modem Status for " & COM_PORT(Port_Number).Name & "  Last Error = " & Error_Text

End If

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_MODEM = Temp_Bool

End Function

Private Function CLEAR_PORT_ERROR(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CLEAR_COM_ERROR"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Long As Long
Dim Temp_Bool As Boolean
Dim Error_Text As String, Result_Text As String, Detail_Text As String, Success_Text As String, Failure_Text As String

Const Temp_Text_1 As String = " Attempting to Clear Comms Error(s)"
Const Temp_Text_2 As String = " Failed to Clear Comms Error(s), "
Const Temp_Text_3 As String = " Clearing Comms Error(s), Result       = "
    
    If Port_Debug Then
    Detail_Text = COM_PORT(Port_Number).Name & Temp_Text_1
    Call PRINT_DEBUG_TEXT(Module_Name, "CLEARING", Detail_Text)

    End If
    
    Temp_Bool = Clear_Com_Error(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Port_Debug Then
    
    Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)
    Failure_Text = Temp_Text_2 & Error_Text & TEXT_SPACE
    Success_Text = Temp_Text_3 & Error_Text
    Result_Text = IIf(Temp_Bool, TEXT_SUCCESS, TEXT_FAILURE)
    Detail_Text = IIf(Temp_Bool, Success_Text, Failure_Text)
    
    Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

    End If
    
CLEAR_PORT_ERROR = Temp_Bool

End Function

Public Function SHOW_PORT_DCB(Port_Number As Long) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_DCB"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------
Dim Show_Result As Boolean
Dim DCB_TEXT(LONG_10 To LONG_36) As String
Dim DCB_VALUE(LONG_10 To LONG_36) As String
Dim DCB_LOWBITS As Long, DCB_COUNTER As Long

Const Temp_Spaces As String = "     "
Const Temp_Device As String = "DEVICE CONTROL"
Const Temp_Text_1 As String = Temp_Spaces & " 0=Disable, 1=Enable, 2=Handshake"
Const Temp_Text_2 As String = Temp_Spaces & " 0=Disable, 1=Enable, 2=Handshake, 3=Toggle"
Const Temp_Text_3 As String = Temp_Spaces & " Bits 16-32 unused"
Const Temp_Text_4 As String = Temp_Spaces & " 0=None, 1=Odd, 2=Even, 3=Mark, 4=Space"
Const Temp_Text_5 As String = Temp_Spaces & " 0=1 Stop Bit, 1=1.5 Stop Bits, 2=2 Stop Bits"

Show_Result = False

If Port_Valid Then

If Port_Started(Port_Number) Then

Show_Result = True

DCB_LOWBITS = COM_PORT(Port_Number).DCB.BIT_FIELD

DCB_TEXT(10) = " DCB Length                   ": DCB_VALUE(10) = COM_PORT(Port_Number).DCB.LENGTH_DCB
DCB_TEXT(11) = " Binary Mode                  ": DCB_VALUE(11) = IIf(DCB_LOWBITS And HEX_01, TEXT_ON, TEXT_OFF)
DCB_TEXT(12) = " Parity Checking              ": DCB_VALUE(12) = IIf(DCB_LOWBITS And HEX_02, TEXT_ON, TEXT_OFF)
DCB_TEXT(13) = " CTS Output Flow Control      ": DCB_VALUE(13) = IIf(DCB_LOWBITS And HEX_04, TEXT_ON, TEXT_OFF)
DCB_TEXT(14) = " DSR Output Flow Control      ": DCB_VALUE(14) = IIf(DCB_LOWBITS And HEX_08, TEXT_ON, TEXT_OFF)
DCB_TEXT(15) = " DTR Control Bits             ": DCB_VALUE(15) = Int((DCB_LOWBITS And HEX_30) / &HF) & Temp_Text_1
DCB_TEXT(16) = " DSR Sensitivity              ": DCB_VALUE(16) = IIf(DCB_LOWBITS And HEX_40, TEXT_ON, TEXT_OFF)
DCB_TEXT(17) = " TX Continue                  ": DCB_VALUE(17) = IIf(DCB_LOWBITS And HEX_80, TEXT_ON, TEXT_OFF)
DCB_TEXT(18) = " XON/XOFF Output Flow Control ": DCB_VALUE(18) = IIf(DCB_LOWBITS And HEX_100, TEXT_ON, TEXT_OFF)
DCB_TEXT(19) = " XON/XOFF Input Flow Control  ": DCB_VALUE(19) = IIf(DCB_LOWBITS And HEX_200, TEXT_ON, TEXT_OFF)

DCB_TEXT(20) = " Parity Error - Replace Bytes ": DCB_VALUE(20) = IIf(DCB_LOWBITS And HEX_400, TEXT_ON, TEXT_OFF)
DCB_TEXT(21) = " Discard Null Characters      ": DCB_VALUE(21) = IIf(DCB_LOWBITS And HEX_800, TEXT_ON, TEXT_OFF)
DCB_TEXT(22) = " RTS Control Bits             ": DCB_VALUE(22) = Int((DCB_LOWBITS And HEX_3000) / &HFFF) & Temp_Text_2
DCB_TEXT(23) = " Abort on Error               ": DCB_VALUE(23) = IIf(DCB_LOWBITS And HEX_4000, TEXT_ON, TEXT_OFF)
DCB_TEXT(24) = " Bits 16-32                   ": DCB_VALUE(24) = Int(DCB_LOWBITS And HEX_C000) & Temp_Text_3
DCB_TEXT(25) = " Reserved Word                ": DCB_VALUE(25) = COM_PORT(Port_Number).DCB.Reserved
DCB_TEXT(26) = " XON Limit                    ": DCB_VALUE(26) = COM_PORT(Port_Number).DCB.LIMIT_XON
DCB_TEXT(27) = " XOFF Limit                   ": DCB_VALUE(27) = COM_PORT(Port_Number).DCB.LIMIT_XOFF
DCB_TEXT(28) = " Byte Size                    ": DCB_VALUE(28) = COM_PORT(Port_Number).DCB.BYTE_SIZE
DCB_TEXT(29) = " Parity                       ": DCB_VALUE(29) = COM_PORT(Port_Number).DCB.PARITY & Temp_Text_4

DCB_TEXT(30) = " Stop Bits                    ": DCB_VALUE(30) = COM_PORT(Port_Number).DCB.STOP_BITS & Temp_Text_5
DCB_TEXT(31) = " XON Character                ": DCB_VALUE(31) = COM_PORT(Port_Number).DCB.CHAR_XON
DCB_TEXT(32) = " XOFF Character               ": DCB_VALUE(32) = COM_PORT(Port_Number).DCB.CHAR_XOFF
DCB_TEXT(33) = " Error Char                   ": DCB_VALUE(33) = COM_PORT(Port_Number).DCB.CHAR_ERROR
DCB_TEXT(34) = " End Of File Character        ": DCB_VALUE(34) = COM_PORT(Port_Number).DCB.CHAR_EOF
DCB_TEXT(35) = " Event Character              ": DCB_VALUE(35) = COM_PORT(Port_Number).DCB.CHAR_EVENT
DCB_TEXT(36) = " Reserved 1 Word              ": DCB_VALUE(36) = COM_PORT(Port_Number).DCB.RESERVED_1

For DCB_COUNTER = LONG_10 To LONG_36
Call PRINT_SHOW_TEXT(Module_Name, Temp_Device, COM_PORT(Port_Number).Name & DCB_TEXT(DCB_COUNTER), TEXT_EQUALS_SPACE & DCB_VALUE(DCB_COUNTER))
Next DCB_COUNTER

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_DCB = Show_Result

End Function

Public Function SHOW_PORT_STATUS(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_STATUS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-comstat

Dim Temp_Wait(LONG_1 To LONG_9) As String
Dim Temp_Bitmap As Long, Temp_Bool As Boolean

Const Temp_Text_1 As String = "COM_PORT_STATUS"
Const Temp_Text_2 As String = "FLOW CONTROL"
Const Temp_Text_3 As String = "END OF FILE "
Const Temp_Text_4 As String = "RX INTERRUPT"
Const Temp_Text_5 As String = "QUEUE LENGTH"

Temp_Bool = False

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Bool = CLEAR_PORT_ERROR(Port_Number)

If Temp_Bool Then

Temp_Bitmap = COM_PORT(Port_Number).Status.BIT_FIELD And HEX_7F

Temp_Wait(LONG_1) = IIf(Temp_Bitmap And HEX_01, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_2) = IIf(Temp_Bitmap And HEX_02, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_3) = IIf(Temp_Bitmap And HEX_04, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_4) = IIf(Temp_Bitmap And HEX_08, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_5) = IIf(Temp_Bitmap And HEX_10, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_6) = IIf(Temp_Bitmap And HEX_20, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_7) = IIf(Temp_Bitmap And HEX_40, TEXT_TRUE, TEXT_FALSE)
Temp_Wait(LONG_8) = COM_PORT(Port_Number).Status.QUEUE_IN
Temp_Wait(LONG_9) = COM_PORT(Port_Number).Status.QUEUE_OUT

Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Transmission Waiting for CTS     ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_1))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Transmission Waiting for DSR     ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_2))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Transmission Waiting for RLSD    ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_3))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Transmission Waiting (XOFF Hold) ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_4))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & " Transmission Waiting (XOFF Sent) ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_5))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, COM_PORT(Port_Number).Name & " End Of File (EOF) Received       ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_6))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_4, COM_PORT(Port_Number).Name & " Priority / Interrupt Character   ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_7))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, COM_PORT(Port_Number).Name & " Input Queue (awaiting ReadFile)  ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_8))
Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_5, COM_PORT(Port_Number).Name & " Output Queue (for Transmission)  ", TEXT_EQUALS_SPACE & Temp_Wait(LONG_9))

Else

Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & " Failed to Clear Com Errors")

End If

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_STATUS = Temp_Bool

End Function

Public Function DEVICE_READY(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "DEVICE_READY"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' returns True if port valid, started and COM Port DSR signal is asserted.
' DSR = Data Set Ready,from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow this function to recalculate in Excel Worksheet cell.

Dim Temp_Text As String
Dim Temp_Bool As Boolean, Signal_State As Boolean

Signal_State = False

If Port_Valid Then

    If Port_Started(Port_Number) Then

    Temp_Bool = Get_Modem_Status(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)

    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Temp_Bool Then

        Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.DSR_ON, True, False)

    Else
    
        Temp_Text = "Error Retreiving Modem Status, Last Error = " & Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "MODEM_STATUS", COM_PORT(Port_Number).Name & Temp_Text)

    End If

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DEVICE_READY = Signal_State

End Function

Public Function CLEAR_TO_SEND(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CLEAR_TO_SEND"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' returns True if port valid, started and COM Port CTS signal is asserted.
' CTS = Clear To Send, from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow this function to recalculate in Excel Worksheet cell.

Dim Temp_Text As String
Dim Temp_Bool As Boolean, Signal_State As Boolean

Signal_State = False

If Port_Valid Then

    If Port_Started(Port_Number) Then

    Temp_Bool = Get_Modem_Status(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)

    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Temp_Bool Then

        Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.CTS_ON, True, False)

    Else
    
        Temp_Text = "Error Retreiving Modem Status, Last Error = " & Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "MODEM_STATUS", COM_PORT(Port_Number).Name & Temp_Text)

    End If

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

CLEAR_TO_SEND = Signal_State

End Function

Private Function GET_PORT_SETTINGS_FROM_DCB(Port_Number As Long) As String
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_SETTINGS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Temp_Bool As Boolean
Dim Temp_String As String, Error_Text As String

If Port_Valid Then

Temp_Bool = Get_Com_State(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError
Error_Text = Decode_System_Errors(COM_PORT(Port_Number).DLL_Error)

If Temp_Bool Then

Temp_String = vbNullString
Temp_String = Temp_String & " BAUD=" & COM_PORT(Port_Number).DCB.Baud_Rate
Temp_String = Temp_String & " DATA=" & COM_PORT(Port_Number).DCB.BYTE_SIZE
Temp_String = Temp_String & " PARITY=" & CONVERT_PARITY(COM_PORT(Port_Number).DCB.PARITY)
Temp_String = Temp_String & " STOP=" & CONVERT_STOPBITS(COM_PORT(Port_Number).DCB.STOP_BITS)
'Temp_String = Temp_String & " X_IN=" & IIf(COM_PORT(Port_Number).DCB.BIT_FIELD And &H200, TEXT_ON, TEXT_OFF)
'Temp_String = Temp_String & " X_OUT=" & IIf(COM_PORT(Port_Number).DCB.BIT_FIELD And &H100, TEXT_ON, TEXT_OFF)

Else

Temp_String = "ERROR-NO-DCB-DATA"

Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & " Error in Get Com State for DCB data, Last Error " & Error_Text)

End If
   
Else

   Temp_String = "ERROR-INVALID-PORT"

End If

GET_PORT_SETTINGS_FROM_DCB = Temp_String

End Function

Private Function GET_FRAME_TIME(Port_Number As Long) As Single
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_FRAME_TIME"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Baud_Rate As Long, Length_Start As Long, Length_Data As Long, Length_Parity As Long, Length_Stop As Long
Dim Frame_Length As Long, Frame_Duration As Single
Dim Error_Text As String, Temp_Text As String

Baud_Rate = COM_PORT(Port_Number).DCB.Baud_Rate

Length_Start = LONG_1
Length_Data = COM_PORT(Port_Number).DCB.BYTE_SIZE
Length_Parity = IIf(COM_PORT(Port_Number).DCB.PARITY = LONG_0, LONG_0, LONG_1)
Length_Stop = IIf(COM_PORT(Port_Number).DCB.STOP_BITS = LONG_0, LONG_1, LONG_2)

Frame_Length = Length_Start + Length_Data + Length_Parity + Length_Stop

Frame_Duration = Frame_Length / Baud_Rate * LONG_1E6   ' frame (character) duration in microseconds

Temp_Text = " Baud Rate=" & Baud_Rate & ", Frame Length=" & Frame_Length & ", Frame Duration=" & Frame_Duration

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, "FRAME_INFO", COM_PORT(Port_Number).Name & Temp_Text & " uS ")

GET_FRAME_TIME = Frame_Duration

End Function

Public Function Port_Number_Valid(Port_Number As Long) As Boolean

Port_Number_Valid = IIf((Port_Number < COM_PORT_MIN) Or (Port_Number > COM_PORT_MAX), False, True)

End Function

Private Function Port_Started(Port_Number As Long) As Boolean

Port_Started = IIf(COM_PORT(Port_Number).Handle > LONG_0, True, False)

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

Private Sub PORT_MICROSECONDS_NOW(Port_Number As Long)

QPC COM_PORT(Port_Number).Timers.Timing_QPC_Now

End Sub

Private Sub PORT_MICROSECONDS_END(Port_Number As Long)

QPC COM_PORT(Port_Number).Timers.Timing_QPC_End

End Sub

Private Function DELTA_MICROSECONDS(Port_Number As Long) As Currency

DELTA_MICROSECONDS = COM_PORT(Port_Number).Timers.Timing_QPC_End - COM_PORT(Port_Number).Timers.Timing_QPC_Now

End Function

Private Function PORT_MICROSECONDS(Port_Number As Long) As Currency

PORT_MICROSECONDS = Int(DELTA_MICROSECONDS(Port_Number) * LONG_1000)

End Function

Private Function PORT_MILLISECONDS(Port_Number As Long) As Long

PORT_MILLISECONDS = Int(DELTA_MICROSECONDS(Port_Number))

End Function

Public Function TIMESTAMP() As String

Dim Local_System_Time As SYSTEMTIME

Get_System_Time Local_System_Time

TIMESTAMP = Extend_String(Time() & TEXT_DOT & Local_System_Time.Milliseconds, LONG_14)

End Function

Public Function GET_HOST_MILLISECONDS() As Long

' Application.Volatile  ' optional - remove comment mark to allow this function to recalculate in Excel worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MILLISECONDS = Int(Temp_QPC)

End Function

Public Function GET_HOST_MICROSECONDS() As Currency

' Application.Volatile  ' optional - remove comment mark to allow this function to recalculate in Excel worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter

Const QPF As Long = LONG_1000

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MICROSECONDS = Int(Temp_QPC * QPF)

End Function

Public Function Extend_String(Input_String As Variant, Return_Length As Long) As String

Dim Input_Length As Long, Extend_Text As String
Dim Delta_Length As Long, Spaces_Length As Long

Dim Spaces_String As String * LONG_100

Input_String = CStr(Input_String)
Input_Length = Len(Input_String)
Delta_Length = Return_Length - Input_Length

Spaces_Length = IIf(Delta_Length > LONG_0, Delta_Length, LONG_0)
Extend_Text = IIf(Return_Length > Input_Length, Left$(Spaces_String, Spaces_Length), vbNullString)

Extend_String = Input_String & Extend_Text

End Function

Public Sub DECODE_PORT_ERRORS(ERROR_DATA As Long)

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-clearcommerror

Debug.Print TIMESTAMP & "Input Buffer Overflow             = " & IIf(ERROR_DATA And PORT_EVENT.RX_80_FULL, True, False)
Debug.Print TIMESTAMP & "Character Buffer Over-Run         = " & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & "Hardware Parity Error             = " & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & "Hardware Framing Error            = " & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & "Hardware Break Signal             = " & IIf(ERROR_DATA And PORT_EVENT.BREAK, True, False)

End Sub

Public Function Decode_System_Errors(Error_Code As Long) As String

Dim Temp_String As String

Select Case Error_Code

Case SYSTEM_ERRORS.SUCCESS:              Temp_String = "SUCCESS"
Case SYSTEM_ERRORS.NO_ACCESS:            Temp_String = "NO_ACCESS"
Case SYSTEM_ERRORS.IO_PENDING:           Temp_String = "IO PENDING"
Case SYSTEM_ERRORS.IO_INCOMPLETE:        Temp_String = "IO INCOMPLETE"
Case SYSTEM_ERRORS.INVALID_DATA:         Temp_String = "INVALID DATA"
Case SYSTEM_ERRORS.ACCESS_DENIED:        Temp_String = "ACCESS DENIED"
Case SYSTEM_ERRORS.PATH_NOT_FOUND:       Temp_String = "PATH NOT FOUND"
Case SYSTEM_ERRORS.FILE_NOT_FOUND:       Temp_String = "PORT NOT FOUND"
Case SYSTEM_ERRORS.DEVICE_NOT_READY:     Temp_String = "DEVICE NOT READY"
Case SYSTEM_ERRORS.INVALID_HANDLE:       Temp_String = "INVALID HANDLE"
Case SYSTEM_ERRORS.INVALID_FUNCTION:     Temp_String = "INVALID FUNCTION"
Case SYSTEM_ERRORS.INVALID_PARAMETER:    Temp_String = "INVALID PARAMETER"
Case SYSTEM_ERRORS.OPERATION_ABORTED:    Temp_String = "OPERATION ABORTED"
Case SYSTEM_ERRORS.TOO_MANY_OPEN_FILES:  Temp_String = "TOO MANY OPEN FILES"
Case SYSTEM_ERRORS.INSUFFICIENT_BUFFER:  Temp_String = "INSUFFICIENT BUFFER"

Case Else: Temp_String = "UNKNOWN SYSTEM ERROR CODE " & Error_Code

End Select

Decode_System_Errors = Temp_String

End Function

Private Sub PRINT_STOPPED_TEXT(Module_Text As String, Port_Number As Long)

Const Temp_Text_1 As String = "PORT_STOPPED"
Const Temp_Text_2 As String = "COM Port "
Const Temp_Text_3 As String = ", Port Not Started"

Call PRINT_DEBUG_TEXT(Module_Text, Temp_Text_1, Temp_Text_2 & Port_Number & Temp_Text_3)

End Sub

Private Sub PRINT_INVALID_TEXT(Module_Text As String, Port_Number As Long)

Const Temp_Text_1 As String = "INVALID_PORT"
Const Temp_Text_2 As String = "Port Number "
Const Temp_Text_3 As String = " Invalid, Defined Port Range = "

Call PRINT_DEBUG_TEXT(Module_Text, Temp_Text_1, Temp_Text_2 & Port_Number & Temp_Text_3 & COM_PORT_RANGE)

End Sub

Private Sub PRINT_DEBUG_TEXT(Module_Text As String, Result_Text As String, Message_Text As String)

Const COLUMN_WIDTH_1 As Long = 18
Const COLUMN_WIDTH_2 As Long = 18

Debug.Print TIMESTAMP & Extend_String(Module_Text, COLUMN_WIDTH_1) & Extend_String(Result_Text, COLUMN_WIDTH_2) & Message_Text

End Sub

Private Sub PRINT_SHOW_TEXT(DEVICE_TEXT As String, Prefix_Text As String, Detail_Text As String, Result_Text As Variant)

Const COLUMN_WIDTH_1 As Long = 18
Const COLUMN_WIDTH_2 As Long = 18
Const COLUMN_WIDTH_3 As Long = 50

Dim Temp_Text As String

Temp_Text = Extend_String(DEVICE_TEXT, COLUMN_WIDTH_1) & Extend_String(Prefix_Text, COLUMN_WIDTH_2) & Extend_String(Detail_Text, COLUMN_WIDTH_3) & CStr(Result_Text)

Debug.Print TIMESTAMP & Temp_Text

End Sub

Public Function CONVERT_PARITY(DCB_PARITY As Byte) As String

Dim Temp_String As String

Select Case DCB_PARITY

Case PORT_FRAMING.PARITY_ODD: Temp_String = "O"
Case PORT_FRAMING.PARITY_NONE: Temp_String = "N"
Case PORT_FRAMING.PARITY_EVEN: Temp_String = "E"
Case PORT_FRAMING.PARITY_MARK: Temp_String = "M"
Case PORT_FRAMING.PARITY_SPACE: Temp_String = "S"

Case Else: Temp_String = "?"

End Select

CONVERT_PARITY = Temp_String

End Function

Public Function CONVERT_STOPBITS(DCB_STOPBITS As Byte) As String

Dim Temp_String As String

Select Case DCB_STOPBITS

Case PORT_FRAMING.STOP_BITS_ONE: Temp_String = "1"
Case PORT_FRAMING.STOP_BITS_TWO: Temp_String = "2"
Case PORT_FRAMING.STOP_BITS_1P5: Temp_String = "1.5"

Case Else: Temp_String = "?"

End Select

CONVERT_STOPBITS = Temp_String

End Function

Public Function CONVERT_LINE_ERROR(LINE_ERROR As Byte) As String

Dim Temp_String As String

Select Case LINE_ERROR

Case Port_Errors.BREAK: Temp_String = "BREAK"
Case Port_Errors.FRAME: Temp_String = "FRAME"
Case Port_Errors.OVERFLOW: Temp_String = "OVERFLOW"
Case Port_Errors.OVERRUN: Temp_String = "OVERRUN"
Case Port_Errors.PARITY: Temp_String = "PARITY"

Case Else: Temp_String = "UNKNOWN"

End Select

CONVERT_LINE_ERROR = Temp_String

End Function

Public Function TEMPLATE(Port_Number As Long) As String
'----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "TEMPLATE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'----------------------------------------------------------------------

If Port_Valid Then

If Port_Started(Port_Number) Then




Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

If Port_Debug Then Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

TEMPLATE = COM_PORT(Port_Number).Name & " function result goes here "

End Function

Public Function EXAMPLE(Port_Number As Long) As String

' Example showing how to read data from a theoretical digital voltmeter with a serial port connected to COM Port 1
' To demonstrate, connect a terminal emulator or similar device to the local COM Port 1 on this machine
' From the VBA Immediate Window (Control-G), type ?EXAMPLE(1) and wait for MEASURE VOLTAGE to appear on the emulator
' When it appears, respond immediately with a reply e.g. 1234. This should display after a short delay in the VBA window.
' The function could be called from a larger VBA routine to populate Excel cells or a Word Document with readings etc.
' COM Port can optionally be started with parameters - e.g. START_COM_PORT(Port_Number, "Baud=1200 Data=7 Parity=E")
' Note that VBA remains responsive during wait_for_com function, and also during any extended read/write activites.

Dim VOLTS As String

Const READ_VOLTS_COMMAND As String = "MEASURE VOLTAGE" & vbCr

Const Temp_Text_1 As String = "Sending command string to device on COM Port "
Const Temp_Text_2 As String = "Waiting for response from device on COM Port "
Const Temp_Text_3 As String = "Timed out waiting for new data from COM Port "
Const Temp_Text_4 As String = "Failed to Start COM Port "
Const Temp_Text_5 As String = "Example function complete , closing COM Port "

DEBUG_COM_PORT 1, False

'DEBUG_COM_PORT 1, True                                           ' optional - shows port activities and wait countdown loop counter

If START_COM_PORT(Port_Number) Then                               ' continue if port starts correctly.

    Kernel_Sleep_Milliseconds 200                                 ' allow ports to stabilise after opening (optional)

    Debug.Print TIMESTAMP & Temp_Text_1 & Port_Number
    
    TRANSMIT_COM_PORT Port_Number, READ_VOLTS_COMMAND             ' send command to remote device
    
    Debug.Print TIMESTAMP & Temp_Text_2 & Port_Number
    
    If WAIT_COM_PORT(Port_Number, 10000) Then                     ' wait up to 10 seconds (without blocking VBA) for first character

        Kernel_Sleep_Milliseconds 1000                            ' wait 1 second for more characters to arrive before reading (can adjust for your device)

        VOLTS = RECEIVE_COM_PORT(Port_Number)                     ' read response back into string variable VOLTS

    Else

        Debug.Print TIMESTAMP & Temp_Text_3 & Port_Number

    End If


Else

    Debug.Print TIMESTAMP & Temp_Text_4 & Port_Number

End If

Debug.Print TIMESTAMP & Temp_Text_5 & Port_Number
Debug.Print

Kernel_Sleep_Milliseconds 200                                    ' allow ports to stabilise before closing (optional)

STOP_COM_PORT Port_Number

EXAMPLE = VOLTS                                                  ' return read result back to variable for subsequent use.

End Function

