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
Private Const LONG_6  As Long = 6
Private Const LONG_7  As Long = 7
Private Const LONG_8  As Long = 8
Private Const LONG_9  As Long = 9
Private Const LONG_10 As Long = 10
Private Const LONG_14 As Long = 14
Private Const LONG_18 As Long = 18
Private Const LONG_20 As Long = 20
Private Const LONG_21 As Long = 21
Private Const LONG_30 As Long = 30
Private Const LONG_36 As Long = 36
Private Const LONG_40 As Long = 40
Private Const LONG_50 As Long = 50
Private Const LONG_52 As Long = 52
Private Const LONG_54 As Long = 54
Private Const LONG_60 As Long = 60

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

Private Const TEXT_ON As String = "On"              ' some text string constants for minor gains.
Private Const TEXT_MS As String = " mS"
Private Const TEXT_US As String = " µS"
Private Const TEXT_TO As String = " To "
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

Private Const TEXT_ERROR  As String = "ERROR"
Private Const TEXT_CONFIG As String = "CONFIG"
Private Const TEXT_FAILED As String = "FAILED"
Private Const TEXT_RESULT As String = "RESULT"
Private Const TEXT_SINGLE As String = "SINGLE"
Private Const TEXT_TIMING As String = "TIMING"

Private Const TEXT_CLEARED As String = "CLEARED"
Private Const TEXT_CLOSING As String = "CLOSING"
Private Const TEXT_FAILURE As String = "FAILURE"
Private Const TEXT_LOOPING As String = "LOOPING"
Private Const TEXT_NO_DATA As String = "NO_DATA"
Private Const TEXT_READING As String = "READING"
Private Const TEXT_STARTED As String = "STARTED"
Private Const TEXT_STARTUP As String = "STARTUP"
Private Const TEXT_SUCCESS As String = "SUCCESS"
Private Const TEXT_TIMEOUT As String = "TIMEOUT"
Private Const TEXT_WAITING As String = "WAITING"
Private Const TEXT_WRITING As String = "WRITING"

Private Const TEXT_CLEARING As String = "CLEARING"
Private Const TEXT_DURATION As String = "DURATION"
Private Const TEXT_FINISHED As String = "FINISHED"
Private Const TEXT_RECEIVED As String = "RECEIVED"
Private Const TEXT_SETTINGS As String = "SETTINGS"
Private Const TEXT_STARTING As String = "STARTING"
Private Const TEXT_STOPPING As String = "STOPPING"

Private Const TEXT_COM_PORT As String = "COM Port "
Private Const COM_PORT_RANGE As String = COM_PORT_MIN & " to " & COM_PORT_MAX

Private Type SYSTEMTIME

             Year    As Integer
             Month   As Integer
             WeekDay As Integer
             Day     As Integer
             Hour    As Integer
             Minute  As Integer
             Second  As Integer
             MilliSeconds As Integer                      ' used for debug timestamp
End Type

Private Type DEVICE_CONTROL_BLOCK                         ' DCB  - Check latest Microsoft documentation

             LENGTH_DCB As Long
             BAUD_RATE  As Long
             BIT_FIELD  As Long
             RESERVED_0 As Integer
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

             Name As String
             Debug As Boolean
             Handle As LongPtr
             DLL_Error As Long
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

Private Declare PtrSafe Sub Get_System_Time Lib "Kernel32.dll" Alias "GetSystemTime" (ByRef System_Time As SYSTEMTIME)
Private Declare PtrSafe Sub Kernel_Sleep_MilliSeconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_MilliSeconds As Long)
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Boolean
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_Frequency As Currency) As Boolean

' https://docs.microsoft.com/en-us/windows/win32/devio/communications-functions

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
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "START_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------
' Port_Settings if supplied should have the same structure as the equivalent command-line Mode arguments for a COM Port:
' [baud=b][parity=p][data=d][stop=s][to={on|off}][xon={on|off}][odsr={on|off}][octs={on|off}][dtr={on|off|hs}][rts={on|off|hs|tg}][idsr={on|off}]
' For example, to configure a baud rate of 1200, no parity, 8 data bits, and 1 stop bit, Port_Settings text is "baud=1200 parity=N data=8 stop=1"
' Alternatively, can assign port settings to a VBA String constant or variable (e.g. SCANNER) and call start_com_port(1,SCANNER) from VBA routine

Dim Temp_Result As Boolean, Port_Handle As LongPtr
Dim Temp_Port_Name As String, Result_Text As String, Detail_Text As String

Const Start_Text_1 As String = " Attempting to Start and Configure COM Port "
Const Start_Text_2 As String = " Started and Configured COM Port "
Const Start_Text_3 As String = " with Handle "
Const Start_Text_4 As String = " Failed to Configure COM Port "
Const Start_Text_5 As String = " Failed to Open and Create COM Port "
Const Start_Text_6 As String = " Failed to Start COM Port, Existing Port Handle = "
Const Start_Text_7 As String = " Port Number Invalid, Defined Port Number Range = "

Result_Text = TEXT_FAILURE

Temp_Port_Name = TEXT_COM_PORT & CStr(Port_Number) & TEXT_COMMA

If Port_Valid Then

Port_Handle = COM_PORT(Port_Number).Handle                 ' existing port handle if port open

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTING, Temp_Port_Name & Start_Text_1 & Port_Number)
    
If COM_PORT_CLOSED(Port_Number) Then
If COM_PORT_CREATE(Port_Number) Then
If COM_PORT_CONFIGURE(Port_Number, Port_Setttings) Then
        
        PURGE_BUFFERS Port_Number
        
        Temp_Result = True
        Result_Text = TEXT_SUCCESS
        Port_Handle = COM_PORT(Port_Number).Handle          ' new port handle
        Detail_Text = Start_Text_2 & Port_Number & Start_Text_3 & Port_Handle

Else
        STOP_COM_PORT Port_Number                           ' close com port if configure failed
        
        Detail_Text = Start_Text_4 & Port_Number
End If
            
Else:   Detail_Text = Start_Text_5 & Port_Number:       End If
Else:   Detail_Text = Start_Text_6 & Port_Handle:       End If
Else:   Detail_Text = Start_Text_7 & COM_PORT_RANGE:    End If

If Port_Debug Or Not Port_Valid Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, Temp_Port_Name & Detail_Text)

START_COM_PORT = Temp_Result

End Function

Private Function COM_PORT_CREATE(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "COM_PORT_CREATE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Handle As LongPtr
Dim Temp_Result As Boolean
Dim CREATE_FILE_FLAGS As Long
Dim Device_Path As String, Error_Text As String, Result_Text As String, Detail_Text As String, Temp_Name As String

Const DEVICE_PREFIX As String = "\\.\COM"
Const Create_Text_1 As String = "CREATING"
Const Create_Text_2 As String = "PORT_MODE"
Const Create_Text_3 As String = " Attempting to Open Port with Device Path "
Const Create_Text_4 As String = " Port Open for Exclusive Access, Handle = "
Const Create_Text_5 As String = " Failed to Open COM Port, "
Const Create_Text_6 As String = " Creating Synchronous (non-overlapped) mode Port "

Device_Path = DEVICE_PREFIX & CStr(Port_Number)

CREATE_FILE_FLAGS = PORT_FILE_FLAGS.SYNCHRONOUS_MODE

Temp_Name = TEXT_COM_PORT & CStr(Port_Number) & TEXT_COMMA

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Create_Text_1, Temp_Name & Create_Text_3 & Device_Path)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Create_Text_2, Temp_Name & Create_Text_6)

Temp_Handle = Com_Port_Open(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, CREATE_FILE_FLAGS)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Select Case Temp_Handle

Case HANDLE_INVALID

    Temp_Result = False
    Result_Text = TEXT_FAILURE
    Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
    Detail_Text = Create_Text_5 & Error_Text
    COM_PORT(Port_Number).Name = vbNullString
    COM_PORT(Port_Number).Handle = LONG_0

Case Else

    Temp_Result = True
    Result_Text = TEXT_SUCCESS
    Detail_Text = Create_Text_4 & Temp_Handle
    COM_PORT(Port_Number).Name = Temp_Name
    COM_PORT(Port_Number).Handle = Temp_Handle

End Select

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, Temp_Name & Detail_Text)

COM_PORT_CREATE = Temp_Result

End Function

Private Function COM_PORT_CONFIGURE(Port_Number As Long, Optional Port_Settings As String) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "COM_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Clean_Settings As String, Selected_Text As String, Result_Text As String, Detail_Text As String

Const Config_Text_1 As String = " Attempting to Configure Port With Supplied Settings "
Const Config_Text_2 As String = " Attempting to Configure Port With Existing Settings "
Const Config_Text_3 As String = " Configured COM Port with Settings = "
Const Config_Text_4 As String = " Failed to Set Port Values "
Const Config_Text_5 As String = " Failed to Set Port Timers "
Const Config_Text_6 As String = " Failed to Set Port with Supplied Settings, "

Result_Text = TEXT_FAILURE

Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)

Selected_Text = IIf(Len(Clean_Settings) > LONG_4, Config_Text_1, Config_Text_2)

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTUP, COM_PORT(Port_Number).Name & Selected_Text)

If SET_PORT_CONFIG(Port_Number, Clean_Settings) Then
If SET_PORT_TIMERS(Port_Number) Then
If SET_PORT_VALUES(Port_Number) Then

      Temp_Result = True
      Result_Text = TEXT_SUCCESS
      Detail_Text = Config_Text_3 & GET_PORT_SETTINGS(Port_Number)
      
Else: Detail_Text = Config_Text_4: End If
Else: Detail_Text = Config_Text_5: End If
Else: Detail_Text = Config_Text_6 & Port_Settings: End If
     
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

COM_PORT_CONFIGURE = Temp_Result

End Function

Private Function SET_PORT_CONFIG(Port_Number As Long, Optional Port_Settings As String) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-buildcommdcba

Dim Temp_Result As Boolean, Temp_Build As Boolean, New_Settings As Boolean
Dim Temp_String As String, Result_Text As String, Detail_Text As String

Const Set_Text_1 As String = " Attempting to Set Port With Supplied Settings "
Const Set_Text_2 As String = " Attempting to use Existing Port Settings, "
Const Set_Text_3 As String = " Building new Device Control Block (DCB), result = "
Const Set_Text_4 As String = " Supplied Settings applied to Port = "
Const Set_Text_5 As String = " Using Existing Port Settings, "
Const Set_Text_6 As String = " Failed to build DCB, "
Const Set_Text_7 As String = " Failed to apply configuration settings, "
Const Set_Text_8 As String = " Failed to get existing Device Control Block (DCB) Settings "

If GET_PORT_CONFIG(Port_Number) Then          ' get existing com port config (baud, parity etc.) into device control block

    New_Settings = IIf(Len(Port_Settings) > LONG_4, True, False)
    
    Temp_String = IIf(New_Settings, Set_Text_1 & Port_Settings, Set_Text_2 & GET_PORT_SETTINGS(Port_Number))

    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Temp_String)

    If New_Settings Then

        Temp_Build = Build_Port_DCB(Port_Settings, COM_PORT(Port_Number).DCB)
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Set_Text_3 & Temp_Build)
        
        If Temp_Build Then
        
            Temp_Result = Apply_Port_DCB(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
            COM_PORT(Port_Number).DLL_Error = Err.LastDllError
                 
            If Temp_Result Then
            
                Result_Text = TEXT_SUCCESS
                Detail_Text = Set_Text_4 & GET_PORT_SETTINGS(Port_Number)
            
            Else
                
                Result_Text = TEXT_FAILURE
                Detail_Text = Set_Text_7 & DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
            
            End If
                       
        Else
        
            Temp_Result = False
            Result_Text = TEXT_FAILURE
            Detail_Text = Set_Text_6 & DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
        
        End If
        
    Else

        Temp_Result = True
        Result_Text = TEXT_SUCCESS
        Detail_Text = Set_Text_5 & GET_PORT_SETTINGS_FROM_DCB(Port_Number)

    End If

Else

    Temp_Result = False
    Result_Text = TEXT_FAILURE
    Detail_Text = Set_Text_8
   
End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

SET_PORT_CONFIG = Temp_Result

End Function

Private Function SET_PORT_VALUES(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_VALUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------
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

Const TEXT_READ_BYTES As String = "READ_BYTES"
Const TEXT_READ_BUFFER As String = "READ_BUFFER"

Const Values_Text_1 As String = " Insufficient Read Buffer Size, Buffer Length = "
Const Values_Text_2 As String = " Synchronous Read Buffer Length (Max Timeslice Bytes) = "
Const Values_Text_3 As String = " Setting Timeslice Bytes per Synchronous Read / Write = "

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

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READ_BUFFER, COM_PORT(Port_Number).Name & Values_Text_2 & Read_Buffer_Length)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READ_BYTES, COM_PORT(Port_Number).Name & Values_Text_3 & Effective_Byte_Count)

Else   ' read buffer size not > 0

Temp_Result = False

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READ_BUFFER, COM_PORT(Port_Number).Name & Values_Text_1 & Read_Buffer_Length)

End If

SET_PORT_VALUES = Temp_Result

End Function

Public Function SHOW_PORT_VALUES(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_VALUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Const Print_Text_1 As String = "PORT_VALUES"
Const Print_Text_2 As String = "FRAME_TIME"
Const Print_Text_3 As String = "SPEED"
Const Print_Text_4 As String = "BUFFER"
Const Print_Text_5 As String = "TIMING"
Const Print_Text_6 As String = "SETTINGS"

Const Value_Text_01 As String = " Standard Port Settings                "
Const Value_Text_02 As String = " MilliSeconds per Read/Write character "
Const Value_Text_03 As String = " MicroSeconds per Read/Write character "
Const Value_Text_04 As String = " Read/Write speed in Bytes per Second  "
Const Value_Text_05 As String = " Exit Loop Wait Time MilliSeconds      "
Const Value_Text_06 As String = " Char Loop Wait Time MilliSeconds      "
Const Value_Text_07 As String = " Data Loop Wait Time MilliSeconds      "
Const Value_Text_08 As String = " Line Loop Wait Time MilliSeconds      "
Const Value_Text_09 As String = " Synch. Read Timeout MicroSeconds      "
Const Value_Text_10 As String = " Read/Write 1-Second Timeslice Bytes   "
Const Value_Text_11 As String = " Maximum Synchronous Read Buffer Size  "

Dim Temp_Result As Boolean, Port_Name As String

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
    Temp_Result = True

    Port_Name = COM_PORT(Port_Number).Name
    
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_6, Port_Name & Value_Text_01, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Settings)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Value_Text_02, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Frame_MilliSeconds)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Value_Text_03, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Frame_MicroSeconds)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_3, Port_Name & Value_Text_04, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Bytes_Per_Second)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_05, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Exit_Loop_Wait)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_06, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Char_Loop_Wait)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_07, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Data_Loop_Wait)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_08, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Line_Loop_Wait)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_09, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Read_Wait_Time)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Value_Text_10, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Timers.Timeslice_Bytes)
    Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_4, Port_Name & Value_Text_11, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Buffers.Read_Buffer_Length)
    
    Else
        
    Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If

Else
    
    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_VALUES = Temp_Result

End Function

Private Function CLOSE_PORT_HANDLE(Port_Number As Long) As Boolean
'----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "RELEASE_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'----------------------------------------------------------------------

Dim Port_Handle As LongPtr
Dim Temp_Result As Boolean
Dim Result_Text As String, Detail_Text As String

Const Close_Text_1 As String = " Attempting to Close Synchronous Port Handle "
Const Close_Text_2 As String = " Closed Synchronous Port Handle "
Const Close_Text_3 As String = " Error Closing Port, "

Port_Handle = COM_PORT(Port_Number).Handle

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CLOSING, COM_PORT(Port_Number).Name & Close_Text_1 & Port_Handle)

Temp_Result = Com_Port_Close(Port_Handle)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

    Result_Text = TEXT_SUCCESS
    Detail_Text = Close_Text_2 & Port_Handle

Else

    Result_Text = TEXT_FAILURE
    Detail_Text = Close_Text_3 & DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
    
End If

CLOSE_PORT_HANDLE = Temp_Result

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

End Function

Public Function STOP_COM_PORT(Port_Number As Long) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "STOP_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Temp_Handle As LongPtr
Dim Temp_Result As Boolean
Dim Result_Text As String, Detail_Text As String, Temp_Name As String

Const Stop_Text_1 As String = " Attempting to Stop COM Port "
Const Stop_Text_2 As String = " Stopped COM Port with Handle "
Const Stop_Text_3 As String = " Error Closing Port with Handle "
Const Stop_Text_4 As String = " with Handle "

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Name = COM_PORT(Port_Number).Name
Temp_Handle = COM_PORT(Port_Number).Handle

If Port_Debug Then
    Detail_Text = Temp_Name & Stop_Text_1 & Port_Number & Stop_Text_4 & Temp_Handle
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STOPPING, Detail_Text)
End If

PURGE_BUFFERS Port_Number
PURGE_COM_PORT Port_Number

If CLOSE_PORT_HANDLE(Port_Number) Then

    COM_PORT(Port_Number).Name = vbNullString
    COM_PORT(Port_Number).Handle = LONG_0
    Detail_Text = Stop_Text_2 & Temp_Handle
    Result_Text = TEXT_SUCCESS
    Temp_Result = True
     
Else

    Temp_Result = False
    Result_Text = TEXT_FAILURE
    Detail_Text = Stop_Text_3 & Temp_Handle

End If

        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, Temp_Name & Detail_Text)

Else
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
End If

Else
        Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
End If

STOP_COM_PORT = Temp_Result

End Function

Public Function WAIT_COM_PORT(Port_Number As Long, Optional Wait_MilliSeconds As Long = LONG_333) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "WAIT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Wait_String As String
Dim WAIT_RESULT As Boolean

Const Wait_String_1 As String = " Waiting for Receive Data, Wait Time = "
Const Wait_String_2 As String = " mS for Receive Data, Result = "
Const Wait_String_3 As String = " Waited "

If Port_Valid Then
 
    If Port_Started(Port_Number) Then
    
        If Port_Debug Then
        
            Wait_String = Wait_String_1 & Wait_MilliSeconds & TEXT_MS
            Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTED, COM_PORT(Port_Number).Name & Wait_String)
            
            PORT_MICROSECONDS_NOW Port_Number
            
            WAIT_RESULT = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)
            
            PORT_MICROSECONDS_END Port_Number
            
            Wait_String = Wait_String_3 & PORT_MILLISECONDS(Port_Number) & Wait_String_2 & WAIT_RESULT
            Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RESULT, COM_PORT(Port_Number).Name & Wait_String)

        Else
        
            WAIT_RESULT = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)
        
        End If
        
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If

Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

WAIT_COM_PORT = WAIT_RESULT

End Function

Private Function SYNCHRONOUS_WAIT_COM_PORT(Port_Number As Long, Wait_MilliSeconds As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_WAIT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Debug_Text As String, Error_Text As String, Port_Name As String
Dim Data_Waiting As Boolean, Wait_Expired As Boolean, Clear_Result As Boolean
Dim Loop_Iteration As Long, Wait_Remaining As Long, Queue_Length As Long
Dim Loop_Remainder As Long, Loop_Wait_Time As Long, Sleep_Time As Long

Const Loop_Time As Long = LONG_100                        ' MilliSeconds

Const Wait_Text_1 As String = " Approximate Wait Time "
Const Wait_Text_2 As String = " mS, Loop Count = "
Const Wait_Text_3 As String = " Loop Countdown "
Const Wait_Text_4 As String = " Wait Time Remaining = "
Const Wait_Text_5 As String = " Receive Data Queue Length = "
Const Wait_Text_6 As String = " Synchronous Wait, Wait Time Remaining = "
Const Wait_Text_7 As String = " Clear Comms Error Failed, "
Const Wait_Text_8 As String = " Clear Comms Error Failed, Input Queue Data not available"

Wait_Remaining = IIf(Wait_MilliSeconds < LONG_1, LONG_1, Wait_MilliSeconds)
Loop_Wait_Time = IIf(Wait_MilliSeconds < Loop_Time, Wait_Remaining, Loop_Time)
Loop_Remainder = IIf(Wait_Remaining Mod Loop_Wait_Time > LONG_0, LONG_1, LONG_0)
Loop_Iteration = Int(Wait_Remaining / Loop_Wait_Time) + Loop_Remainder

If Port_Debug Then
    Port_Name = COM_PORT(Port_Number).Name
    Debug_Text = Wait_Text_1 & Wait_Remaining & Wait_Text_2 & Loop_Iteration
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, Port_Name & Debug_Text)
End If

Do

Clear_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Clear_Result Then

    Queue_Length = COM_PORT(Port_Number).Status.QUEUE_IN
    Data_Waiting = IIf(Queue_Length > LONG_0, True, False)
    
    If Not Data_Waiting Then
    
        Wait_Expired = IIf(Wait_Remaining < LONG_1, True, False)
        
        If Not Wait_Expired Then
    
            If Port_Debug Then
                Debug_Text = Wait_Text_3 & Loop_Iteration & TEXT_COMMA & Wait_Text_4 & Wait_Remaining & TEXT_MS
                Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, Port_Name & Debug_Text)
            End If
            
            Sleep_Time = IIf(Wait_Remaining < Loop_Wait_Time, Wait_Remaining, Loop_Wait_Time)
            
            Kernel_Sleep_MilliSeconds Sleep_Time
            Loop_Iteration = Loop_Iteration - LONG_1
            Wait_Remaining = Wait_Remaining - Sleep_Time
        
        Else
      
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMEOUT, Port_Name & Wait_Text_6 & Wait_Remaining & TEXT_MS)
      
        End If
       
    Else

        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RESULT, Port_Name & Wait_Text_5 & Queue_Length)

    End If
    
Else

    Wait_Expired = True
    Data_Waiting = False
    
    If Port_Debug Then
        Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, Port_Name & Wait_Text_7 & Error_Text)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, Port_Name & Wait_Text_8)
    End If
    
End If

DoEvents

Loop Until Data_Waiting Or Wait_Expired Or Not Clear_Result

SYNCHRONOUS_WAIT_COM_PORT = Data_Waiting

End Function

Public Function READ_COM_PORT(Port_Number As Long, Optional Number_Characters As Long) As String
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "READ_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Read_Timeslice_Bytes As Long
Dim Read_Character_Count As Long
Dim Read_Character_String As String, Port_Name As String

Const Temp_Text_1 As String = " Port Settings = "
Const Temp_Text_2 As String = " Characters Requested = "
Const Temp_Text_3 As String = " Synchronous Read, Result = "
Const Temp_Text_4 As String = " Synchronous Read Failed "

If Port_Valid Then

    If Port_Started(Port_Number) Then
        
        Read_Character_String = vbNullString
        Read_Character_Count = Number_Characters
        Read_Timeslice_Bytes = COM_PORT(Port_Number).Timers.Timeslice_Bytes
    
        If Number_Characters < LONG_1 Or Number_Characters > Read_Timeslice_Bytes Then Read_Character_Count = Read_Timeslice_Bytes
    
        If Port_Debug Then
        
            Port_Name = COM_PORT(Port_Number).Name
            Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SETTINGS, Port_Name & Temp_Text_1 & COM_PORT(Port_Number).Settings)
            Call PRINT_DEBUG_TEXT(Module_Name, TEXT_STARTING, Port_Name & Temp_Text_2 & Read_Character_Count)
            
        End If
        
        Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, Read_Character_Count)
            
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READING, Port_Name & Temp_Text_3 & Temp_Result)
            
        If Temp_Result Then
            
            If Not COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Then Read_Character_String = COM_PORT(Port_Number).Buffers.Read_Result
                   
        Else
            
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, Port_Name & Temp_Text_4)
    
        End If
    
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If
      
Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If
    
READ_COM_PORT = Read_Character_String

End Function

Public Function RECEIVE_COM_PORT(Port_Number As Long) As String
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "RECEIVE_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Receive_MicroSeconds As Currency
Dim Receive_Byte_Count As Long, Bytes_Per_Second As Long

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
Const Temp_Text_11 As String = " Last Data MicroSeconds = "
Const Temp_Text_12 As String = " Read Wait MicroSeconds = "
Const Temp_Text_13 As String = " Receive   MicroSeconds = "
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
            Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, COM_PORT(Port_Number).Timers.Timeslice_Bytes)
            
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_READING, COM_PORT(Port_Number).Name & Temp_Text_03 & Temp_Result)
            
            If Temp_Result Then
            
            If Not COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Then
            
                COM_PORT(Port_Number).Buffers.Receive_Result = COM_PORT(Port_Number).Buffers.Receive_Result & COM_PORT(Port_Number).Buffers.Read_Result
                
                Select Case COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read
                
                Case Is < LONG_4                                         ' assume manual data entry, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SINGLE, COM_PORT(Port_Number).Name & Temp_Text_06 & COM_PORT(Port_Number).Timers.Char_Loop_Wait & TEXT_MS)
                Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Char_Loop_Wait
            
                Case Is < LONG_21                                        ' assume continuous data ending, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text_07 & COM_PORT(Port_Number).Timers.Data_Loop_Wait & TEXT_MS)
                Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Data_Loop_Wait
                
                Case Is = COM_PORT(Port_Number).Timers.Timeslice_Bytes   ' assume more data available immediately, improve responsiveness.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_09)
            
                Case Else                                                ' assume more data from continuous source, allow buffer to refill.
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WAITING, COM_PORT(Port_Number).Name & Temp_Text_08 & COM_PORT(Port_Number).Timers.Line_Loop_Wait & TEXT_MS)
                Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Line_Loop_Wait
                
                End Select
                                            
                DoEvents
                
            End If
                    
            Else
            
                If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_15)
    
            End If
                        
        Loop Until COM_PORT(Port_Number).Buffers.Read_Buffer_Empty Or Not Temp_Result
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_11 & COM_PORT(Port_Number).Timers.Last_Data_Time)
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_LOOPING, COM_PORT(Port_Number).Name & Temp_Text_12 & COM_PORT(Port_Number).Timers.Read_Wait_Time)
        
        If Not COM_PORT(Port_Number).Timers.Read_Timeout Then
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_NO_DATA, COM_PORT(Port_Number).Name & Temp_Text_05 & COM_PORT(Port_Number).Timers.Exit_Loop_Wait & TEXT_MS)
        Kernel_Sleep_MilliSeconds COM_PORT(Port_Number).Timers.Exit_Loop_Wait
        
        End If
                
     Loop Until COM_PORT(Port_Number).Timers.Read_Timeout Or Not Temp_Result
     
    If Port_Debug Then
    
        PORT_MICROSECONDS_END Port_Number
        Receive_MicroSeconds = PORT_MICROSECONDS(Port_Number)
        Receive_Byte_Count = Len(COM_PORT(Port_Number).Buffers.Receive_Result)
        Bytes_Per_Second = Receive_Byte_Count / (Receive_MicroSeconds / LONG_1E6)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_DURATION, COM_PORT(Port_Number).Name & Temp_Text_13 & Receive_MicroSeconds)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RECEIVED, COM_PORT(Port_Number).Name & Temp_Text_14 & Receive_Byte_Count)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FINISHED, COM_PORT(Port_Number).Name & Temp_Text_10 & Bytes_Per_Second)
    
    End If
    
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If
      
Else

   Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If
    
RECEIVE_COM_PORT = COM_PORT(Port_Number).Buffers.Receive_Result

End Function

Public Function TRANSMIT_COM_PORT(Port_Number As Long, Transmit_Text As String) As Boolean
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "TRANSMIT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Debug_String As String
Dim Transmit_Time As Currency
Dim Loop_Closing As Boolean, Temp_Result As Boolean
Dim Temp_Pointer As Long, Transmit_Length As Long, Bytes_Per_Second As Long
Dim Byte_Pointer As Long, Timeslice_Bytes As Long, Byte_Count As Long, Loop_Counter As Long

Const Temp_Text_1 As String = " Port Settings = "
Const Temp_Text_2 As String = " Timeslice Bytes/Second = "
Const Temp_Text_3 As String = " Transmitting Bytes "
Const Temp_Text_4 As String = " Transmit Time for "
Const Temp_Text_5 As String = " Bytes = "
Const Temp_Text_6 As String = " Effective Bytes per Second = "
Const Temp_Text_7 As String = " ("
Const Temp_Text_8 As String = " Bytes)"

If Port_Valid Then

If Port_Started(Port_Number) Then

Transmit_Length = Len(Transmit_Text)
Timeslice_Bytes = COM_PORT(Port_Number).Timers.Timeslice_Bytes

If Port_Debug Then

    Debug_String = COM_PORT(Port_Number).Name & Temp_Text_1 & COM_PORT(Port_Number).Settings
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SETTINGS, Debug_String)
    
    Debug_String = COM_PORT(Port_Number).Name & Temp_Text_2 & Timeslice_Bytes
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, Debug_String)

    PORT_MICROSECONDS_NOW Port_Number

End If

For Loop_Counter = LONG_1 To Transmit_Length Step Timeslice_Bytes

    Byte_Pointer = (Loop_Counter + Timeslice_Bytes) - LONG_1
    Loop_Closing = IIf(Transmit_Length - Loop_Counter < Timeslice_Bytes, True, False)
    Temp_Pointer = IIf(Loop_Closing, Transmit_Length, Byte_Pointer)
    Byte_Count = Temp_Pointer - Loop_Counter + LONG_1
    
    If Port_Debug Then
    
    Debug_String = Temp_Text_3 & Loop_Counter & TEXT_TO & Temp_Pointer & Temp_Text_7 & Byte_Count & Temp_Text_8
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_WRITING, COM_PORT(Port_Number).Name & Debug_String)
    
    End If
    
    COM_PORT(Port_Number).Buffers.Write_Buffer = Mid$(Transmit_Text, Loop_Counter, Timeslice_Bytes)

    Temp_Result = SYNCHRONOUS_WRITE_COM_PORT(Port_Number)

    DoEvents

Next Loop_Counter

If Port_Debug Then
    
    PORT_MICROSECONDS_END Port_Number
    
    Transmit_Time = PORT_MICROSECONDS(Port_Number)
    Bytes_Per_Second = Transmit_Length / Transmit_Time * LONG_1E6
    Debug_String = Temp_Text_4 & Transmit_Length & Temp_Text_5 & Int(Transmit_Time / LONG_1000) & TEXT_MS

    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, COM_PORT(Port_Number).Name & Debug_String)
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_TIMING, COM_PORT(Port_Number).Name & Temp_Text_6 & Bytes_Per_Second)

End If

Else
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
End If

Else
        Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
End If

DoEvents

TRANSMIT_COM_PORT = Temp_Result

End Function

Private Function GET_PORT_CONFIG(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_CONFIG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getcommstate
' get existing com port config and write to port's device control block

Dim Temp_Result As Boolean
Dim Result_Text As String, Detail_Text As String

Const Temp_Text_1 As String = " Existing Port Settings, "
Const Temp_Text_2 As String = " Port not started, no data available, "

Temp_Result = Query_Port_DCB(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

    Result_Text = TEXT_SUCCESS
    Detail_Text = Temp_Text_1 & GET_PORT_SETTINGS_FROM_DCB(Port_Number)

Else

    Result_Text = TEXT_FAILURE
    Detail_Text = Temp_Text_2 & DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CONFIG, COM_PORT(Port_Number).Name & Detail_Text)


GET_PORT_CONFIG = Temp_Result

End Function

Public Function GET_PORT_SETTINGS(Port_Number As Long) As String
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_SETTINGS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Port_Settings As String

Const TEXT_PORT_INVALID As String = "INVALID-PORT"
Const TEXT_NOT_STARTED As String = "PORT-NOT-STARTED"

If Port_Valid Then

    If Port_Started(Port_Number) Then

        Port_Settings = vbNullString
        Port_Settings = Port_Settings & COM_PORT(Port_Number).DCB.BAUD_RATE & TEXT_DASH
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
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_READ"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Error_Text As String
Dim Temp_Result As Boolean

Const TEXT_SYNC_READ As String = "SYNC_READ"

Const Temp_Text_1 As String = " Synchronous Read, Bytes = "
Const Temp_Text_2 As String = " Read FAILED, "

Temp_Result = Synchronous_Read(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Read_Buffer, Read_Bytes_Requested, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SYNC_READ, COM_PORT(Port_Number).Name & Temp_Text_1 & COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Read)

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

    Temp_Result = False
    COM_PORT(Port_Number).Timers.Read_Timeout = True
    COM_PORT(Port_Number).Buffers.Read_Buffer_Empty = True
    COM_PORT(Port_Number).Buffers.Read_Result = vbNullString
    Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
    
    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SYNC_READ, COM_PORT(Port_Number).Name & Temp_Text_2 & Error_Text)
    
End If

DoEvents

SYNCHRONOUS_READ_COM_PORT = Temp_Result

End Function

Private Function SYNCHRONOUS_WRITE_COM_PORT(Port_Number As Long) As Boolean
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SYNCHRONOUS_WRITE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Error_Text As String
Dim Write_Buffer_Length As Long
Dim Write_Complete As Boolean, Temp_Result As Boolean

Const TEXT_SYNC_WRITE As String = "SYNC_WRITE"

Const Temp_Text_1 As String = " Synchronous Write, Write Length  = "
Const Temp_Text_2 As String = " Synchronous Write, Bytes Written = "
Const Temp_Text_3 As String = " Write FAILED, "

Write_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Write_Buffer)

Temp_Result = Synchronous_Write(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Buffers.Write_Buffer, Write_Buffer_Length, COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent = Write_Buffer_Length Then

Write_Complete = True
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SYNC_WRITE, COM_PORT(Port_Number).Name & Temp_Text_1 & Write_Buffer_Length)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SYNC_WRITE, COM_PORT(Port_Number).Name & Temp_Text_2 & COM_PORT(Port_Number).Buffers.Synchronous_Bytes_Sent)

Else

Write_Complete = False
Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SYNC_WRITE, COM_PORT(Port_Number).Name & Temp_Text_3 & Error_Text)

End If

DoEvents

SYNCHRONOUS_WRITE_COM_PORT = Write_Complete

End Function

Public Function SEND_COM_PORT(Port_Number As Long, Send_Variable As Variant) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SEND_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/return-values-for-the-cstr-function

Dim Send_Result As Boolean
Const Temp_Text As String = " Transmit Result = "

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
        Send_Result = TRANSMIT_COM_PORT(Port_Number, CStr(Send_Variable))
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_RESULT, COM_PORT(Port_Number).Name & Temp_Text & Send_Result)

    Else
    
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
    End If

Else

   Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SEND_COM_PORT = Send_Result

End Function

Public Function PUT_COM_PORT(Port_Number As Long, Put_String As String) As Boolean
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "PUT_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Write_Result As Boolean
Dim Write_Byte_Count As Long
Dim Error_Text As String, Put_Character As String

If Port_Valid Then

    If Port_Started(Port_Number) Then

        Put_Character = Left$(Put_String, LONG_1)
        
        Write_Result = Synchronous_Write(COM_PORT(Port_Number).Handle, Put_Character, LONG_1, Write_Byte_Count)
      
    Else
    
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
        
    End If

Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
        
End If

PUT_COM_PORT = Write_Result

End Function

Public Function GET_COM_PORT(Port_Number As Long) As String
'-----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-----------------------------------------------------------------------

Dim Read_Byte_Count As Long
Dim Get_Character As String * LONG_1               ' must be fixed length 1

If Port_Valid Then

    If Port_Started(Port_Number) Then

        Synchronous_Read COM_PORT(Port_Number).Handle, Get_Character, LONG_1, Read_Byte_Count
      
    Else

        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
        
    End If
        
Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
        
End If

GET_COM_PORT = Get_Character

End Function

Private Function PURGE_COM_PORT(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "PURGE_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Result_Text As String, Detail_Text As String, Error_Text As String

Const Purge_Text_1 As String = "PURGING"
Const Purge_Text_2 As String = " Purge Com Port Success, Result = "
Const Purge_Text_3 As String = " Purge Com Port Failed, "

Temp_Result = Com_Port_Purge(COM_PORT(Port_Number).Handle, PORT_CONTROL.PURGE_ALL)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError
Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

If Temp_Result Then

        Result_Text = TEXT_SUCCESS
        Detail_Text = Purge_Text_2 & Temp_Result

Else

        Result_Text = TEXT_FAILURE
        Detail_Text = Purge_Text_3 & Error_Text
    
End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

DoEvents

PURGE_COM_PORT = Temp_Result

End Function

Private Sub PURGE_BUFFERS(Port_Number As Long)
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "PURGE_BUFFERS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

Const Result_Text As String = "BUFFERS"
Const Detail_Text As String = " Purged Port Profile Read and Write Data Buffers "

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

    If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

End Sub

Private Function SET_PORT_TIMERS(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_TIMERS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Result_Text As String, Timer_Text As String
Dim Detail_Text As String, Error_Text As String

Const NO_TIMEOUT As Long = MAXDWORD
Const WRITE_CONSTANT As Long = LONG_3000
Const TEXT_READ_EQUALS As String = " Read = "
Const TEXT_WRITE_EQUALS As String = TEXT_COMMA & " Write = "

Const Set_Text_1 As String = " Port Read/Write Timers Applied,"
Const Set_Text_2 As String = " Port Timers Not Set, "

COM_PORT(Port_Number).Timeouts.Read_Interval_Timeout = NO_TIMEOUT              ' Timeouts not used for file reads.
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Constant = LONG_0            '
COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT
COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Multiplier = LONG_0

Temp_Result = Set_Com_Timers(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Timeouts)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

Result_Text = TEXT_SUCCESS
Timer_Text = TEXT_READ_EQUALS & NO_TIMEOUT & TEXT_MS & TEXT_WRITE_EQUALS & WRITE_CONSTANT & TEXT_MS
Detail_Text = Set_Text_1 & Timer_Text

Else

Result_Text = TEXT_FAILURE
Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
Detail_Text = Set_Text_2 & Error_Text

End If

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)

SET_PORT_TIMERS = Temp_Result

End Function

Public Function SHOW_PORT_TIMERS(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_TIMERS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Port_Name As String
Dim Error_Text As String
Dim Temp_Result As Boolean
Dim Temp_Timer(LONG_1 To LONG_5) As String

Const Print_Text_1 As String = "COM_PORT_TIMERS"
Const Print_Text_2 As String = "TIMER READ"
Const Print_Text_3 As String = "TIMER WRITE"
Const Print_Text_4 As String = " Error retrieving Timer Settings for "

Const Timer_Text_1 As String = " Read Interval "
Const Timer_Text_2 As String = " Read Constant "
Const Timer_Text_3 As String = " Read Multiplier "
Const Timer_Text_4 As String = " Write Constant "
Const Timer_Text_5 As String = " Write Multiplier "

If Port_Valid Then

    If Port_Started(Port_Number) Then

        Port_Name = COM_PORT(Port_Number).Name

        Temp_Result = Get_Com_Timers(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Timeouts)
        
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError
       
        If Temp_Result Then

        Temp_Timer(LONG_1) = COM_PORT(Port_Number).Timeouts.Read_Interval_Timeout
        Temp_Timer(LONG_2) = COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Constant
        Temp_Timer(LONG_3) = COM_PORT(Port_Number).Timeouts.Read_Total_Timeout_Multiplier
        Temp_Timer(LONG_4) = COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Constant
        Temp_Timer(LONG_5) = COM_PORT(Port_Number).Timeouts.Write_Total_Timeout_Multiplier

        Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Timer_Text_1, TEXT_EQUALS_SPACE & Temp_Timer(LONG_1))
        Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Timer_Text_2, TEXT_EQUALS_SPACE & Temp_Timer(LONG_2))
        Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Timer_Text_3, TEXT_EQUALS_SPACE & Temp_Timer(LONG_3))
        Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_3, Port_Name & Timer_Text_4, TEXT_EQUALS_SPACE & Temp_Timer(LONG_4))
        Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_3, Port_Name & Timer_Text_5, TEXT_EQUALS_SPACE & Temp_Timer(LONG_5))

        Else
        
        Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
        Call PRINT_DEBUG_TEXT(Module_Name, TEXT_ERROR, Print_Text_4 & Port_Name & TEXT_COMMA & Error_Text)

        End If

    Else

        Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

    End If

Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_TIMERS = Temp_Result

End Function

Public Function SHOW_PORT_QUEUES(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_QUEUES"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

Dim Temp_Result As Boolean

Const Temp_Text_1 As String = "COM PORT QUEUE"
Const Temp_Text_2 As String = "QUEUE_IN  "
Const Temp_Text_3 As String = "QUEUE_OUT "
Const Temp_Text_4 As String = "= NO QUEUE DATA "
Const Temp_Text_5 As String = " Input  Queue "
Const Temp_Text_6 As String = " Output Queue "
Const Temp_Text_7 As String = " Clear Comms Error Failed "

If Port_Valid Then

    If Port_Started(Port_Number) Then

    If CLEAR_PORT_ERROR(Port_Number) Then

    Debug.Print
    
        Temp_Result = True
        Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_2, COM_PORT(Port_Number).Name & Temp_Text_5, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Status.QUEUE_IN)
        Call PRINT_SHOW_TEXT(Temp_Text_1, Temp_Text_3, COM_PORT(Port_Number).Name & Temp_Text_6, TEXT_EQUALS_SPACE & COM_PORT(Port_Number).Status.QUEUE_OUT)

    Else
        
        Call PRINT_SHOW_TEXT(Module_Name, TEXT_FAILED, COM_PORT(Port_Number).Name & Temp_Text_7, Temp_Text_4)
    
    End If
    
    Else
    
        Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    
    End If
     
Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_QUEUES = Temp_Result

End Function

Public Function CHECK_COM_PORT(Port_Number As Long) As Long
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CHECK_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_Result As Boolean, Temp_Queue As Long, Error_Text As String

Const Check_Text_1 As String = " Receive characters waiting to be read = "
Const Check_Text_2 As String = " Clear Comms Error Failed, Queue Data not available"
Const Check_Text_3 As String = " Clear Comms Error Failed, "

Temp_Queue = LONG_NEG_1

If Port_Valid Then

If Port_Started(Port_Number) Then
    
        Temp_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)
        
        COM_PORT(Port_Number).DLL_Error = Err.LastDllError
        
        If Temp_Result Then
    
            Temp_Queue = COM_PORT(Port_Number).Status.QUEUE_IN
        
            If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SUCCESS, COM_PORT(Port_Number).Name & Check_Text_1 & Temp_Queue)

        Else
        
            If Port_Debug Then
            
                Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
                Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & Check_Text_2)
                Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & Check_Text_3 & Error_Text)
                
            End If
            
        End If
        
Else

   If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If
        

Else

   Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

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
Dim Port_Name As String
Dim Error_Text As String
Dim Temp_Result As Boolean
Dim Temp_Error(LONG_1 To LONG_5) As String

Const Prefix_Text_1 As String = "OVERFLOW"
Const Prefix_Text_2 As String = "OVERRUN"
Const Prefix_Text_3 As String = "PARITY"
Const Prefix_Text_4 As String = "FRAMING"
Const Prefix_Text_5 As String = "BREAK"

Const Error_Text_1 As String = " Input Buffer Overflow "
Const Error_Text_2 As String = " Character Buffer Over-Run "
Const Error_Text_3 As String = " Hardware Parity Error "
Const Error_Text_4 As String = " Hardware Framing Error "
Const Error_Text_5 As String = " Hardware Break Signal "

Const TEXT_PORT_ERRORS As String = "COM_PORT_ERRORS"

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError

Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

Port_Name = COM_PORT(Port_Number).Name
Port_Error = COM_PORT(Port_Number).Port_Errors

Temp_Error(LONG_1) = IIf(Port_Error And Port_Errors.OVERFLOW, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_2) = IIf(Port_Error And Port_Errors.OVERRUN, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_3) = IIf(Port_Error And Port_Errors.PARITY, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_4) = IIf(Port_Error And Port_Errors.FRAME, TEXT_TRUE, TEXT_FALSE)
Temp_Error(LONG_5) = IIf(Port_Error And Port_Errors.BREAK, TEXT_TRUE, TEXT_FALSE)

Call PRINT_SHOW_TEXT(TEXT_PORT_ERRORS, Prefix_Text_1, Port_Name & Error_Text_1, TEXT_EQUALS_SPACE & Temp_Error(LONG_1))
Call PRINT_SHOW_TEXT(TEXT_PORT_ERRORS, Prefix_Text_2, Port_Name & Error_Text_2, TEXT_EQUALS_SPACE & Temp_Error(LONG_2))
Call PRINT_SHOW_TEXT(TEXT_PORT_ERRORS, Prefix_Text_3, Port_Name & Error_Text_3, TEXT_EQUALS_SPACE & Temp_Error(LONG_3))
Call PRINT_SHOW_TEXT(TEXT_PORT_ERRORS, Prefix_Text_4, Port_Name & Error_Text_4, TEXT_EQUALS_SPACE & Temp_Error(LONG_4))
Call PRINT_SHOW_TEXT(TEXT_PORT_ERRORS, Prefix_Text_5, Port_Name & Error_Text_5, TEXT_EQUALS_SPACE & Temp_Error(LONG_5))

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_ERRORS = Temp_Result

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
Dim Port_Signals As Long
Dim Temp_Result As Boolean
Dim SIGNAL_CTS As String, SIGNAL_DSR As String, SIGNAL_RNG As String, SIGNAL_RLS As String

Const Modem_Text_1 As String = "COM_PORT_MODEM"
Const Modem_Text_2 As String = "MODEM (In)  "
Const Modem_Text_3 As String = " Clear to Send (from modem/device) CTS  "
Const Modem_Text_4 As String = " Data Set (modem/device) Ready     DSR  "
Const Modem_Text_5 As String = " Ring Signal (from modem)          RING "
Const Modem_Text_6 As String = " Receive Line Signal Detect        RLSD "
Const Modem_Text_7 As String = " Error retrieving Modem Status, "

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

Port_Signals = COM_PORT(Port_Number).Port_Signals

SIGNAL_CTS = IIf(Port_Signals And PORT_CONTROL.CTS_ON, TEXT_ON, TEXT_OFF)
SIGNAL_DSR = IIf(Port_Signals And PORT_CONTROL.DSR_ON, TEXT_ON, TEXT_OFF)
SIGNAL_RNG = IIf(Port_Signals And PORT_CONTROL.RING_ON, TEXT_ON, TEXT_OFF)
SIGNAL_RLS = IIf(Port_Signals And PORT_CONTROL.RLSD_ON, TEXT_ON, TEXT_OFF)

Call PRINT_SHOW_TEXT(Modem_Text_1, Modem_Text_2, COM_PORT(Port_Number).Name & Modem_Text_3, TEXT_EQUALS_SPACE & SIGNAL_CTS)
Call PRINT_SHOW_TEXT(Modem_Text_1, Modem_Text_2, COM_PORT(Port_Number).Name & Modem_Text_4, TEXT_EQUALS_SPACE & SIGNAL_DSR)
Call PRINT_SHOW_TEXT(Modem_Text_1, Modem_Text_2, COM_PORT(Port_Number).Name & Modem_Text_5, TEXT_EQUALS_SPACE & SIGNAL_RNG)
'Call PRINT_SHOW_TEXT(Modem_Text_1, Modem_Text_2, COM_PORT(Port_Number).Name & Modem_Text_6, TEXT_EQUALS_SPACE & SIGNAL_RLS)

Else

Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

Call PRINT_DEBUG_TEXT(Module_Name, TEXT_ERROR, COM_PORT(Port_Number).Name & Modem_Text_7 & Error_Text)

End If

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_MODEM = Temp_Result

End Function

Private Function CLEAR_PORT_ERROR(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CLEAR_COM_ERROR"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Result_Text As String, Detail_Text As String, Error_Text As String

Const Clear_Text_1 As String = " Attempting to Clear Comms Error(s)"
Const Clear_Text_2 As String = " Comms Error(s) Cleared Successfully "

If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, TEXT_CLEARING, COM_PORT(Port_Number).Name & Clear_Text_1)

Temp_Result = Com_Port_Clear(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Errors, COM_PORT(Port_Number).Status)

COM_PORT(Port_Number).DLL_Error = Err.LastDllError
Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

If Temp_Result Then

    Result_Text = TEXT_CLEARED
    Detail_Text = Clear_Text_2

Else

    Result_Text = TEXT_FAILED
    Detail_Text = TEXT_SPACE & Error_Text

End If

Call PRINT_SHOW_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text, TEXT_EQUALS_SPACE & Result_Text)

CLEAR_PORT_ERROR = Temp_Result

End Function

Public Static Function SHOW_PORT_DCB(Port_Number As Long) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_DCB"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

Dim Temp_Result As Boolean
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

If Port_Valid Then

If Port_Started(Port_Number) Then

Temp_Result = True

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
DCB_TEXT(25) = " Reserved Word 0              ": DCB_VALUE(25) = COM_PORT(Port_Number).DCB.RESERVED_0
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
DCB_TEXT(36) = " Reserved Word 1              ": DCB_VALUE(36) = COM_PORT(Port_Number).DCB.RESERVED_1

For DCB_COUNTER = LONG_10 To LONG_36
Call PRINT_SHOW_TEXT(Module_Name, Temp_Device, COM_PORT(Port_Number).Name & DCB_TEXT(DCB_COUNTER), TEXT_EQUALS_SPACE & DCB_VALUE(DCB_COUNTER))
Next DCB_COUNTER

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_DCB = Temp_Result

End Function

Public Static Function SHOW_PORT_STATUS(Port_Number As Long) As Boolean
'--------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SHOW_PORT_STATUS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-comstat

Dim Port_Name As String
Dim Temp_Wait(LONG_1 To LONG_9) As String
Dim Temp_Result As Boolean, Temp_Bitmap As Long

Const Print_Text_1 As String = "COM_PORT_STATUS"
Const Print_Text_2 As String = "FLOW CONTROL"
Const Print_Text_3 As String = "END OF FILE "
Const Print_Text_4 As String = "RX INTERRUPT"
Const Print_Text_5 As String = "QUEUE LENGTH"

Const Failed_Text_1 As String = " Failed to Clear Com Errors       "
Const Status_Text_1 As String = " Transmission Waiting for CTS     "
Const Status_Text_2 As String = " Transmission Waiting for DSR     "
Const Status_Text_3 As String = " Transmission Waiting for RLSD    "
Const Status_Text_4 As String = " Transmission Waiting (XOFF Hold) "
Const Status_Text_5 As String = " Transmission Waiting (XOFF Sent) "
Const Status_Text_6 As String = " End Of File (EOF) Received       "
Const Status_Text_7 As String = " Priority / Interrupt Character   "
Const Status_Text_8 As String = " Input Queue (awaiting ReadFile)  "
Const Status_Text_9 As String = " Output Queue (for Transmission)  "

If Port_Valid Then

If Port_Started(Port_Number) Then

Port_Name = COM_PORT(Port_Number).Name

Temp_Result = CLEAR_PORT_ERROR(Port_Number)

If Temp_Result Then

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

Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Status_Text_1, TEXT_EQUALS_SPACE & Temp_Wait(LONG_1))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Status_Text_2, TEXT_EQUALS_SPACE & Temp_Wait(LONG_2))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Status_Text_3, TEXT_EQUALS_SPACE & Temp_Wait(LONG_3))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Status_Text_4, TEXT_EQUALS_SPACE & Temp_Wait(LONG_4))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_2, Port_Name & Status_Text_5, TEXT_EQUALS_SPACE & Temp_Wait(LONG_5))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_3, Port_Name & Status_Text_6, TEXT_EQUALS_SPACE & Temp_Wait(LONG_6))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_4, Port_Name & Status_Text_7, TEXT_EQUALS_SPACE & Temp_Wait(LONG_7))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Status_Text_8, TEXT_EQUALS_SPACE & Temp_Wait(LONG_8))
Call PRINT_SHOW_TEXT(Print_Text_1, Print_Text_5, Port_Name & Status_Text_9, TEXT_EQUALS_SPACE & Temp_Wait(LONG_9))

Else

Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, Port_Name & Failed_Text_1)

End If

Else

Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If

Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

SHOW_PORT_STATUS = Temp_Result

End Function

Public Function DEVICE_READY(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "DEVICE_READY"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' returns True if port valid, started and COM Port DSR signal is asserted.
' DSR = Data Set Ready,from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Error_Text As String, Debug_Text As String
Dim Temp_Result As Boolean, Signal_State As Boolean

Const Status_Text As String = "DSR_STATE"
Const Device_State As String = " Data Set Ready, State = "
Const Error_Prefix As String = " Error Retreiving Modem Status, "

If Port_Valid Then

    If Port_Started(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Temp_Result Then

        Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.DSR_ON, True, False)
        If Port_Debug Then
            Debug_Text = COM_PORT(Port_Number).Name & Device_State & Signal_State
            Call PRINT_DEBUG_TEXT(Module_Name, Status_Text, Debug_Text)
        End If
    
    Else
    
        Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
        Debug_Text = COM_PORT(Port_Number).Name & Error_Prefix & Error_Text
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Status_Text, Debug_Text)

    End If

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DEVICE_READY = Signal_State

End Function

Public Function DEVICE_CALLING(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "RING_INDICATE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' returns True if port valid, started and COM Port RI signal is asserted.
' Ring Indicator, from attached modem, serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Error_Text As String, Debug_Text As String
Dim Temp_Result As Boolean, Signal_State As Boolean

Const Status_Text As String = "RING_STATE"
Const Device_State As String = " Ring Indicator, State = "
Const Error_Prefix As String = " Error Retreiving Modem Status, "

If Port_Valid Then

    If Port_Started(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Temp_Result Then

        Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.RING_ON, True, False)
        If Port_Debug Then
            Debug_Text = COM_PORT(Port_Number).Name & Device_State & Signal_State
            Call PRINT_DEBUG_TEXT(Module_Name, Status_Text, Debug_Text)
        End If
        
    Else
    
        Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
        Debug_Text = COM_PORT(Port_Number).Name & Error_Prefix & Error_Text
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Status_Text, Debug_Text)

    End If

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

DEVICE_CALLING = Signal_State

End Function

Public Function CLEAR_TO_SEND(Port_Number As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "CLEAR_TO_SEND"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' returns True if port valid, started and COM Port CTS signal is asserted.
' CTS = Clear To Send, from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Error_Text As String
Dim Temp_Result As Boolean, Signal_State As Boolean

Const Status_Text As String = "MODEM_STATUS"
Const Error_Prefix As String = "Error Retreiving Modem Status, "

If Port_Valid Then

    If Port_Started(Port_Number) Then

    Temp_Result = Get_Port_Modem(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).Port_Signals)
    
    COM_PORT(Port_Number).DLL_Error = Err.LastDllError

    If Temp_Result Then

        Signal_State = IIf(COM_PORT(Port_Number).Port_Signals And PORT_CONTROL.CTS_ON, True, False)

    Else
    
        Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)

        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Status_Text, COM_PORT(Port_Number).Name & Error_Prefix & Error_Text)

    End If

Else

If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

CLEAR_TO_SEND = Signal_State

End Function

Public Function SIGNAL_COM_PORT(Port_Number As Long, Signal_Function As Long) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SIGNAL_COM_PORT"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction

Dim Signal_Result As Boolean
Dim Signal_Valid As Boolean
Dim Result_Text As String, Detail_Text As String, Error_Text As String

Const Signal_Text_1 As String = "SIGNALLING"
Const Signal_Text_2 As String = " Signal Com Port Success, Result = "
Const Signal_Text_3 As String = " Signal Function Invalid, "
Const Signal_Text_4 As String = " Signal Com Port Failed, "

Signal_Valid = IIf(Signal_Function < LONG_10 And Signal_Function > LONG_0, True, False)

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
        If Signal_Valid Then
        
            Signal_Result = Set_Com_Signal(COM_PORT(Port_Number).Handle, Signal_Function)
            
            COM_PORT(Port_Number).DLL_Error = Err.LastDllError
            Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
                
            If Signal_Result Then
                                    Result_Text = TEXT_SUCCESS
                                    Detail_Text = Signal_Text_2 & Signal_Result
            Else
                                    Result_Text = TEXT_FAILURE
                                    Detail_Text = Signal_Text_4 & Error_Text
            End If
                 
        Else
                Result_Text = TEXT_FAILURE
                Detail_Text = Signal_Text_3 & Signal_Function
        End If
        
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)
        
    Else
            If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
    End If

Else
        Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
End If

SIGNAL_COM_PORT = Signal_Result

End Function

Public Function REQUEST_TO_SEND_COM_PORT(Port_Number As Long, RTS_State As Boolean) As Boolean
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "REQUEST_TO_SEND"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction

Dim RTS_Signal As Long
Dim RTS_String As String
Dim RTS_Result As Boolean
Dim Result_Text As String, Detail_Text As String, Debug_Text As String, Error_Text As String

Const SIGNAL_RTS_1 As Long = LONG_3
Const SIGNAL_RTS_0 As Long = LONG_4

Const TEXT_RTS_1 As String = " RTS = On "
Const TEXT_RTS_0 As String = " RTS = Off "

Const Signal_Text_1 As String = "SIGNALLING"
Const Signal_Text_2 As String = " Sending Port RTS Function "
Const Signal_Text_3 As String = " Signal Com Port RTS Success,"
Const Signal_Text_4 As String = " RTS Failed, "

RTS_String = IIf(RTS_State, TEXT_RTS_1, TEXT_RTS_0)
RTS_Signal = IIf(RTS_State, SIGNAL_RTS_1, SIGNAL_RTS_0)

If Port_Valid Then

    If Port_Started(Port_Number) Then
    
            If Port_Debug Then
                Debug_Text = COM_PORT(Port_Number).Name & Signal_Text_2 & RTS_Signal & TEXT_COMMA & RTS_String
                PRINT_DEBUG_TEXT Module_Name, Signal_Text_1, Debug_Text
            End If
            
            RTS_Result = Set_Com_Signal(COM_PORT(Port_Number).Handle, RTS_Signal)
            
            COM_PORT(Port_Number).DLL_Error = Err.LastDllError
            Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
                
            If RTS_Result Then
            
                Result_Text = TEXT_SUCCESS
                Detail_Text = Signal_Text_3 & RTS_String
                Kernel_Sleep_MilliSeconds LONG_50                ' optional - allow local and remote hardware devices to settle.
            
            Else
            
                Result_Text = TEXT_FAILURE
                Detail_Text = Signal_Text_4 & Error_Text
            
            End If
                    
        If Port_Debug Then Call PRINT_DEBUG_TEXT(Module_Name, Result_Text, COM_PORT(Port_Number).Name & Detail_Text)
        
    Else
    
        If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)
        
    End If

Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
        
End If

REQUEST_TO_SEND_COM_PORT = RTS_Result

End Function

Private Function GET_PORT_SETTINGS_FROM_DCB(Port_Number As Long) As String
'---------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_PORT_SETTINGS"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'---------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Temp_String As String, Error_Text As String

Const TEXT_BAUD_EQUALS   As String = "BAUD="
Const TEXT_DATA_EQUALS   As String = "DATA="
Const TEXT_STOP_EQUALS   As String = "STOP="
Const TEXT_PARITY_EQUALS As String = "PARITY="
Const TEXT_NO_DCB_DATA   As String = "ERROR-NO-DCB-DATA"
Const TEXT_INVALID_PORT  As String = "ERROR-INVALID-PORT"

Const Error_Prefix As String = " Error in Get Com State for DCB data, "

If Port_Valid Then

Temp_Result = Query_Port_DCB(COM_PORT(Port_Number).Handle, COM_PORT(Port_Number).DCB)
COM_PORT(Port_Number).DLL_Error = Err.LastDllError

If Temp_Result Then

Temp_String = vbNullString
Temp_String = Temp_String & TEXT_BAUD_EQUALS & COM_PORT(Port_Number).DCB.BAUD_RATE & TEXT_SPACE
Temp_String = Temp_String & TEXT_DATA_EQUALS & COM_PORT(Port_Number).DCB.BYTE_SIZE & TEXT_SPACE
Temp_String = Temp_String & TEXT_PARITY_EQUALS & CONVERT_PARITY(COM_PORT(Port_Number).DCB.PARITY) & TEXT_SPACE
Temp_String = Temp_String & TEXT_STOP_EQUALS & CONVERT_STOPBITS(COM_PORT(Port_Number).DCB.STOP_BITS) & TEXT_SPACE
'Temp_String = Temp_String & "X_IN=" & IIf(COM_PORT(Port_Number).DCB.BIT_FIELD And &H200, TEXT_ON, TEXT_OFF) & TEXT_SPACE
'Temp_String = Temp_String & "X_OUT=" & IIf(COM_PORT(Port_Number).DCB.BIT_FIELD And &H100, TEXT_ON, TEXT_OFF) & TEXT_SPACE

Else

Temp_String = TEXT_NO_DCB_DATA
Error_Text = DLL_ERROR_TEXT(COM_PORT(Port_Number).DLL_Error)
Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FAILURE, COM_PORT(Port_Number).Name & Error_Prefix & Error_Text)

End If
   
Else

Temp_String = TEXT_INVALID_PORT

End If

GET_PORT_SETTINGS_FROM_DCB = Temp_String

End Function

Private Function GET_FRAME_TIME(Port_Number As Long) As Single
'-------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "GET_FRAME_TIME"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'-------------------------------------------------------------------------

Dim Frame_Duration As Single
Dim Error_Text As String, Temp_String As String
Dim Frame_Length As Long, Length_Stop As Long, BAUD_RATE As Long
Dim Length_Start As Long, Length_Data As Long, Length_Parity As Long

Const TEXT_BAUD_RATE As String = " Baud Rate="
Const TEXT_FRAME_INFO As String = "FRAME_INFO"
Const TEXT_FRAME_TIME As String = TEXT_COMMA & " Frame Duration="
Const TEXT_FRAME_LENGTH As String = TEXT_COMMA & " Frame Length="

BAUD_RATE = COM_PORT(Port_Number).DCB.BAUD_RATE

Length_Start = LONG_1
Length_Data = COM_PORT(Port_Number).DCB.BYTE_SIZE
Length_Stop = IIf(COM_PORT(Port_Number).DCB.STOP_BITS = LONG_0, LONG_1, LONG_2)
Length_Parity = IIf(COM_PORT(Port_Number).DCB.PARITY = LONG_0, LONG_0, LONG_1)

Frame_Length = Length_Start + Length_Data + Length_Parity + Length_Stop
Frame_Duration = Frame_Length / BAUD_RATE * LONG_1E6   ' frame (character) duration in MicroSeconds

If Port_Debug Then

    Temp_String = TEXT_BAUD_RATE & BAUD_RATE & TEXT_FRAME_LENGTH & Frame_Length & TEXT_FRAME_TIME & Frame_Duration

    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_FRAME_INFO, COM_PORT(Port_Number).Name & Temp_String & TEXT_US)

End If

GET_FRAME_TIME = Frame_Duration

End Function

Public Function DEBUG_COM_PORT(Port_Number As Long, Optional Debug_State As Variant) As Boolean
'------------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "SET_PORT_DEBUG"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'------------------------------------------------------------------------

Dim COM_PORT_STRING As String
Const TEXT_SET_DEBUG  As String = "SET_DEBUG"
Const TEXT_NEW_STATE  As String = TEXT_COMMA & " New Debug State = "

If Port_Valid Then
   
    If IsMissing(Debug_State) Then
    
            Port_Debug = Not COM_PORT(Port_Number).Debug
    Else
            Port_Debug = CBool(Debug_State)
    End If
    
    COM_PORT(Port_Number).Debug = Port_Debug
    COM_PORT_STRING = TEXT_COM_PORT & CStr(Port_Number) & TEXT_NEW_STATE
    Call PRINT_DEBUG_TEXT(Module_Name, TEXT_SET_DEBUG, COM_PORT_STRING & Port_Debug)
    
Else

    Call PRINT_INVALID_TEXT(Module_Name, Port_Number)
    
End If

DEBUG_COM_PORT = Port_Debug

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

Public Function Port_Number_Valid(Port_Number As Long) As Boolean

Port_Number_Valid = IIf((Port_Number < COM_PORT_MIN) Or (Port_Number > COM_PORT_MAX), False, True)

End Function

Private Function Port_Started(Port_Number As Long) As Boolean

Port_Started = IIf(COM_PORT(Port_Number).Handle > LONG_0, True, False)

End Function

Private Function COM_PORT_CLOSED(Port_Number As Long) As Boolean

COM_PORT_CLOSED = IIf(COM_PORT(Port_Number).Handle < LONG_1, True, False)

End Function

Private Sub PORT_MICROSECONDS_NOW(Port_Number As Long)

QPC COM_PORT(Port_Number).Timers.Timing_QPC_Now

End Sub

Private Sub PORT_MICROSECONDS_END(Port_Number As Long)

QPC COM_PORT(Port_Number).Timers.Timing_QPC_End

End Sub

Private Function DELTA_MICROSECONDS(Port_Number As Long) As Currency

With COM_PORT(Port_Number).Timers

DELTA_MICROSECONDS = .Timing_QPC_End - .Timing_QPC_Now

End With

End Function

Private Function PORT_MICROSECONDS(Port_Number As Long) As Currency

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/currency-data-type

PORT_MICROSECONDS = Int(DELTA_MICROSECONDS(Port_Number) * LONG_1000)

End Function

Private Function PORT_MILLISECONDS(Port_Number As Long) As Long

PORT_MILLISECONDS = Int(DELTA_MICROSECONDS(Port_Number))

End Function

Public Function TIMESTAMP() As String

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim TIMESTAMP_TIME As SYSTEMTIME
Dim TIMESTAMP_STRING As String * LONG_14

Get_System_Time TIMESTAMP_TIME

TIMESTAMP_STRING = Time() & TEXT_DOT & TIMESTAMP_TIME.MilliSeconds

TIMESTAMP = TIMESTAMP_STRING

End Function

Public Function GET_HOST_MILLISECONDS() As Long

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/currency-data-type

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MILLISECONDS = Int(Temp_QPC)

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

Public Sub DECODE_PORT_ERRORS(ERROR_DATA As Long)

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-clearcommerror

Const Error_Text_1 As String = "Input Buffer Overflow             = "
Const Error_Text_2 As String = "Character Buffer Over-Run         = "
Const Error_Text_3 As String = "Hardware Parity Error             = "
Const Error_Text_4 As String = "Hardware Framing Error            = "
Const Error_Text_5 As String = "Hardware Break Signal             = "

Debug.Print TIMESTAMP & Error_Text_1 & IIf(ERROR_DATA And PORT_EVENT.RX_80_FULL, True, False)
Debug.Print TIMESTAMP & Error_Text_2 & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & Error_Text_3 & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & Error_Text_4 & IIf(ERROR_DATA And PORT_EVENT.LINE_ERROR, True, False)
Debug.Print TIMESTAMP & Error_Text_5 & IIf(ERROR_DATA And PORT_EVENT.BREAK, True, False)

End Sub

Public Static Function DLL_ERROR_TEXT(Error_Code As Long) As String

Dim Error_Text As String

Const ERROR_NUM_PREFIX As String = "Last DLL Error "
Const ERROR_NUM_SUFFIX As String = " = "

Const ERROR_TEXT_000 As String = "SUCCESS"
Const ERROR_TEXT_001 As String = "INVALID FUNCTION"
Const ERROR_TEXT_002 As String = "PORT NOT FOUND"
Const ERROR_TEXT_003 As String = "PATH NOT FOUND"
Const ERROR_TEXT_004 As String = "TOO MANY OPEN FILES"
Const ERROR_TEXT_005 As String = "ACCESS DENIED"
Const ERROR_TEXT_006 As String = "INVALID HANDLE"
Const ERROR_TEXT_013 As String = "INVALID DATA"
Const ERROR_TEXT_015 As String = "DEVICE NOT READY"
Const ERROR_TEXT_087 As String = "INVALID PARAMETER"
Const ERROR_TEXT_122 As String = "INSUFFICIENT BUFFER"
Const ERROR_TEXT_995 As String = "OPERATION ABORTED"
Const ERROR_TEXT_996 As String = "IO INCOMPLETE"
Const ERROR_TEXT_997 As String = "IO PENDING"
Const ERROR_TEXT_998 As String = "NO ACCESS"

Const ERROR_UNKNOWN As String = "UNKNOWN SYSTEM ERROR CODE "

Select Case Error_Code

Case SYSTEM_ERRORS.SUCCESS:              Error_Text = ERROR_TEXT_000
Case SYSTEM_ERRORS.NO_ACCESS:            Error_Text = ERROR_TEXT_998
Case SYSTEM_ERRORS.IO_PENDING:           Error_Text = ERROR_TEXT_997
Case SYSTEM_ERRORS.IO_INCOMPLETE:        Error_Text = ERROR_TEXT_996
Case SYSTEM_ERRORS.INVALID_DATA:         Error_Text = ERROR_TEXT_013
Case SYSTEM_ERRORS.ACCESS_DENIED:        Error_Text = ERROR_TEXT_005
Case SYSTEM_ERRORS.PATH_NOT_FOUND:       Error_Text = ERROR_TEXT_003
Case SYSTEM_ERRORS.FILE_NOT_FOUND:       Error_Text = ERROR_TEXT_002
Case SYSTEM_ERRORS.DEVICE_NOT_READY:     Error_Text = ERROR_TEXT_015
Case SYSTEM_ERRORS.INVALID_HANDLE:       Error_Text = ERROR_TEXT_006
Case SYSTEM_ERRORS.INVALID_FUNCTION:     Error_Text = ERROR_TEXT_001
Case SYSTEM_ERRORS.INVALID_PARAMETER:    Error_Text = ERROR_TEXT_087
Case SYSTEM_ERRORS.OPERATION_ABORTED:    Error_Text = ERROR_TEXT_995
Case SYSTEM_ERRORS.TOO_MANY_OPEN_FILES:  Error_Text = ERROR_TEXT_004
Case SYSTEM_ERRORS.INSUFFICIENT_BUFFER:  Error_Text = ERROR_TEXT_122

Case Else:                               Error_Text = ERROR_UNKNOWN

End Select

DLL_ERROR_TEXT = ERROR_NUM_PREFIX & CStr(Error_Code) & ERROR_NUM_SUFFIX & Error_Text

End Function

Private Sub PRINT_STOPPED_TEXT(Module_Text As String, Port_Number As Long)

Const Print_Text_1 As String = "PORT_STOPPED"
Const Print_Text_2 As String = "COM Port "
Const Print_Text_3 As String = TEXT_COMMA & " Port Not Started"

Call PRINT_DEBUG_TEXT(Module_Text, Print_Text_1, Print_Text_2 & Port_Number & Print_Text_3)

End Sub

Private Sub PRINT_INVALID_TEXT(Module_Text As String, Port_Number As Long)

Const Print_Text_1 As String = "INVALID_PORT"
Const Print_Text_2 As String = "Port Number "
Const Print_Text_3 As String = " Invalid, Defined Port Range = "

Call PRINT_DEBUG_TEXT(Module_Text, Print_Text_1, Print_Text_2 & Port_Number & Print_Text_3 & COM_PORT_RANGE)

End Sub

Public Sub PRINT_DEBUG_TEXT(Module_Text As String, Result_Text As String, Message_Text As String)

Dim TEXT_COLUMN_1 As String * LONG_18
Dim TEXT_COLUMN_2 As String * LONG_18

TEXT_COLUMN_1 = Module_Text
TEXT_COLUMN_2 = Result_Text

Debug.Print TIMESTAMP & TEXT_COLUMN_1 & TEXT_COLUMN_2 & Message_Text

End Sub

Public Sub PRINT_SHOW_TEXT(Device_Text As String, Prefix_Text As String, Detail_Text As String, Result_Text As Variant)

Dim TEXT_COLUMN_1 As String * LONG_18
Dim TEXT_COLUMN_2 As String * LONG_18
Dim TEXT_COLUMN_3 As String * LONG_52

TEXT_COLUMN_1 = Device_Text
TEXT_COLUMN_2 = Prefix_Text
TEXT_COLUMN_3 = Detail_Text

Debug.Print TIMESTAMP & TEXT_COLUMN_1 & TEXT_COLUMN_2 & TEXT_COLUMN_3 & CStr(Result_Text)

End Sub

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

Public Function CONVERT_LINE_ERROR(LINE_ERROR As Byte) As String

Dim Error_Text As String

Const Error_Text_1 As String = "BREAK"
Const Error_Text_2 As String = "FRAME"
Const Error_Text_3 As String = "OVERFLOW"
Const Error_Text_4 As String = "OVERRUN"
Const Error_Text_5 As String = "PARITY"
Const Error_Text_6 As String = "UNKNOWN"

Select Case LINE_ERROR

Case Port_Errors.BREAK:             Error_Text = Error_Text_1
Case Port_Errors.FRAME:             Error_Text = Error_Text_2
Case Port_Errors.OVERFLOW:          Error_Text = Error_Text_3
Case Port_Errors.OVERRUN:           Error_Text = Error_Text_4
Case Port_Errors.PARITY:            Error_Text = Error_Text_5

Case Else:                          Error_Text = Error_Text_6

End Select

CONVERT_LINE_ERROR = Error_Text

End Function

Public Function TEMPLATE(Port_Number As Long) As String
'----------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "TEMPLATE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'----------------------------------------------------------------------

Dim Temp_String As String

Const Temp_Text_1 As String = " Port Valid and Started @ "

If Port_Valid Then

If Port_Started(Port_Number) Then

' function code starts here, e.g.

    Temp_String = Application.Name & TEXT_COMMA & Temp_Text_1 & Time() & vbCrLf

Else

    If Port_Debug Then Call PRINT_STOPPED_TEXT(Module_Name, Port_Number)

End If


Else

Call PRINT_INVALID_TEXT(Module_Name, Port_Number)

End If

TEMPLATE = Temp_String

End Function

Public Function EXAMPLE(Port_Number As Long) As String
'--------------------------------------------------------------------
Dim Port_Debug As Boolean: Const Module_Name As String = "EXAMPLE"
Dim Port_Valid As Boolean: Port_Valid = Port_Number_Valid(Port_Number)
If Port_Valid Then Port_Debug = COM_PORT(Port_Number).Debug
'--------------------------------------------------------------------

' Example showing how to read data from a theoretical digital voltmeter with a serial port connected to COM Port 1
' To demonstrate, connect a terminal emulator or similar device to the local COM Port 1 on this machine
' From the VBA Immediate Window (Control-G), type ?EXAMPLE(1) and wait for MEASURE VOLTAGE to appear on the emulator
' When it appears, respond immediately with a reply e.g. 123. This should display after a short delay in the VBA window.
' Function could be called from a larger VBA routine to populate Excel cells or a Word Document with readings etc.
' COM Port can optionally be started with parameters - e.g. START_COM_PORT(Port_Number, "Baud=1200 Data=7 Parity=E")
' Note that VBA remains responsive during wait_for_com function, and also during any extended read/write activites.

Dim VOLTAGE As String: VOLTAGE = "0.000"

Const READ_VOLTS_COMMAND As String = "MEASURE VOLTAGE" & vbCr
Const READ_VOLTS_RESULT As String = "VOLTAGE = "

Const Temp_Text_1 As String = "COM Port Started, Settings = "
Const Temp_Text_2 As String = "Enter a response to MEASURE VOLTAGE on device "
Const Temp_Text_3 As String = "Sending command string to device on COM Port "
Const Temp_Text_4 As String = "Waiting for response from device on COM Port "
Const Temp_Text_5 As String = "Measure Voltage read response from Device = "
Const Temp_Text_6 As String = "Timed out waiting for response from COM Port "
Const Temp_Text_7 As String = "Example function complete , closing COM Port "
Const Temp_Text_8 As String = "Failed to Start COM Port "
Const Temp_Text_9 As String = "Port Number Invalid, Defined Port Number Range = "

Debug.Print

'DEBUG_COM_PORT Port_Number, True                       ' optional - shows port activities and wait countdown loop counter

If Port_Valid Then

    If Not Port_Started(Port_Number) Then START_COM_PORT Port_Number
    
    Kernel_Sleep_MilliSeconds 500                       ' allow local and remote ports to stabilise

    If Port_Started(Port_Number) Then

    Call PRINT_DEBUG_TEXT(Module_Name, "STARTED", Temp_Text_1 & GET_PORT_SETTINGS(Port_Number))
    Call PRINT_DEBUG_TEXT(Module_Name, "RESPOND", Temp_Text_2)
    Call PRINT_DEBUG_TEXT(Module_Name, "SENDING", Temp_Text_3 & Port_Number)

    TRANSMIT_COM_PORT Port_Number, READ_VOLTS_COMMAND   ' send read volts command to remote device

    Call PRINT_DEBUG_TEXT(Module_Name, "WAITING", Temp_Text_4 & Port_Number)
    
    If WAIT_COM_PORT(Port_Number, 10000) Then           ' wait up to 10 seconds (without blocking VBA) for first character
    
            Kernel_Sleep_MilliSeconds 1000              ' allow user 1 second to finish typing any remaining characters
            
            VOLTAGE = RECEIVE_COM_PORT(Port_Number)     ' receive device response back into VOLTAGE variable
            
            Call PRINT_DEBUG_TEXT(Module_Name, "VOLTAGE", Temp_Text_5 & VOLTAGE)
                
    Else
            Call PRINT_DEBUG_TEXT(Module_Name, "TIMEOUT", Temp_Text_6 & Port_Number)
    End If
    
    Call PRINT_DEBUG_TEXT(Module_Name, "STOPPING", Temp_Text_7 & Port_Number)

    STOP_COM_PORT Port_Number

    Kernel_Sleep_MilliSeconds 500                       ' allow local and remote ports to stabilise

Else

    Call PRINT_DEBUG_TEXT(Module_Name, "FAILED", Temp_Text_8 & Port_Number)

End If
 
Else                                                    ' Invalid port number - configure at start of this module

    Call PRINT_DEBUG_TEXT(Module_Name, "INVALID", Temp_Text_9 & COM_PORT_RANGE)

End If

Debug.Print

EXAMPLE = READ_VOLTS_RESULT & VOLTAGE                   ' return read volts result back to function for subsequent use.

End Function

