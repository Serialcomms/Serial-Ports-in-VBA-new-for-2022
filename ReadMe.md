# VBA Serial Port routines for Microsoft Office
## New for 2022 - Windows 10, Office 2019 (Excel, Word, Access)

Getting Serial (COM) Ports working as intended in VBA can be surprisingly difficult in certain usage scenarios. 

The legacy nature of serial comms is such that many existing VBA solutions are now rather dated with references to defunct web sites etc. 

These new VBA routines are a fresh start for 2022 and based largely on Microsoft's Win32 API calls and documentation. 

Functions work in Excel, Word and Access (Windows versions only) with Macros enabled.

No plug-ins, DLLs, ActiveX, licences, payments or registrations are required.  

Developed on Windows 10 (64-Bit) with a local Microsoft Office 2019 Professional (32-Bit) installation.  

Tested on Office 2016 Professional (64-bit) and Office 2019 Professional (32-Bit).

Functions are straightforward to use and intended to help implement ad-hoc projects for serial data acquisition or transfer.
Coding style supports infrequent VBA users and developers.

Standard in-built VBA functions to handle COM port data can suffer from two issues :-

1. Setting port parameters with the VBA open command may not work in some Windows versions e.g.

   `Open "COM1:9600,N,8,1" For Read Access As #1`       \
     _(command line workaround known, settings can revert after reboot)_

2. Attempting to read data when there is none waiting will cause VBA to hang with a 'not responding' message.  
  
   `Get #1, , Read_Data_Byte`  
  
The new functions address both of these issues, and also where data transfers take longer than the 5-6 second VBA timeout.

Debugging can be set on/off per port with results shown in the VBA immediate window. Extensive debug functionality makes several modules quite verbose, performance impact is however minimal. 

Performance on a modern PC is good, with software timing delays required to allow the relatively slow serial com ports to catch up.  Multiple com ports are supported, including physical hardware ports and virtual software ports. 

All read and write functions are synchronous, in part because not all serial ports support overlapped operation. Reading, Writing and Waiting are 'timesliced' to ensure that VBA remains responsive during any extended data transfers or waiting times. 

Optional steps for Excel only - 

- Remove comment mark before `Option Private Module` to prevent function names appearing in cell formula drop-down lists. 
- Remove comment mark before `Application.Volatile` where indicated to refresh results when functions are used in cells and the worksheet is recalculated (e.g. with F9 key).

Main user-defined functions are as follows. First parameter is a valid COM Port number on host. COM 1 is used here for example. Min/Max com port numbers are defined in the declarations section at the start of the module.

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `debug_com_port(1)`                  | Toggles debug messaging on/off (debug results in VBA Immediate window)                                        |
| `debug_com_port(1,True)`             | Set port debug messaging on                                                                                   |          
| `debug_com_port(1,False)`            | Set port debug messaging off                                                                                  |  
| `start_com_port(1)`                  | Starts port with existing settings. Returns `True` if successful, `False` if start fails for any reason.      | 
| `start_com_port(1,"Baud=1200")`      | Starts port with settings as supplied. Returns `True` or `False` as above.                                    |
| `check_com_port(1)`                  | Returns number of input characters waiting to be read (no delay). Return value -1 indicates error.            |
| `wait_com_port(1)`                   | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`               | As above, can optionally specify wait time (500) in milliseconds. Timesliced to avoid VBA hanging.            |  
| `get_com_port(1)`                    | Receives a single character string from a started com port.                                                   |
| `put_com_port(1,"A")`                | Sends a single character to a started com port. Returns `True` if successful, `False` if fail.                |
| `read_com_port(1,20)`                | Reads up to specified number (20) of characters. No delay, max characters = approx 1 second timeslice.        |
| `send_com_port(1,V)`                 | Sends variable V. Function converts V to String and calls transmit_com_port.                                  |
| `receive_com_port(1)`                | Receives all data from port, timesliced for low port speeds and/or large data transfers.                      |
| `transmit_com_port(1,"QWERTY")`      | Sends string to port, in timeslices of approx 1 second to avoid VBA 'not responding'                          |
| `device_ready(1)`                    | Returns `True` if port started and DSR Signal (input) active.                                                 |
| `clear_to_send(1)`                   | Returns `True` if port started and CTS Signal (input) active.                                                 |
| `stop_com_port(1)`                   | Stops port and hands control of it back to Windows.                                                           |

Other Public functions such as `show_port_errors(1)` etc. should only be used in the Immediate window for further information if required.
Private functions are not intended to be called directly by users.
