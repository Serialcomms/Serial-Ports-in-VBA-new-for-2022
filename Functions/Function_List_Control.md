# VBA Serial Port Functions

## start, stop and debug COM ports

First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `debug_com_port(1)`                  | Toggles debug messaging on/off                                                                                |
| `debug_com_port(1,True)`             | Set port debug messaging on                                                                                   |
| `debug_com_port(1,False)`            | Set port debug messaging off                                                                                  |
| `start_com_port(1)`                  | Starts port with existing settings                                                                            |
| `start_com_port(1,"Baud=1200")`      | Starts port with settings as supplied                                                                         |
| `start_com_port(1,SCANNER)`          | Starts port with settings defined in string constant or variable e.g. SCANNER                                 |
| `get_port_settings(1)`               | Returns string [^2] with port settings or error text                                                          |
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows. Queued transmit/receive data is purged (abort close)        |
| `stop_com_port(1,2000)`              | As above, but first waits up to 2000mS for queued transmit data to be sent (graceful close)                   |
| `port_busy(1)`                       | Returns `True` while an operation on the port is in progress [^3]                                             |
| `cancel_com_port(1)`                 | Requests cooperative cancel of the running operation; it unwinds cleanly at its next check [^4]               |
| `clear_port_busy(1)`                 | Recovery only - clears the port operation guard after a VBE Ctrl-Break/Reset [^3]                             |
| `abi_self_test()`                    | Verifies Win32 structure sizes (DCB=28, COMSTAT=12, COMMTIMEOUTS=20); runs automatically before first start   |
| `port_error_history(1)`              | Returns OR-accumulated comms-error bits captured by status polls (1=overflow 2=overrun 4=parity 8=frame 16=break) |
| `port_error_history(1,True)`         | As above, and clears the accumulated history after reading                                                    |

* Debug results are shown in the VBA Immediate Window (Control-G)
* Debug functions return `True` or `False` to indicate debug state
* Other functions return `True` or `False` to indicate success or failure

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 
[^2]: e.g. 9600-8-N-1 , PORT-NOT-STARTED, INVALID-PORT
[^3]: Each port has one operation guard. A port function called again for the same port (for example from a timer,
      event handler or Ribbon callback while an earlier call is still running) is rejected rather than allowed to
      interfere with the operation in progress, and returns its usual failure value (`False`, empty string, or -1).
[^4]: `stop_com_port` on a busy port automatically requests cancellation and returns `False`; call it again once
      the running operation has unwound. Data received or sent before the cancel point is preserved.
