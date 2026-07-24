# VBA Serial Port Functions
## send, receive, check and wait for data

First parameter (1) is a valid[^1] and started COM Port number on host PC.

| VBA Function                    |  TS  | Description                                                                                                   |
| :-------------------------------|:----:| :-------------------------------------------------------------------------------------------------------------|
| `check_com_port(1)`             | No   | Returns number of input characters waiting to be read (no wait). Return value -1 indicates error.             |
| `wait_com_port(1)`              | Yes  | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`          | Yes  | As above, specify maximum wait time (500) in milliseconds.                                                    |
| `get_com_port(1)`               | No   | Receives a single-character string.                                                                           |
| `put_com_port(1,"A")`           | No   | Sends a single-character string.                                                                              |
| `read_com_port(1)`              | No   | Reads an unspecified number [^2] of waiting characters.                                                       |
| `read_com_port(1,20)`           | No   | Reads up to specified number [^2] (20) of waiting characters.                                                 |
| `send_com_port(1,V)`            | Yes  | Sends variable V. Function converts V to string and calls `transmit_com_port`. [^4]                           |
| `send_com_port(1,$B$5)`         | Yes  | Sends contents of Cell $B$5 [^5] to com port (Excel Worksheet Only)                                           |
| `receive_com_port(1)`           | Yes  | Receives all data from port [^4][^3]                                                                          |
| `receive_com_port(1,4096)`      | Yes  | As above, but stops once at least 4096 bytes have been received [^6]                                          |
| `receive_com_port(1,0,5000)`    | Yes  | As above, but stops after 5000mS total, however much data has arrived [^6]                                    |
| `transmit_com_port(1,"QWERTY")` | Yes  | Sends supplied string QWERTY to port [^4]                                                                     |
| `transmit_com_port(1,COMMANDS)` | Yes  | Sends supplied string constant or variable COMMANDS to port [^4]                                              |
| `transmit_bytes(1,B)`           | Yes  | Binary-safe: sends Byte array B exactly as supplied [^7]                                                      |
| `receive_bytes(1,B)`            | Yes  | Binary-safe: receives raw bytes into Byte array B, returns byte count [^7]                                    |
| `receive_bytes(1,B,4096,5000)`  | Yes  | As above with optional max-byte and total-millisecond bounds [^6][^7]                                         |


* Functions shown as TS=No return within a few milliseconds. 
* Functions shown as TS=Yes are timesliced to avoid VBA hanging with a 'not responding' message.

[^1]:  Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 

[^2]:  Maximum number of waiting characters read is approximately = (baud rate / 10)  
       
[^3]:  Function includes read wait and exit timers and returns when exit timer expires.  

[^4]:  Function can block for extended periods with VBA remaining responsive before returning.  

[^5]:  Excel will re-send if Cell $B$5 value changes

[^6]:  Optional receive bounds. Both default to zero (no limit), which keeps the original behaviour of returning
       only when the read exit timer expires. Use them when a continuous data source could otherwise keep
       `receive_com_port` running indefinitely. The byte limit can be overshot by up to one read timeslice,
       because data already taken from the driver is never discarded.

[^7]:  The string functions above are for SBCS/ANSI text: VBA strings pass through an ANSI conversion at the
       native boundary, so non-ASCII, DBCS and arbitrary binary payloads can be transformed. For binary
       protocols and full 0-255 byte values use `transmit_bytes` / `receive_bytes`, which move raw `Byte()`
       data with no conversion. `receive_bytes` returns the byte count (0 = no data, -1 = failure/busy) and
       sizes the array to exactly the bytes received.
       
