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
| `receive_com_port(1)`           | Yes  | Receives all data from port [^4][^3]                                                                          |
| `transmit_com_port(1,"QWERTY")` | Yes  | Sends supplied string QWERTY to port [^4]                                                                     |
| `transmit_com_port(1,COMMANDS)` | Yes  | Sends supplied string constant or variable COMMANDS to port [^4]                                              |

* Functions shown as TS=No return within a few milliseconds. 
* Functions shown as TS=Yes are timesliced to avoid VBA hanging with a 'not responding' message.

[^1]:  Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 

[^2]:  Maximum number of characters read is approximately = (baud rate / 10)  
       
[^3]:  Function includes read wait and exit timers and returns when exit timer expires.  

[^4]:  Function can block for extended periods with VBA remaining responsive before returning.  
       
