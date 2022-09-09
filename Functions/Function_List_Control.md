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
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows                                                              |

* Debug results are shown in the VBA Immediate Window (Control-G)
* Debug functions return `True` or `False` to indicate debug state
* Other functions return `True` or `False` to indicate success or failure

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 
  
