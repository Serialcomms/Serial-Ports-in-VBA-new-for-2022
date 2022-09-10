# VBA Serial Port routines for Microsoft Office
## New for 2022 - Windows 10, Office 2019 (Excel, Word, Access)

Getting Serial (COM) Ports working as intended in VBA can be surprisingly difficult in certain usage scenarios. 

New VBA routines here will help resolve these issues in Excel, Word and Access (Windows PC versions only).

Functions are straightforward to use with coding style to support infrequent VBA users and developers.

Intended to help implement ad-hoc projects for serial data acquisition or transfer.

No plug-ins, DLLs, ActiveX, licences, payments or registrations are required.  

<details><summary>More Information</summary>
<p>
   
<details><summary>VBA Issues</summary>
<p>

The in-built VBA functions for COM Port data can suffer from the following issues :- 
   
1. Setting port parameters with the VBA open command may not work in some Windows versions e.g.

   `Open "COM1:9600,N,8,1" For Read Access As #1`       \
     _(command line workaround known, settings can revert after reboot)_

2. Attempting to read data when there is none waiting will cause VBA to hang with a 'not responding' message.  
  
   `Get #1, , Read_Data_Byte`  
   
   The new functions address both of these issues, and also where data transfers take longer than the 5-6 second VBA timeout.
   
</p>
</details>   

<details><summary>Background</summary>  
<p>

The legacy of serial comms means that many online solution searches are now time-expired with links to defunct web sites etc.    

New functions here are therefore a fresh start for 2022 and are based largely on Microsoft's Win32 API calls and documentation. 

Developed on Windows 10 (64-Bit) with a local Microsoft Office 2019 Professional (32-Bit VBA7) installation.  

Tested on Office 2016 Professional (64-bit VBA7) and Office 2019 Professional (32-Bit VBA7)    

</p>
</details>

<details><summary>COM Ports</summary>
<p>

Multiple com ports are supported, including physical hardware ports and synthetic virtual software ports. 

All read and write functions are synchronous, in part because not all serial port types support overlapped operation.

Performance on a modern PC is good, with software timing delays required to allow the relatively slow serial com ports to catch up. 

Reading, Writing and Waiting are 'timesliced' to ensure that VBA remains responsive during any extended data transfers or waiting times. 

</p>
</details>

<details><summary>Debugging</summary>
<p>

* Debugging can be set on/off per port with results shown in the VBA immediate window. 

* Extensive debug functionality makes several modules quite verbose. 

* A far more compact version without debug is available in the No-Debug folder. 

</p>
</details>   

<details><summary>Optional steps for Excel only</summary>
<p>

- Remove comment mark before `Option Private Module` to prevent function names appearing in cell formula drop-down lists. 
- Remove comment mark before `Application.Volatile` where indicated to refresh results when functions are used in cells and the worksheet is recalculated (e.g. with F9 key).

</p>
</details>
   
<details><summary>Function List</summary>
<p>   

[COM Port Control](Functions/Function_List_Control.md)
   
[Read/Write/Check Data](Functions/Function_List_Data.md)
   
[Port Signalling Functions](Functions/Function_List_Signalling.md)

[Show Functions](Functions/Function_List_Show.md)

Private functions are not intended to be called directly by users.
  
</p>
</details>   
   
</p>
</details>   
