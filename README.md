# VBA Serial Port routines for Microsoft Office
## Windows 10/11, Office 2016-2021 (Excel, Word, Access) — hardened 2026 revision

Getting Serial (COM) Ports working as intended in VBA can be surprisingly difficult in certain usage scenarios. 

New VBA routines here will help resolve these issues in Excel, Word and Access (Windows PC versions only).

Functions are straightforward to use with coding style to support infrequent VBA users and developers.

Intended to help implement ad-hoc projects for serial data acquisition or transfer.

No plug-ins, DLLs, ActiveX, licences, payments or registrations are required.  

**2026 hardening revision** — following an independent verification & validation review, both modules were
reworked for correctness and robustness: exact Win32 declarations for 32/64-bit Office, a per-port operation
guard against re-entrant calls, cooperative cancellation, runtime error handling on every public entry point,
partial-write recovery, bounded receives, binary-safe `Byte()` transport, and corrected timing/error decoding.
See [CHANGELOG.md](CHANGELOG.md) for the full list.

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

New functions here are therefore a fresh start and are based largely on Microsoft's Win32 API calls and documentation. 

Originally developed on Windows 10 (64-Bit) with Microsoft Office 2019 Professional (32-Bit VBA7).

Tested on Office 2016 Professional (64-bit VBA7) and Office 2019 Professional (32-Bit VBA7).

The 2026 hardened revision compiles clean on Office 2016+ (VBA7); Win32 `BOOL` results are declared `As Long`,
native NULL pointers are passed as explicit `LongPtr` zeros, and a built-in `ABI_SELF_TEST` verifies structure
sizes (`DCB=28, COMSTAT=12, COMMTIMEOUTS=20`) before the first port is opened on any host.

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

<details><summary>Text vs Binary data</summary>
<p>

* The classic string functions (`read/receive/transmit/send/get/put_com_port`) are for **SBCS/ANSI text**:
  VBA strings pass through an ANSI conversion at the native boundary, so non-ASCII, DBCS and arbitrary
  binary payloads can be transformed.
* For binary protocols and full 0–255 byte values use **`transmit_bytes` / `receive_bytes`**, which move raw
  `Byte()` arrays with no conversion — embedded NUL and all byte values are preserved.

</p>
</details>

<details><summary>Concurrency, cancellation and recovery</summary>
<p>

Every port operation yields to Windows (`DoEvents`) so VBA stays responsive — which also means a timer, event
handler or Ribbon callback can call back into the module while an operation is still running. Each port
therefore has a non-blocking operation guard:

* A second call on the **same** port while one is active is rejected cleanly (returns `False`, an empty
  string, or `-1`) instead of corrupting the operation in progress. Different ports stay independent.
* `port_busy(1)` tests the guard; `cancel_com_port(1)` asks a long receive/transmit/wait to unwind cleanly
  at its next check, keeping any partial data.
* `stop_com_port(1)` on a busy port automatically requests cancellation and returns `False` — call it again
  once the operation has unwound. `stop_com_port(1, 2000)` drains the transmit queue (graceful close) before
  closing; the default remains an abort close.
* `clear_port_busy(1)` is a recovery-only escape hatch after a VBE Ctrl-Break/Reset.
* `port_error_history(1)` returns hardware error bits (overflow/overrun/parity/framing/break) accumulated by
  status polling, so no error evidence is lost between checks.

</p>
</details>

<details><summary>Debugging</summary>
<p>

* Debugging can be set on/off per port with results shown in the VBA immediate window. 

* Extensive debug functionality makes several modules quite verbose. 

* Runtime errors are trapped at every public entry point, recorded per port, and reported as the function's
  normal failure value — no unhandled VBA error escapes to the caller.

</p>
</details>  

<details><summary>Compile test workbooks</summary>
<p>

`SERIAL_PORT_TEST_MAIN.xlsm` and `SERIAL_PORT_TEST_NODEBUG.xlsm` contain the two modules imported and
compiled on Office 16.0, with headless smoke tests passed (invalid port numbers only — no hardware touched).

**Never import both modules into the same VBA project** — they share public names by design.

</p>
</details>

<details><summary>Other Versions</summary>
<p> 

* [Original](https://github.com/Serialcomms/Serial-Ports-in-VBA-new-for-2022/tree/Original) (original 2022 version)

* [No-Debug](No-Debug)      (more compact)

* [Simplified](https://github.com/Serialcomms/Serial-Ports-in-VBA-Simple-2022)   (single com port)

* [Minimal](https://github.com/Serialcomms/Serial-Ports-in-VBA-Extra-Simple-2022)   (single com port, no settings)

* [VBA6 / 32-Bit](https://github.com/Serialcomms/Serial-Ports-in-VBA6-legacy-for-2022)  (legacy Windows/Office)

</p>
</details> 
   
<details><summary>Alternatives (Excel Only)</summary>
<p> 

* [Microsoft Data Streamer for Excel](https://learn.microsoft.com/en-us/microsoft-365/education/data-streamer/)   
 
</p>
</details>   
     
<details><summary>Optional steps for Excel only</summary>
<p>  

- Functions can be used directly in Worksheet cells as formulas where appropriate.  
- Remove comment mark before `Option Private Module` to prevent function names appearing in cell formula drop-down lists.  
- Remove comment mark before `Application.Volatile` where indicated to refresh results when functions are used in cells and the worksheet is recalculated (e.g. with F9 key).
- Port status functions (`device_ready`, `clear_to_send`, etc.) are safe to use in cells while transfers run —
  they never disturb an operation in progress.

</p>
</details>

<details><summary>Optional Ribbon Customisation</summary>
<p>

[Office 2010 XML](/Ribbon/RIBBON_2010.xml) and [SERIAL_PORT_RIBBON](/Ribbon/SERIAL_PORT_RIBBON.bas) example files are available in the [Ribbon](/Ribbon) folder. 
   
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
