# Changelog

All notable changes to the serial-port modules (`SERIAL_PORT_VBA.bas` and
`No-Debug/SERIAL_PORT_VBA_NO_DEBUG.bas`) are documented here. Unless a change is marked
*(main module only)*, it was applied identically to both modules.

## [2026-07-23] — Hardening revision

Driven by an independent verification & validation review (20 findings: 2 critical, 22 high,
11 medium, 1 low across both modules) and a follow-up best-practice integration pass. Both
modules import and **compile clean on Office 16.0**, verified via COM automation with 50+
headless smoke tests (`SERIAL_PORT_TEST_MAIN.xlsm` / `SERIAL_PORT_TEST_NODEBUG.xlsm`).

### Fixed — defects

- **Crash on invalid port number in `RECEIVE_COM_PORT`** (critical path). The function indexed
  the port array outside the validity check, raising a VBA subscript error for out-of-range
  ports. All results are now accumulated in locals; the array is never indexed unvalidated.
- **Win32 `BOOL` declared as VBA `Boolean`.** All 13 native `BOOL`-returning declarations are
  now `As Long` (Win32 BOOL is 32-bit; VBA Boolean is 16-bit) and every call site tests
  `<> 0` explicitly.
- **Native NULL pointers passed as untyped Optional Variants.** `lpSecurityAttributes`,
  `hTemplateFile` and `lpOverlapped` are now `ByVal ... As LongPtr` with an explicit
  `NULL_POINTER` (0). An omitted Optional Variant passes a pointer to a "missing" Variant —
  a non-NULL value — where the API contract requires NULL.
- **`DCBlength` never initialised.** `DCB.LENGTH_DCB = LenB(DCB)` is set before `GetCommState`
  and re-set after `BuildCommDCB` (which does not populate it) before `SetCommState`.
- **`WriteFile` success return ignored; partial writes reported as success.** Synchronous write
  now requires API success *and* bytes-written = bytes-requested, retries short writes for the
  remaining bytes (bounded by a stall cap), and resets the byte count before every call so a
  failure can never be judged against stale state.
- **Chunked transmit masked earlier failures.** The chunk loop returned only the last chunk's
  status. It now fails fast on the first failed chunk and returns success only when the total
  bytes sent equal the requested length; the total is retained in the port profile.
- **QueryPerformanceCounter ticks treated as fixed 10 MHz.** `QueryPerformanceFrequency` is now
  queried once, cached, and all tick→time conversions use the counter/frequency ratio. Hosts
  and VMs with other counter frequencies now time correctly.
- **`GET_HOST_MILLISECONDS` overflow.** Return type changed `Long` → `Currency`; the previous
  absolute-millisecond `Long` overflowed after ~24.9 days of host uptime.
- **1.5 stop bits calculated as 2** in frame timing. DCB stop-bits value 1 now contributes 1.5
  bit times (`Double` arithmetic); baud rate validated > 0 before division.
- **`DECODE_PORT_ERRORS` decoded the wrong bit masks** *(main module only)*. It tested event
  (EV_) masks — using the same bit for overrun, parity and framing — instead of the
  ClearCommError error (CE_) masks 1/2/4/8/16. Each bit is now tested independently, so
  combined masks decode correctly.
- **`GET_COM_PORT` ignored the `ReadFile` return.** A byte is returned only when the API
  succeeded and reported exactly one byte; the native error is recorded otherwise.
- **`PUT_COM_PORT` sent a NUL for empty input and never checked bytes written.** Empty input is
  rejected; success requires API success and bytes-written = 1.
- **`SHOW_PORT_ALL` had no return type or result.** Now `As Boolean`, aggregating the child
  show functions' results.
- **Stale `Err.LastDllError` reported as current diagnostics.** Still captured immediately after
  every native call, but cleared to `SUCCESS` when the call succeeded; error text is only
  formatted on failure paths.
- **Stale byte counts.** `Synchronous_Bytes_Read`/`_Sent` are zeroed before each native call.
- **Timestamp mixed time bases** *(main module only)*. `TIMESTAMP` combined local `Time()` with
  UTC `GetSystemTime` milliseconds; now uses `GetLocalTime` with zero-padded milliseconds.
- **Divide-by-zero paths** guarded: frame time with baud 0, port values with unusable byte
  rates (configuration now fails instead of starting a port with a zero timeslice), and debug
  throughput calculations with zero elapsed time.
- **`For ... Step 0` hang.** Chunked transmit refuses to run when `Timeslice_Bytes < 1`.
- **Latent `Static` state leak** *(main module only)*. `SHOW_PORT_DCB` and `SHOW_PORT_STATUS`
  were `Static` functions whose persisted locals could return a stale `True` after an
  invalid-port call. Now plain functions.

### Added — concurrency, lifecycle and safety

- **Per-port operation guard** (`PORT_ENTER`/`PORT_LEAVE` on a `Port_Busy` flag). Every public
  operation that yields through `DoEvents` — start, stop, wait, read, receive, transmit, send,
  put, get, check, signal, RTS, and the state-mutating show functions — claims the port for one
  operation. A re-entrant call on the same port (timer, event, Ribbon callback, nested macro)
  is rejected with the function's normal failure value instead of corrupting the operation in
  progress. Ports are independent. Guarded public functions call private cores
  (`STOP_PORT_CORE`, `TRANSMIT_PORT_CORE`) rather than other guarded publics.
- **Cooperative cancellation** (`CANCEL_COM_PORT`). Long operations check for cancel after
  every yield and unwind cleanly, preserving partial data. `STOP_COM_PORT` on a busy port
  automatically requests cancel and returns `False` for retry.
- **Post-yield state verification** (`OPERATION_INTERRUPTED`). After every `DoEvents`, long
  operations re-verify cancel state and handle identity; if the port was closed or restarted
  underneath them they stop touching it immediately.
- **Graceful close option.** `STOP_COM_PORT(Port, Drain_MilliSeconds)` waits up to the given
  time for the transmit queue to drain before closing. Default (0) remains the original
  purge-then-close abort.
- **Bounded receive.** `RECEIVE_COM_PORT(Port, [Max_Bytes], [Total_MilliSeconds])` — optional
  byte and total-time limits stop a continuous data source from keeping the call running and
  growing memory indefinitely. Defaults (0) preserve the original unbounded behaviour.
- **Binary-safe transport.** `TRANSMIT_BYTES(Port, Data())` and
  `RECEIVE_BYTES(Port, Data(), [Max_Bytes], [Total_MilliSeconds])` move raw `Byte()` arrays via
  `ByRef ... As Any` declarations — no string/code-page conversion, full 00–FF and embedded-NUL
  fidelity. Receive grows a doubling buffer, returns an exact-length array plus count, and
  distinguishes no-data (0) from failure (−1). The classic string functions remain, documented
  as SBCS/ANSI-text transport.
- **ABI self-test.** Public `ABI_SELF_TEST` verifies marshalled structure sizes (`DCB=28`,
  `COMSTAT=12`, `COMMTIMEOUTS=20`, `SYSTEMTIME=16` in main); runs automatically (cached) before
  the first `START_COM_PORT`, which refuses to open a port on mismatch.
- **Error-evidence preservation.** Every `ClearCommError` poll ORs the consumed CE_ bits into a
  per-port `Error_History` (cleared on port open). `SHOW_PORT_ERRORS` *(main)* reports current
  + history exactly once; `PORT_ERROR_HISTORY(Port, [Clear])` is the public accessor. Status
  polling can no longer silently destroy hardware error evidence.
- **Runtime error handling everywhere.** `On Error GoTo Clean_Exit` with a single cleanup path
  (guard release included) on every public entry point that touches the port or converts input
  — 30+ handlers per module. Trapped errors are recorded per port (`VBA_Error`, separate from
  the Win32 `DLL_Error`) and reported as the function's normal failure value. Logging helpers
  (`TIMESTAMP`, `PRINT_SHOW_TEXT`, `DECODE_PORT_ERRORS`) use safe fallbacks so a failure inside
  logging can never cascade.
- **Busy-safe status readers.** `DEVICE_READY`, `DEVICE_CALLING`, `CLEAR_TO_SEND`,
  `CARRIER_DETECT` read modem signals into locals and only persist shared diagnostics when the
  port is idle — a worksheet/timer status poll can never clobber an in-flight operation's
  state, and still returns the true signal.
- **Recovery and introspection.** `Port_Busy(Port)` queries the guard; `CLEAR_PORT_BUSY(Port)`
  recovers from a VBE Ctrl-Break/Reset that bypassed cleanup.
- **Compile test workbooks.** `SERIAL_PORT_TEST_MAIN.xlsm` and `SERIAL_PORT_TEST_NODEBUG.xlsm`
  hold each module imported and compiled on Office 16.0 (separate projects — the modules share
  public names and must never be loaded together).

### Changed — behaviour notes

- A runtime error inside a guarded call is now trapped and returned as a failure result instead
  of raised to the caller.
- A same-port re-entrant call is rejected (`False` / empty string / `-1`) instead of running
  concurrently with the operation in progress.
- `SIGNAL_COM_PORT` retains function code 7: it is `RESETDEV` in the Windows SDK (winbase.h),
  omitted from the online `EscapeCommFunction` table but valid. Codes 1–9 are now documented in
  the module header.
- `DEBUG_COM_PORT` leaves the debug state unchanged (and returns it) when given a value that
  cannot convert to Boolean, instead of raising a type-mismatch error.
- `GET_PORT_SETTINGS` returns `ERROR-READING-SETTINGS` instead of raising on an internal error.
- Function documentation updated: [Control](Functions/Function_List_Control.md),
  [Data](Functions/Function_List_Data.md), [Signalling](Functions/Function_List_Signalling.md).

### Compatibility

All original public function signatures are unchanged; new parameters are `Optional` with
defaults preserving prior behaviour, and new functions are purely additive. Existing callers —
including the [Ribbon](Ribbon) module — compile and run unmodified. `GET_HOST_MILLISECONDS`
callers should note its return type changed `Long` → `Currency` (overflow fix).

### Validation status

Static verification and host compile evidence are complete (see the audit trail in the review
package's `Serial_VBA_Findings_Disposition.md`). Hardware-in-the-loop validation — loopback
00–FF byte integrity, cancellation under live traffic, unplug/recovery, soak — remains to be
executed against a real or virtual COM port pair, on 32-bit and 64-bit Office.

## Earlier history

Pre-2026 changes were not tracked in a changelog; see the
[git commit history](https://github.com/Serialcomms/Serial-Ports-in-VBA-new-for-2022/commits)
for the original 2022 development record.
