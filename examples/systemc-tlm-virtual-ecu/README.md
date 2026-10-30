# SystemC/TLM-2.0 virtual automotive ECU

This runnable teaching example models a CPU, address-decoding bus, RAM, and a
memory-mapped CAN controller. It uses the TLM-2.0 loosely timed coding style.

## Memory map

| Address range | Device |
|---|---|
| `0x4000-0x4fff` | CAN controller |
| `0x8000-0x8fff` | RAM |

CAN registers are `CONTROL` (`+0x00`), `TX_ID` (`+0x04`), `TX_DATA`
(`+0x08`), `COMMAND` (`+0x0c`), and `STATUS` (`+0x10`).

## Build

Install a SystemC release that provides the `SystemCLanguage` CMake package,
then point CMake at the installation:

```powershell
cmake -S . -B build -DCMAKE_PREFIX_PATH=C:/path/to/systemc
cmake --build build --config Release
./build/Release/virtual_ecu.exe
```

On a single-config generator, the executable is usually `build/virtual_ecu`.

## Expected behavior

The CPU enables the controller, programs CAN ID `0x123` and data
`0xDEADBEEF`, and writes the transmit command. After 100 microseconds, the CAN
model raises an interrupt. The CPU reads the completion status and saves it to
RAM.

Concepts demonstrated:

- initiator, target, and hierarchical transaction routing
- memory-mapped registers
- `tlm_generic_payload`
- blocking transport and annotated delay
- `sc_event` for internal device behavior
- signal-based interrupt delivery
- address and command error responses

