# EVOLog

EVOLog is a Serilog-inspired logging library for VBA and Microsoft Office projects. It lets you create a single logger, send one log event through it, and route that event to multiple destinations such as the Immediate window, a text file, a worksheet, a textbox, or another logger.

> **Status:** Archived reference library. This project is no longer in active development, but it remains a useful example of structured-style logging patterns in VBA.

## Why use EVOLog?

EVOLog was built to make logging in VBA less repetitive and more flexible.

- Write a log event once and fan it out to multiple outputs
- Filter output by logging level at each conduit
- Format messages through reusable templates
- Batch log delivery before flushing to outputs
- Capture logs in memory and replay them later

## Getting Started

The repository contains the library in two forms:

- Exported VBA source files in `/home/runner/work/EVOLog/EVOLog/src`
- An Excel workbook version in `/home/runner/work/EVOLog/EVOLog/EVOLog.xlsm`

To adopt EVOLog in your own Office/VBA project:

1. Import the classes and modules you need from `/home/runner/work/EVOLog/EVOLog/src` into your VBA project.
2. Create a `cEvoLogger`.
3. Initialize one or more conduits.
4. Attach those conduits to the logger.
5. Send log artifacts through the logger and flush when needed.

## Core Concepts

### `cEvoLogger`

`cEvoLogger` is the main logging object. It owns the conduit collection, creates and dispatches `cEvoArtifact` log payloads, and manages batching behavior.

Key responsibilities:

- Naming a logger with `Init`
- Attaching and removing conduits
- Logging messages with `LogArtifact`
- Sending either immediate or batched output
- Flushing buffered artifacts with `FlushBatchedLogArtifacts`

By default, the logger starts with batching enabled and a batch size of 20.

### `ILogConduit`

`ILogConduit` is the interface every output destination implements. A conduit decides where a log artifact goes and whether it should be emitted based on its configured logging level.

### `cEvoArtifact`

`cEvoArtifact` is the log event payload. It carries the message, timestamp, logging level, and optional parameters used in message templates.

### Logging Levels

EVOLog defines these levels:

1. `Verbose`
2. `Debugg`
3. `Information`
4. `Warning`
5. `Error`
6. `Fatal`

Conduits can ignore anything below their configured threshold.

### Batching

When batching is enabled, log artifacts are buffered until the batch fills or you explicitly flush them. This makes it possible to accumulate output and emit it together instead of writing one message at a time.

## Built-in Conduits

The library includes these conduit implementations:

- `cLogConduit_Immediate` - writes to the VBA Immediate window
- `cLogConduit_File` - writes to a text file
- `cLogConduit_TextBox` - appends messages to a textbox or userform control
- `cLogConduit_MemoryLogger` - stores artifacts in memory for later replay
- `cLogConduit_Counter` - counts messages by logging level
- `cLogConduit_ExcelWorksheet` - writes log output into a worksheet
- `cLogConduit_EvoLogger` - forwards log artifacts into another `cEvoLogger`

## Simple Usage Flow

Typical usage looks like this:

1. Create a logger and call `Init`
2. Create one or more conduits and initialize each one
3. Attach conduits with `AddConduit`
4. Call `LogArtifact` for each event you want to record
5. Call `FlushBatchedLogArtifacts` to push any buffered output

Example:

```vb
Dim logger As cEvoLogger
Dim immediateConduit As cLogConduit_Immediate

Set logger = New cEvoLogger
Set immediateConduit = New cLogConduit_Immediate

logger.Init "MyLogger"
immediateConduit.Init "Immediate", Information

logger.AddConduit immediateConduit
logger.LogArtifact Information, "Application started for {{0}}", Environ$("Username")
logger.FlushBatchedLogArtifacts
```

## Capabilities

EVOLog currently supports:

- Multiple outputs attached to a single logger
- Per-conduit filtering by logging level
- Message templates with placeholders such as `{{0}}`
- Batched delivery with configurable batch size
- In-memory capture and later rechanneling through `cLogConduit_MemoryLogger`
- Logger-to-logger forwarding for chained output flows

## Repository Layout

- `/home/runner/work/EVOLog/EVOLog/src` - exported VBA source files
- `/home/runner/work/EVOLog/EVOLog/EVOLog.xlsm` - workbook version of the library
- `/home/runner/work/EVOLog/EVOLog/src/UnitTests.bas` - sample/manual test routines

## Intended Audience

EVOLog is most useful for developers maintaining VBA automation, Office add-ins, Excel tooling, or legacy Microsoft Office solutions that need better logging structure without leaving VBA.

## Status and Maintenance

This repository is preserved as stable reference code rather than an actively evolving product.

- No new feature development is planned
- The code remains useful for learning from or adapting into VBA projects
- The workbook and exported source files are retained for inspection and reuse

## License

This project is licensed under the GNU General Public License v2.0. See `/home/runner/work/EVOLog/EVOLog/LICENSE`.
