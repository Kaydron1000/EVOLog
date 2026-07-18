# EVOLog - Evocation Logging

Evoke Verbose Output Logging - Channel multiple log outputs with EVOLog.

> **Status:** This library is no longer in active development.

EVOLog is a VBA logging library inspired by Serilog-style logging. It was built to make structured logging more accessible inside Microsoft Office VBA projects while still keeping the API simple and flexible.

## Overview

EVOLog provides a logger that can send log messages to multiple conduits at once. A conduit is any destination that implements the `ILogConduit` interface, such as:

- Text files
- Text boxes or forms
- Counters or progress displays
- Another `cEVOLogger` instance

This design makes it easy to route the same log message to multiple outputs without duplicating logging logic in your application.

## Main Components

### `cEVOLogger`

`cEVOLogger` is the core logger class. It manages conduits and coordinates how log messages are dispatched.

Key members include:

- `LoggerName`
- `LoggingLevelNames`
- `BatchOutput`
- `BatchSetCount`
- `CounditsCount`
- `GetCouduitNames`
- `AddConduit`
- `GetConduit`
- `RemoveConduit`
- `ClearConduits`
- `LogArtifact`
- `LogArtifactObject`
- `Init`

### `ILogConduit`

`ILogConduit` defines the contract for any logging destination used by EVOLog. If a class implements this interface, it can receive log output from `cEVOLogger`.

## How It Works

1. Create a `cEVOLogger` instance.
2. Add one or more conduits.
3. Send log entries through the logger.
4. Each configured conduit receives the output.

## Example Use Cases

- Writing log output to a file during automation
- Displaying live progress messages in a VBA userform
- Forwarding logs to multiple destinations at the same time

## Project Notes

- This repository is preserved for reference and documentation.
- Existing code may still be useful as a lightweight logging foundation for VBA projects.
- No new feature development is planned.

## License

This project is licensed under the GNU General Public License v2.0.
