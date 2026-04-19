# SSMS Executor

SQL Server Management Studio (SSMS) extension for executing the current statement based on the cursor position - no need to highlight text first.

## Features

### Execute Current Statement

If your script contains multiple SQL statements, place your cursor anywhere within the statement you want to run and press `Ctrl+Shift+E`. The extension automatically detects the statement boundaries and executes only that statement, leaving the rest of your script untouched. Your original cursor position and any text selection are restored after execution.

This is available from the **Tools** menu (**Execute Statement**) or from the toolbar.

### Execute Inner Statement

When working with compound or nested SQL blocks (such as `BEGIN...END`), the **Execute Inner Statement** command lets you run just the innermost statement at the cursor position rather than the entire block. This is useful for debugging or testing individual statements inside stored procedures, `IF` blocks, or transaction wrappers.

This is available from the **Tools** menu (**Execute Inner Statement**).

### Options

You can configure the extension under **Tools > Options > SSMS Executor**:

| Option | Description |
|---|---|
| **Execute inner statements** | When enabled, the default `Ctrl+Shift+E` shortcut will execute the inner statement instead of the full block. |

# Documentation

[Getting started guide](https://github.com/tkwj/ssms-executor/wiki)

[Release notes for released versions and other builds](https://github.com/tkwj/ssms-executor/wiki/Release-notes)

# Downloads/builds

## SQL Server Management Studio (SSMS) 21 & 22 Extension

You can download the extension from the [Releases section](https://github.com/tkwj/ssms-executor/releases)
or build it yourself using the provided source code.

This version only supports SSMS 21 & 22.
For prior SSMS versions, please consult the original version here: https://github.com/devvcat/ssms-executor/
