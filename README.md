# taskTracker-VBA

A smart task tracker system built in Excel using VBA. Designed to help teams manage departmental tasks efficiently.

## Features

- Automatic syncing of tasks to department sheets
- Duplicate detection by Task ID
- Supports task metadata: due dates, priority, status, etc.
- Exported modular code for easy versioning

## Project Structure

- `Excel/` – contains the working `.xlsm` file
- `VBA/` – source code modules and forms
- `Docs/` – screenshots or documentation

## Getting Started

1. Download the `taskTracker.xlsm` file from `Excel/`
2. Enable macros when opening in Excel
3. Run the `SyncTasksToDepartments` macro to organize your task list

## Developer Guide

To view/edit the code:

1. Open `taskTracker.xlsm` in Excel
2. Press `Alt + F11` to open the VBA editor
3. Import modules from the `/VBA/` folder if needed

