# Autostart Manager

**Autostart Manager** is a user-friendly GUI application designed for managing autostart programs on Windows. It allows users to add, remove, and manage applications that launch at system startup, as well as execute programs in batches. The application supports light and dark themes, drag-and-drop functionality, and comprehensive logging of all operations.

## Features

### Autostart Management
- **Add Programs to Autostart**:
  - Supports `.exe`, `.bat`, and `.cmd` files.
  - Option to run programs with administrator privileges via Windows Task Scheduler.
  - Add programs to the Windows Registry for standard autostart.
  - Drag-and-drop support for easy file selection.
- **Remove Programs from Autostart**:
  - Remove entries from the Registry or Task Scheduler.
  - Confirmation prompt to prevent accidental deletions.
- **View Current Autostart Programs**:
  - Displays all programs configured for autostart (from Registry and Task Scheduler).
  - Sortable table for easy navigation.

### Batch Execution
- **Manage Batch Program List**:
  - Drag-and-drop files to add them to the batch execution list.
  - Reorder programs by dragging table rows.
  - Remove individual programs or clear the entire list.
- **Start and Stop Batch Execution**:
  - Launch all programs in the list with a single click.
  - Stop all running batch processes.
  - Real-time status display (Running/Stopped) in the table.

### Additional Features
- **Interface Themes**:
  - Toggle between light and dark themes.
  - Theme preference is saved between sessions.
- **Import/Export Settings**:
  - Export batch program list and theme settings to a JSON file.
  - Import settings from a JSON file for quick configuration.
- **Logging**:
  - Logs all operations to `autostart_manager.log` with rotation (1 MB max, up to 5 backups).
  - View the last 100 log lines in the GUI.
  - Option to clear logs.
- **Self-Autostart Management**:
  - Add or remove Autostart Manager itself from autostart with one click.
