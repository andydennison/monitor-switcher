# Monitor Switcher

A Windows 11 system tray application that automatically switches your monitor input when you use a KM (Keyboard/Mouse) switch between your work laptop and home machine.

## Features

- 🖥️ **Automatic Monitor Switching**: Detects when you use your KM switch and automatically changes monitor input
- 🎛️ **System Tray Interface**: Runs quietly in the background with easy access from the system tray
- ⚙️ **Configurable**: Set which HDMI/DisplayPort your home machine and work laptop use
- 🔧 **Manual Override**: Manually switch inputs from the system tray menu when needed
- 💾 **Persistent Settings**: Configuration is saved and persists across restarts

## How It Works

The application monitors USB device changes on your home machine to detect when you switch your KM (keyboard/mouse) switch. When it detects a switch:
- **Devices appear** → You switched TO your home machine → Monitor switches to home input
- **Devices disappear** → You switched TO your work laptop → Monitor switches to work input

The monitor input is controlled using DDC/CI (Display Data Channel Command Interface), a standard protocol for monitor control.

## Requirements

- Windows 11 (or Windows 10)
- Python 3.8 or higher
- Monitor that supports DDC/CI (most modern monitors do)
- KM switch connected via USB

## Installation

1. **Clone or download this repository**

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify your monitor supports DDC/CI**:
   - Most modern monitors support DDC/CI by default
   - Some monitors have a setting in the OSD menu to enable it (look for "DDC/CI" in settings)

## Configuration

### First Time Setup

1. **Run the application**:
   ```bash
   python monitor_switcher.py
   ```

2. **Configure inputs**:
   - Right-click the system tray icon (look for "M" icon)
   - Click "Configure"
   - Set your **Home Machine Input** (e.g., HDMI-1)
   - Set your **Work Laptop Input** (e.g., HDMI-2)
   - Click "Save"

### Configuration Options

The application creates a `config.json` file with these settings:

- **home_machine_input**: The monitor input your home machine is connected to
  - Options: `HDMI-1`, `HDMI-2`, `DisplayPort-1`, `DisplayPort-2`, `DVI-1`, `VGA-1`

- **work_laptop_input**: The monitor input your work laptop is connected to
  - Same options as above

- **monitor_index**: Which monitor to control if you have multiple (0 = first monitor, 1 = second, etc.)

- **check_interval**: How often to check for device changes (in seconds, default: 2.0)

- **last_active_machine**: Tracks which machine was last active

### Example Configuration

```json
{
    "home_machine_input": "HDMI-1",
    "work_laptop_input": "HDMI-2",
    "monitor_index": 0,
    "check_interval": 2.0,
    "last_active_machine": "home"
}
```

## Usage

### Running the Application

```bash
python monitor_switcher.py
```

The application will:
1. Start monitoring for KM switch changes
2. Create a system tray icon (letter "M")
3. Automatically switch monitor inputs when you use your KM switch

### System Tray Menu

Right-click the system tray icon to access:

- **Switch to Home**: Manually switch to home machine input
- **Switch to Work**: Manually switch to work laptop input
- **Configure**: Open configuration window
- **Quit**: Exit the application

### Running on Windows Startup

To start the application automatically when Windows starts:

1. **Create a shortcut**:
   - Right-click `monitor_switcher.py`
   - Create shortcut

2. **Add to Startup folder**:
   - Press `Win + R`
   - Type `shell:startup` and press Enter
   - Move the shortcut to this folder

Alternatively, compile to an executable (see below) and add that to startup.

## Building an Executable

To create a standalone `.exe` file (so you don't need Python installed):

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Build the executable**:
   ```bash
   pyinstaller --onefile --windowed --icon=NONE --name=MonitorSwitcher monitor_switcher.py
   ```

3. **Find the executable**:
   - Look in the `dist` folder
   - Run `MonitorSwitcher.exe`

## Troubleshooting

### Monitor doesn't switch

1. **Check DDC/CI support**:
   - Verify your monitor supports DDC/CI
   - Check monitor OSD menu for DDC/CI setting and enable it

2. **Try different monitor index**:
   - If you have multiple monitors, try changing `monitor_index` in config
   - Set to 0, 1, 2, etc.

3. **Check input names**:
   - Verify the input names match your monitor's actual inputs
   - Try different options (HDMI-1 vs HDMI-2)

4. **Check logs**:
   - Look at `monitor_switcher.log` for error messages
   - Run from command line to see real-time output

### Application doesn't detect KM switch

1. **USB connection**:
   - Verify your KM switch is connected via USB
   - The detection works by monitoring USB device changes

2. **Adjust check interval**:
   - Edit `config.json` and increase `check_interval` to 3.0 or 5.0
   - Some systems may need more time to detect device changes

3. **Manual switching**:
   - Use the manual "Switch to Home" / "Switch to Work" options from the tray menu
   - This bypasses automatic detection

### Icon doesn't appear in system tray

- On Windows 11, you may need to enable the icon in system tray settings:
  - Settings → Personalization → Taskbar → System tray icons
  - Find "Python" or "MonitorSwitcher" and enable it

## Technical Details

### DDC/CI Control

The application uses the `monitorcontrol` library to send VCP (Virtual Control Panel) commands to your monitor via DDC/CI. This is the same protocol your monitor's physical buttons use.

### Device Detection

The app monitors USB device presence using Windows APIs (`pywin32`). When keyboard/mouse devices disappear (you switched away) or appear (you switched to this machine), it triggers the monitor switch.

### Supported Input Sources

- HDMI-1, HDMI-2
- DisplayPort-1, DisplayPort-2
- DVI-1, DVI-2
- VGA-1

## License

This project is open source and available for personal and commercial use.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## Support

If you encounter issues:
1. Check the `monitor_switcher.log` file for error messages
2. Verify your monitor supports DDC/CI
3. Try manual switching first to isolate the issue
4. Open an issue on GitHub with log details

---

**Note**: This application is designed to run on your home Windows 11 machine. It monitors when you switch TO and FROM your home machine, then automatically changes your monitor input accordingly.
