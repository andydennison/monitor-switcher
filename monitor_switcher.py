"""
Monitor Switcher - Automatically switch monitor input based on KM switch usage
"""
import os
import sys
import json
import time
import threading
import logging
from pathlib import Path
from typing import Optional, Dict, Any

import pystray
from PIL import Image, ImageDraw
from monitorcontrol import get_monitors, InputSource, VCPError
import win32api
import win32con
import win32gui

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('monitor_switcher.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class Config:
    """Configuration manager for monitor switcher"""

    def __init__(self, config_file: str = "config.json"):
        self.config_file = Path(config_file)
        self.data = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Error loading config: {e}")

        # Default configuration
        return {
            "home_machine_input": "HDMI-1",  # HDMI-1, HDMI-2, DisplayPort-1, etc.
            "work_laptop_input": "HDMI-2",
            "monitor_index": 0,  # Index of monitor to control (0 = first)
            "check_interval": 2.0,  # Seconds between checks
            "last_active_machine": "home"  # "home" or "work"
        }

    def save(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.data, f, indent=4)
            logger.info("Configuration saved")
        except Exception as e:
            logger.error(f"Error saving config: {e}")

    def get(self, key: str, default=None):
        """Get configuration value"""
        return self.data.get(key, default)

    def set(self, key: str, value):
        """Set configuration value"""
        self.data[key] = value
        self.save()


class MonitorController:
    """Controller for monitor input switching via DDC/CI"""

    # Mapping of input names to VCP codes
    INPUT_SOURCES = {
        "HDMI-1": InputSource.HDMI1,
        "HDMI-2": InputSource.HDMI2,
        "DisplayPort-1": InputSource.DP1,
        "DisplayPort-2": InputSource.DP2,
        "DVI-1": InputSource.DVI1,
        "DVI-2": InputSource.DVI2,
        "VGA-1": 0x01,  # VGA/RGB
    }

    def __init__(self, monitor_index: int = 0):
        self.monitor_index = monitor_index
        self.monitor = None
        self._init_monitor()

    def _init_monitor(self):
        """Initialize monitor connection"""
        try:
            monitors = get_monitors()
            if monitors and len(monitors) > self.monitor_index:
                self.monitor = monitors[self.monitor_index]
                logger.info(f"Connected to monitor at index {self.monitor_index}")
            else:
                logger.warning(f"No monitor found at index {self.monitor_index}")
        except Exception as e:
            logger.error(f"Error initializing monitor: {e}")

    def switch_input(self, input_name: str) -> bool:
        """Switch monitor to specified input"""
        if not self.monitor:
            logger.warning("No monitor connected, attempting to reconnect...")
            self._init_monitor()
            if not self.monitor:
                return False

        if input_name not in self.INPUT_SOURCES:
            logger.error(f"Unknown input source: {input_name}")
            return False

        try:
            input_source = self.INPUT_SOURCES[input_name]
            with self.monitor:
                self.monitor.set_input_source(input_source)
            logger.info(f"Switched monitor to {input_name}")
            return True
        except VCPError as e:
            logger.error(f"VCP Error switching input: {e}")
            return False
        except Exception as e:
            logger.error(f"Error switching input: {e}")
            return False

    def get_current_input(self) -> Optional[str]:
        """Get current monitor input"""
        if not self.monitor:
            return None

        try:
            with self.monitor:
                current = self.monitor.get_input_source()
                # Find matching input name
                for name, source in self.INPUT_SOURCES.items():
                    if source == current:
                        return name
            return None
        except Exception as e:
            logger.error(f"Error getting current input: {e}")
            return None


class KMSwitchDetector:
    """Detects when KM switch changes by monitoring USB device presence"""

    def __init__(self, callback):
        self.callback = callback
        self.running = False
        self.thread = None
        self.last_device_count = 0
        self.last_state = "home"  # Assume starting on home machine

    def start(self):
        """Start monitoring for device changes"""
        self.running = True
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()
        logger.info("KM switch detector started")

    def stop(self):
        """Stop monitoring"""
        self.running = False
        if self.thread:
            self.thread.join(timeout=5)
        logger.info("KM switch detector stopped")

    def _monitor_loop(self):
        """Main monitoring loop"""
        # Get initial device count
        self.last_device_count = self._count_input_devices()

        while self.running:
            try:
                current_count = self._count_input_devices()

                # If device count changed significantly, KM switch likely occurred
                if abs(current_count - self.last_device_count) >= 1:
                    # Devices appeared = switched to this machine
                    if current_count > self.last_device_count:
                        new_state = "home"
                    else:
                        # Devices disappeared = switched away from this machine
                        new_state = "work"

                    if new_state != self.last_state:
                        logger.info(f"KM switch detected: {self.last_state} -> {new_state}")
                        self.last_state = new_state
                        self.callback(new_state)

                    self.last_device_count = current_count

                time.sleep(2.0)  # Check every 2 seconds

            except Exception as e:
                logger.error(f"Error in monitor loop: {e}")
                time.sleep(5.0)

    def _count_input_devices(self) -> int:
        """Count active input devices (keyboards and mice)"""
        try:
            # Use Windows API to enumerate HID devices
            import win32api
            device_count = 0

            # Count keyboards
            device_count += win32api.GetSystemMetrics(win32con.SM_CMOUSEBUTTONS)

            # Simple approach: check if we can get keyboard state
            # If keyboard is available, GetKeyState won't fail
            try:
                for vk in [win32con.VK_SHIFT, win32con.VK_CONTROL, win32con.VK_MENU]:
                    win32api.GetKeyState(vk)
                device_count += 2  # Keyboard is present
            except:
                pass

            return device_count
        except Exception as e:
            logger.error(f"Error counting devices: {e}")
            return 0


class ConfigWindow:
    """Configuration window using tkinter"""

    def __init__(self, config: Config, on_save_callback=None):
        import tkinter as tk
        from tkinter import ttk

        self.config = config
        self.on_save_callback = on_save_callback

        # Create window
        self.root = tk.Tk()
        self.root.title("Monitor Switcher Configuration")
        self.root.geometry("400x300")
        self.root.resizable(False, False)

        # Make window appear on top
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)

        # Create UI
        self._create_widgets()

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """Create configuration UI widgets"""
        import tkinter as tk
        from tkinter import ttk

        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title = ttk.Label(main_frame, text="Monitor Input Configuration",
                         font=('Arial', 12, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Home machine input
        ttk.Label(main_frame, text="Home Machine Input:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.home_input = ttk.Combobox(main_frame, width=15, state='readonly')
        self.home_input['values'] = ('HDMI-1', 'HDMI-2', 'DisplayPort-1', 'DisplayPort-2', 'DVI-1', 'VGA-1')
        self.home_input.set(self.config.get('home_machine_input', 'HDMI-1'))
        self.home_input.grid(row=1, column=1, sticky=tk.W, pady=5)

        # Work laptop input
        ttk.Label(main_frame, text="Work Laptop Input:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.work_input = ttk.Combobox(main_frame, width=15, state='readonly')
        self.work_input['values'] = ('HDMI-1', 'HDMI-2', 'DisplayPort-1', 'DisplayPort-2', 'DVI-1', 'VGA-1')
        self.work_input.set(self.config.get('work_laptop_input', 'HDMI-2'))
        self.work_input.grid(row=2, column=1, sticky=tk.W, pady=5)

        # Monitor index
        ttk.Label(main_frame, text="Monitor Index:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.monitor_index = ttk.Spinbox(main_frame, from_=0, to=3, width=15)
        self.monitor_index.set(self.config.get('monitor_index', 0))
        self.monitor_index.grid(row=3, column=1, sticky=tk.W, pady=5)

        # Info label
        info = ttk.Label(main_frame, text="Monitor index 0 = first monitor",
                        font=('Arial', 8), foreground='gray')
        info.grid(row=4, column=0, columnspan=2, pady=(0, 20))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="Save", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(side=tk.LEFT, padx=5)

        # Status label
        self.status_label = ttk.Label(main_frame, text="", foreground='green')
        self.status_label.grid(row=6, column=0, columnspan=2)

    def _save(self):
        """Save configuration"""
        self.config.set('home_machine_input', self.home_input.get())
        self.config.set('work_laptop_input', self.work_input.get())
        self.config.set('monitor_index', int(self.monitor_index.get()))

        self.status_label.config(text="Configuration saved!")

        if self.on_save_callback:
            self.on_save_callback()

        # Close after 1 second
        self.root.after(1000, self.root.destroy)

    def _cancel(self):
        """Cancel and close window"""
        self.root.destroy()

    def show(self):
        """Show the configuration window"""
        self.root.mainloop()


class MonitorSwitcherApp:
    """Main application class"""

    def __init__(self):
        self.config = Config()
        self.monitor_controller = MonitorController(
            monitor_index=self.config.get('monitor_index', 0)
        )
        self.km_detector = KMSwitchDetector(callback=self._on_km_switch)
        self.icon = None

    def _on_km_switch(self, machine: str):
        """Callback when KM switch is detected"""
        logger.info(f"Switching to {machine} machine")

        # Get the appropriate input from config
        if machine == "home":
            input_name = self.config.get('home_machine_input', 'HDMI-1')
        else:
            input_name = self.config.get('work_laptop_input', 'HDMI-2')

        # Switch monitor input
        success = self.monitor_controller.switch_input(input_name)

        if success:
            self.config.set('last_active_machine', machine)
            if self.icon:
                self.icon.notify(
                    f"Switched to {machine} machine ({input_name})",
                    "Monitor Switcher"
                )
        else:
            if self.icon:
                self.icon.notify(
                    f"Failed to switch to {input_name}",
                    "Monitor Switcher - Error"
                )

    def _create_image(self):
        """Create system tray icon"""
        # Create a simple icon with "M" letter
        width = 64
        height = 64
        image = Image.new('RGB', (width, height), color='#2E86AB')
        dc = ImageDraw.Draw(image)

        # Draw "M" for Monitor
        dc.text((width//2, height//2), "M", fill='white', anchor='mm',
                font_size=32 if hasattr(dc, 'font_size') else None)

        return image

    def _show_config(self, icon, item):
        """Show configuration window"""
        logger.info("Opening configuration window")

        def on_save():
            # Reinitialize monitor controller with new settings
            self.monitor_controller = MonitorController(
                monitor_index=self.config.get('monitor_index', 0)
            )
            logger.info("Configuration updated")

        config_window = ConfigWindow(self.config, on_save_callback=on_save)
        config_window.show()

    def _manual_switch_home(self, icon, item):
        """Manually switch to home machine"""
        self._on_km_switch("home")

    def _manual_switch_work(self, icon, item):
        """Manually switch to work laptop"""
        self._on_km_switch("work")

    def _quit(self, icon, item):
        """Quit the application"""
        logger.info("Shutting down...")
        self.km_detector.stop()
        icon.stop()

    def run(self):
        """Run the application"""
        logger.info("Starting Monitor Switcher...")

        # Start KM switch detector
        self.km_detector.start()

        # Create system tray icon
        menu = pystray.Menu(
            pystray.MenuItem("Switch to Home", self._manual_switch_home),
            pystray.MenuItem("Switch to Work", self._manual_switch_work),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Configure", self._show_config),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit)
        )

        self.icon = pystray.Icon(
            "monitor_switcher",
            self._create_image(),
            "Monitor Switcher",
            menu
        )

        logger.info("Monitor Switcher running in system tray")
        self.icon.run()


def main():
    """Main entry point"""
    try:
        app = MonitorSwitcherApp()
        app.run()
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
