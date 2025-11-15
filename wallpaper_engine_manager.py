import os
import subprocess
import time
import win32api
import win32con
import win32gui
import psutil
import logging
import threading
from pathlib import Path
from datetime import datetime

# Setup logging
logging.basicConfig(
    filename='wallpaper_engine_manager.log',
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    encoding='utf-8'
)

# Settings
CHECK_INTERVAL = 10                    # Check every 10 seconds
RESTART_DELAY = 3                      # Wait 3 seconds before restarting
LONG_SLEEP_THRESHOLD = 2.5 * 60 * 60  # 2.5 hours in seconds
SHORT_SLEEP_DELAY = 5                  # Wait 5 seconds for short sleeps
LONG_SLEEP_DELAY = 30                  # Wait 30 seconds for long sleeps (2.5+ hours)

class WallpaperEngineManager:
    def __init__(self):
        # Possible Wallpaper Engine paths
        self.wallpaper_paths = [
            r"D:\SteamLibrary\steamapps\common\wallpaper_engine\wallpaper32.exe",
            r"D:\SteamLibrary\steamapps\common\wallpaper_engine\wallpaper64.exe",
            r"C:\Program Files (x86)\Steam\steamapps\common\wallpaper_engine\wallpaper32.exe",
            r"C:\Program Files (x86)\Steam\steamapps\common\wallpaper_engine\wallpaper64.exe",
            r"C:\Program Files\Steam\steamapps\common\wallpaper_engine\wallpaper32.exe",
            r"C:\Program Files\Steam\steamapps\common\wallpaper_engine\wallpaper64.exe",
        ]
        self.wallpaper_exe = self.find_wallpaper_engine()
        self.process_name = "wallpaper32.exe" if "wallpaper32" in self.wallpaper_exe else "wallpaper64.exe"
        self.is_sleeping = False      # Sleep status
        self.monitoring_active = True # Is monitor active
        self.monitor_thread = None    # Monitor thread
        self.sleep_start_time = None  # Track when sleep started
        
    def find_wallpaper_engine(self):
        """Find Wallpaper Engine path"""
        for path in self.wallpaper_paths:
            if os.path.exists(path):
                logging.info(f"Found Wallpaper Engine at: {path}")
                return path
        
        # If not found, try searching in disk drives
        logging.warning("Wallpaper Engine not found in common paths, searching...")
        for drive in ['C:', 'D:', 'E:', 'F:']:
            search_paths = [
                f"{drive}\\SteamLibrary\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
                f"{drive}\\Steam\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
                f"{drive}\\Program Files\\Steam\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
            ]
            for path in search_paths:
                if os.path.exists(path):
                    logging.info(f"Found Wallpaper Engine at: {path}")
                    return path
        
        raise FileNotFoundError("Wallpaper Engine not found! Make sure it's installed.")
    
    def is_running(self):
        """Check if Wallpaper Engine is running"""
        for proc in psutil.process_iter(['name']):
            try:
                if self.process_name.lower() in proc.info['name'].lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return False
    
    def start_wallpaper_engine(self):
        """Start Wallpaper Engine"""
        if not self.is_running():
            try:
                subprocess.Popen(self.wallpaper_exe, shell=False)
                logging.info("Wallpaper Engine started successfully")
                print("✓ Wallpaper Engine started")
            except Exception as e:
                logging.error(f"Error starting Wallpaper Engine: {e}")
                print(f"✗ Error: {e}")
        else:
            logging.info("Wallpaper Engine already running")
    
    def stop_wallpaper_engine(self):
        """Stop Wallpaper Engine"""
        for proc in psutil.process_iter(['name', 'pid']):
            try:
                if self.process_name.lower() in proc.info['name'].lower():
                    proc.terminate()
                    proc.wait(timeout=5)
                    logging.info("Wallpaper Engine stopped")
                    print("✓ Wallpaper Engine stopped")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                pass
    
    def continuous_monitor(self):
        """Continuous monitor - checks if Wallpaper Engine is running"""
        logging.info("Continuous monitor started")
        print(f"🔄 Continuous monitor active - checking every {CHECK_INTERVAL} seconds")
        
        while self.monitoring_active:
            try:
                # If computer is not sleeping and Wallpaper Engine is not running - start it
                if not self.is_sleeping and not self.is_running():
                    logging.warning("Wallpaper Engine not running! Restarting...")
                    print(f"⚠️  Wallpaper Engine closed - restarting...")
                    time.sleep(RESTART_DELAY)
                    self.start_wallpaper_engine()
                
                time.sleep(CHECK_INTERVAL)
                
            except Exception as e:
                logging.error(f"Error in continuous monitor: {e}")
                time.sleep(CHECK_INTERVAL)
    
    def start_monitoring(self):
        """Start continuous monitor in separate thread"""
        if self.monitor_thread is None or not self.monitor_thread.is_alive():
            self.monitoring_active = True
            self.monitor_thread = threading.Thread(target=self.continuous_monitor, daemon=True)
            self.monitor_thread.start()
    
    def stop_monitoring(self):
        """Stop continuous monitor"""
        self.monitoring_active = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=2)


class PowerEventMonitor:
    def __init__(self, manager):
        self.manager = manager
        self.hwnd = None
        
    def create_window(self):
        """Create hidden window to receive system messages"""
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self.wnd_proc
        wc.lpszClassName = "WallpaperEngineMonitor"
        wc.hInstance = win32api.GetModuleHandle(None)
        
        try:
            class_atom = win32gui.RegisterClass(wc)
            self.hwnd = win32gui.CreateWindow(
                class_atom,
                "Wallpaper Engine Monitor",
                0,
                0, 0, 0, 0,
                0,
                0,
                wc.hInstance,
                None
            )
        except Exception as e:
            logging.error(f"Error creating window: {e}")
            raise
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        """Handle system messages"""
        if msg == win32con.WM_POWERBROADCAST:
            if wparam == win32con.PBT_APMRESUMEAUTOMATIC:
                # Computer woke from sleep
                wake_time = datetime.now()
                
                if self.manager.sleep_start_time:
                    # Calculate sleep duration
                    sleep_duration = (wake_time - self.manager.sleep_start_time).total_seconds()
                    hours = sleep_duration / 3600
                    
                    logging.info(f"Computer woke from sleep - slept for {hours:.1f} hours")
                    
                    # Decide delay based on sleep duration
                    if sleep_duration >= LONG_SLEEP_THRESHOLD:
                        # Long sleep (2.5+ hours) - wait longer to prevent flickering
                        delay = LONG_SLEEP_DELAY
                        print(f"⚡ Computer woke up (slept {hours:.1f}h) - waiting {delay}s for display...")
                        logging.info(f"Long sleep detected ({hours:.1f}h) - using {delay}s delay")
                    else:
                        # Short sleep (<2.5 hours) - quick resume
                        delay = SHORT_SLEEP_DELAY
                        print(f"⚡ Computer woke up (slept {hours:.1f}h) - waiting {delay}s...")
                        logging.info(f"Short sleep detected ({hours:.1f}h) - using {delay}s delay")
                else:
                    # No sleep start time recorded - use default long delay for safety
                    delay = LONG_SLEEP_DELAY
                    print(f"⚡ Computer woke up - waiting {delay}s for display...")
                    logging.info("Wake from sleep (unknown duration) - using default delay")
                
                # Wait based on calculated delay
                time.sleep(delay)
                
                self.manager.is_sleeping = False
                self.manager.sleep_start_time = None  # Reset
                print("✓ System ready - continuous monitor will start Wallpaper Engine")
                # Continuous monitor will handle starting after this delay
                
            elif wparam == win32con.PBT_APMSUSPEND:
                # Computer going to sleep
                self.manager.sleep_start_time = datetime.now()
                logging.info("Computer going to sleep")
                print("💤 Computer sleeping - stopping Wallpaper Engine...")
                self.manager.is_sleeping = True
                self.manager.stop_wallpaper_engine()
                
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def run(self):
        """Run main loop"""
        self.create_window()
        logging.info("Monitor started")
        print("🚀 Monitor running! Script waiting for system events...")
        print("   💡 Continuous monitor enabled - Wallpaper Engine will stay on!")
        print(f"   🔧 Smart sleep mode:")
        print(f"      • Short sleep (<2.5h): {SHORT_SLEEP_DELAY}s delay")
        print(f"      • Long sleep (2.5h+): {LONG_SLEEP_DELAY}s delay (prevents flickering)")
        print("   (Press Ctrl+C to stop)")
        
        # Start Wallpaper Engine initially (on boot)
        self.manager.start_wallpaper_engine()
        
        # Start continuous monitor
        self.manager.start_monitoring()
        
        # Message loop
        try:
            win32gui.PumpMessages()
        except KeyboardInterrupt:
            logging.info("Monitor stopped by user")
            print("\n⏹️  Stopping continuous monitor...")
            self.manager.stop_monitoring()
            print("👋 Monitor stopped")


if __name__ == "__main__":
    print("=" * 50)
    print("  Wallpaper Engine Manager - Smart Edition")
    print("=" * 50)
    
    try:
        manager = WallpaperEngineManager()
        monitor = PowerEventMonitor(manager)
        monitor.run()
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
        logging.error(str(e))
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        logging.error(f"Error: {e}")
