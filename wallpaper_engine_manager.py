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

# הגדרת לוג
logging.basicConfig(
    filename='wallpaper_engine_manager.log',
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    encoding='utf-8'
)

# הגדרות
CHECK_INTERVAL = 10  # בדיקה כל 10 שניות
RESTART_DELAY = 3    # המתנה של 3 שניות לפני הפעלה מחדש

class WallpaperEngineManager:
    def __init__(self):
        # נתיבים אפשריים של Wallpaper Engine
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
        self.is_sleeping = False  # סטטוס שינה
        self.monitoring_active = True  # האם המוניטור פעיל
        self.monitor_thread = None  # ה-thread של המוניטור
        
    def find_wallpaper_engine(self):
        """מוצא את הנתיב של Wallpaper Engine"""
        for path in self.wallpaper_paths:
            if os.path.exists(path):
                logging.info(f"נמצא Wallpaper Engine ב: {path}")
                return path
        
        # אם לא נמצא, נסה לחפש בכונני הדיסק
        logging.warning("לא נמצא Wallpaper Engine בנתיבים הרגילים, מחפש...")
        for drive in ['C:', 'D:', 'E:']:
            search_paths = [
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
                f"{drive}\\Program Files\\Steam\\steamapps\\common\\wallpaper_engine\\wallpaper32.exe",
            ]
            for path in search_paths:
                if os.path.exists(path):
                    logging.info(f"נמצא Wallpaper Engine ב: {path}")
                    return path
        
        raise FileNotFoundError("לא נמצא Wallpaper Engine! וודא שהתוכנה מותקנת.")
    
    def is_running(self):
        """בודק אם Wallpaper Engine רץ"""
        for proc in psutil.process_iter(['name']):
            try:
                if self.process_name.lower() in proc.info['name'].lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return False
    
    def start_wallpaper_engine(self):
        """מפעיל את Wallpaper Engine"""
        if not self.is_running():
            try:
                subprocess.Popen(self.wallpaper_exe, shell=False)
                logging.info("Wallpaper Engine הופעל בהצלחה")
                print("✓ Wallpaper Engine הופעל")
            except Exception as e:
                logging.error(f"שגיאה בהפעלת Wallpaper Engine: {e}")
                print(f"✗ שגיאה: {e}")
        else:
            logging.info("Wallpaper Engine כבר רץ")
    
    def stop_wallpaper_engine(self):
        """סגירת Wallpaper Engine"""
        for proc in psutil.process_iter(['name', 'pid']):
            try:
                if self.process_name.lower() in proc.info['name'].lower():
                    proc.terminate()
                    proc.wait(timeout=5)
                    logging.info("Wallpaper Engine נסגר")
                    print("✓ Wallpaper Engine נסגר")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                pass
    
    def continuous_monitor(self):
        """מוניטור רציף - בודק כל הזמן אם Wallpaper Engine רץ"""
        logging.info("מוניטור רציף התחיל")
        print("🔄 מוניטור רציף פעיל - בודק כל {} שניות".format(CHECK_INTERVAL))
        
        while self.monitoring_active:
            try:
                # אם המחשב לא בשינה וWallpaper Engine לא רץ - הפעל אותו
                if not self.is_sleeping and not self.is_running():
                    logging.warning("Wallpaper Engine לא רץ! מפעיל מחדש...")
                    print(f"⚠️  Wallpaper Engine נסגר - מפעיל מחדש...")
                    time.sleep(RESTART_DELAY)
                    self.start_wallpaper_engine()
                
                time.sleep(CHECK_INTERVAL)
                
            except Exception as e:
                logging.error(f"שגיאה במוניטור הרציף: {e}")
                time.sleep(CHECK_INTERVAL)
    
    def start_monitoring(self):
        """מתחיל את המוניטור הרציף ב-thread נפרד"""
        if self.monitor_thread is None or not self.monitor_thread.is_alive():
            self.monitoring_active = True
            self.monitor_thread = threading.Thread(target=self.continuous_monitor, daemon=True)
            self.monitor_thread.start()
    
    def stop_monitoring(self):
        """עוצר את המוניטור הרציף"""
        self.monitoring_active = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=2)


class PowerEventMonitor:
    def __init__(self, manager):
        self.manager = manager
        self.hwnd = None
        
    def create_window(self):
        """יוצר חלון נסתר לקבלת הודעות מערכת"""
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
            logging.error(f"שגיאה ביצירת חלון: {e}")
            raise
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        """מטפל בהודעות מערכת"""
        if msg == win32con.WM_POWERBROADCAST:
            if wparam == win32con.PBT_APMRESUMEAUTOMATIC:
                # המחשב התעורר משינה
                logging.info("המחשב התעורר משינה")
                print("⚡ המחשב התעורר - מפעיל Wallpaper Engine...")
                self.manager.is_sleeping = False
                time.sleep(2)  # המתן קצר לאחר התעוררות
                self.manager.start_wallpaper_engine()
                
            elif wparam == win32con.PBT_APMSUSPEND:
                # המחשב נכנס לשינה
                logging.info("המחשב נכנס לשינה")
                print("💤 המחשב נכנס לשינה - סוגר Wallpaper Engine...")
                self.manager.is_sleeping = True
                self.manager.stop_wallpaper_engine()
                
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def run(self):
        """מריץ את הלולאה הראשית"""
        self.create_window()
        logging.info("המוניטור התחיל לרוץ")
        print("🚀 המוניטור פועל! הסקריפט ממתין לאירועי מערכת...")
        print("   💡 מוניטור רציף מופעל - Wallpaper Engine יהיה דלוק תמיד!")
        print("   (לחץ Ctrl+C לעצירה)")
        
        # הפעל Wallpaper Engine בהפעלה ראשונית
        self.manager.start_wallpaper_engine()
        
        # התחל מוניטור רציף
        self.manager.start_monitoring()
        
        # לולאת הודעות
        try:
            win32gui.PumpMessages()
        except KeyboardInterrupt:
            logging.info("המוניטור הופסק על ידי המשתמש")
            print("\n⏹️  עוצר מוניטור רציף...")
            self.manager.stop_monitoring()
            print("👋 המוניטור נעצר")


if __name__ == "__main__":
    print("=" * 50)
    print("  Wallpaper Engine Manager")
    print("=" * 50)
    
    try:
        manager = WallpaperEngineManager()
        monitor = PowerEventMonitor(manager)
        monitor.run()
    except FileNotFoundError as e:
        print(f"\n❌ שגיאה: {e}")
        logging.error(str(e))
    except Exception as e:
        print(f"\n❌ שגיאה לא צפויה: {e}")
        logging.error(f"שגיאה: {e}")
