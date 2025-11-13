import os
import shutil
import winreg
from pathlib import Path

def install_to_startup():
    """Install the script to run automatically"""
    
    print("=" * 60)
    print("  Wallpaper Engine Manager - Installation")
    print("=" * 60)
    print()
    
    # Get current directory
    current_dir = Path(__file__).parent.absolute()
    
    # Target directory (in AppData)
    appdata = os.getenv('APPDATA')
    target_dir = Path(appdata) / "WallpaperEngineManager"
    
    # Create directory if it doesn't exist
    target_dir.mkdir(parents=True, exist_ok=True)
    print(f"✓ Created directory: {target_dir}")
    
    # Copy files
    files_to_copy = [
        "wallpaper_engine_manager.py",
        "start_wallpaper_manager.bat"
    ]
    
    for file in files_to_copy:
        src = current_dir / file
        dst = target_dir / file
        if src.exists():
            shutil.copy2(src, dst)
            print(f"✓ Copied: {file}")
        else:
            print(f"✗ Not found: {file}")
            return False
    
    # Add to Startup Registry
    bat_path = str(target_dir / "start_wallpaper_manager.bat")
    
    try:
        # Open Registry Startup key
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE
        )
        
        # Add value
        winreg.SetValueEx(
            key,
            "WallpaperEngineManager",
            0,
            winreg.REG_SZ,
            bat_path
        )
        
        winreg.CloseKey(key)
        print(f"✓ Added to Windows Startup")
        print()
        print("=" * 60)
        print("✅ Installation completed successfully!")
        print()
        print("The script will run automatically on every system boot.")
        print(f"File location: {target_dir}")
        print()
        print("To uninstall:")
        print("1. Run 'uninstall.py'")
        print("2. Or remove manually from Task Manager > Startup")
        print("=" * 60)
        
        # Ask if run now
        print()
        response = input("Start the monitor now? (y/n): ").lower()
        if response in ['y', 'yes']:
            os.startfile(bat_path)
            print("✓ Monitor started!")
        
        return True
        
    except Exception as e:
        print(f"✗ Error adding to Startup: {e}")
        return False


def create_uninstaller():
    """Create uninstall script"""
    appdata = os.getenv('APPDATA')
    target_dir = Path(appdata) / "WallpaperEngineManager"
    
    uninstall_script = '''import winreg
import os
import shutil
from pathlib import Path

print("=" * 60)
print("  Uninstalling Wallpaper Engine Manager")
print("=" * 60)
print()

try:
    # Remove from Registry
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0,
        winreg.KEY_ALL_ACCESS
    )
    
    try:
        winreg.DeleteValue(key, "WallpaperEngineManager")
        print("✓ Removed from Windows Startup")
    except FileNotFoundError:
        print("⚠ Not found in Startup")
    
    winreg.CloseKey(key)
    
    # Stop processes
    import psutil
    for proc in psutil.process_iter(['name', 'cmdline']):
        try:
            if 'wallpaper_engine_manager.py' in ' '.join(proc.info['cmdline'] or []):
                proc.terminate()
                print("✓ Process stopped")
        except:
            pass
    
    print()
    print("✅ Uninstallation complete!")
    print()
    print("You can manually delete the folder:")
    print(f"   {Path(os.getenv('APPDATA')) / 'WallpaperEngineManager'}")
    
except Exception as e:
    print(f"✗ Error: {e}")

input("\\nPress Enter to close...")
'''
    
    uninstall_path = target_dir / "uninstall.py"
    with open(uninstall_path, 'w', encoding='utf-8') as f:
        f.write(uninstall_script)
    
    print(f"✓ Created uninstall script: {uninstall_path}")


if __name__ == "__main__":
    try:
        # Check required packages
        print("Checking required packages...")
        required_packages = ['psutil', 'pywin32']
        missing_packages = []
        
        for package in required_packages:
            try:
                __import__(package if package != 'pywin32' else 'win32api')
            except ImportError:
                missing_packages.append(package)
        
        if missing_packages:
            print(f"\n⚠ Missing packages: {', '.join(missing_packages)}")
            print("\nInstall them using:")
            print(f"   pip install {' '.join(missing_packages)}")
            print()
            input("Press Enter after installation...")
        
        # Continue with installation
        if install_to_startup():
            create_uninstaller()
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
    
    input("\nPress Enter to close...")
```    """Install the script to run automatically"""