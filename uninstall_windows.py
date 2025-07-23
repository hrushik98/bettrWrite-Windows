#!/usr/bin/env python3
"""
BettrWrite Windows Uninstaller
Removes bettrWrite and its configuration
"""

import os
import shutil
from pathlib import Path


class Colors:
    """Console colors for Windows"""
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    END = '\033[0m'


def print_colored(text, color=Colors.END):
    """Print colored text"""
    print(f"{color}{text}{Colors.END}")


def remove_directory(path, description):
    """Remove a directory if it exists"""
    if path.exists():
        try:
            shutil.rmtree(path)
            print_colored(f"✓ Removed {description}: {path}", Colors.GREEN)
            return True
        except Exception as e:
            print_colored(f"✗ Failed to remove {description}: {e}", Colors.RED)
            return False
    else:
        print_colored(f"- {description} not found: {path}", Colors.YELLOW)
        return True


def remove_file(path, description):
    """Remove a file if it exists"""
    if path.exists():
        try:
            os.remove(path)
            print_colored(f"✓ Removed {description}: {path}", Colors.GREEN)
            return True
        except Exception as e:
            print_colored(f"✗ Failed to remove {description}: {e}", Colors.RED)
            return False
    else:
        print_colored(f"- {description} not found: {path}", Colors.YELLOW)
        return True


def main():
    """Main uninstaller function"""
    print_colored("\n=== bettrWrite Windows Uninstaller ===", Colors.RED)
    
    confirm = input("\nThis will remove bettrWrite and all its configuration. Continue? (y/N): ")
    if confirm.lower() != 'y':
        print("Uninstallation cancelled.")
        return
    
    print_colored("\nRemoving bettrWrite...", Colors.YELLOW)
    
    # Configuration directory
    config_dir = Path.home() / ".config" / "bettrwrite"
    remove_directory(config_dir, "configuration directory")
    
    # Log directory
    log_dir = Path.home() / "bettrwrite_logs"
    remove_directory(log_dir, "log directory")
    
    # Desktop shortcut
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        shortcut_path = Path(desktop) / "bettrWrite.lnk"
        remove_file(shortcut_path, "desktop shortcut")
    except:
        print_colored("- Could not check for desktop shortcut", Colors.YELLOW)
    
    # Startup shortcut
    startup_dir = Path(os.environ['APPDATA']) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    startup_shortcut = startup_dir / "bettrWrite.lnk"
    remove_file(startup_shortcut, "startup shortcut")
    
    print_colored("\n=== Uninstallation Complete ===", Colors.GREEN)
    
    print("\nPython packages were not removed. To remove them manually, run:")
    print("pip uninstall keyboard pyperclip pyautogui plyer requests")
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()