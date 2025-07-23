#!/usr/bin/env python3
"""
BettrWrite Windows Installer
Sets up bettrWrite with configuration and dependencies
"""

import json
import sys
import os
import subprocess
from pathlib import Path
import getpass


class Colors:
    """Console colors for Windows"""
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BLUE = '\033[94m'
    END = '\033[0m'


def print_colored(text, color=Colors.END):
    """Print colored text"""
    print(f"{color}{text}{Colors.END}")


def install_dependencies():
    """Install required Python packages"""
    print_colored("\nInstalling Python dependencies...", Colors.YELLOW)
    
    required_packages = [
        "keyboard",
        "pyperclip",
        "pyautogui",
        "requests",
        "plyer"
    ]
    
    for package in required_packages:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print_colored(f"✓ {package} installed successfully", Colors.GREEN)
        except subprocess.CalledProcessError:
            print_colored(f"✗ Failed to install {package}", Colors.RED)
            return False
    
    # Optional: Install pywin32 for desktop shortcuts
    print_colored("\nOptional: Install pywin32 for desktop shortcuts?", Colors.YELLOW)
    install_pywin32 = input("Install pywin32? (y/N): ").lower()
    if install_pywin32 == 'y':
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            print_colored("✓ pywin32 installed successfully", Colors.GREEN)
        except subprocess.CalledProcessError:
            print_colored("✗ Failed to install pywin32 - shortcuts will be created without icons", Colors.YELLOW)
    
    return True


def create_config_directory():
    """Create configuration directory"""
    config_dir = Path.home() / ".config" / "bettrwrite"
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def configure_openai():
    """Configure OpenAI backend"""
    print_colored("\nConfiguring OpenAI backend...", Colors.BLUE)
    
    api_key = ""
    use_env = input("Do you want to use OPENAI_API_KEY environment variable? (y/N): ").lower()
    
    if use_env == 'y':
        env_key = os.getenv("OPENAI_API_KEY")
        if env_key:
            print_colored("✓ Using OPENAI_API_KEY from environment", Colors.GREEN)
            api_key = "USE_ENV_VARIABLE"
        else:
            print_colored("OPENAI_API_KEY not found in environment", Colors.YELLOW)
    
    if not api_key:
        api_key = getpass.getpass("Enter your OpenAI API key (hidden): ").strip()
        if not api_key:
            print_colored("No API key provided. You'll need to set it later.", Colors.YELLOW)
            api_key = "YOUR_OPENAI_API_KEY_OR_NULL"
    
    return api_key


def configure_ollama():
    """Configure Ollama backend"""
    print_colored("\nConfiguring Ollama backend...", Colors.BLUE)
    
    base_url = input("Enter Ollama server URL (default: http://localhost:11434): ").strip()
    if not base_url:
        base_url = "http://localhost:11434"
    
    # Test connection
    print("Testing Ollama connection...")
    try:
        import requests
        response = requests.get(f"{base_url}/api/tags", timeout=5)
        if response.status_code == 200:
            print_colored("✓ Ollama server is accessible", Colors.GREEN)
            
            # Show available models
            models = response.json().get("models", [])
            if models:
                print("\nAvailable models:")
                for model in models:
                    print(f"  - {model['name']}")
            else:
                print_colored("No models found. Please pull models using 'ollama pull <model>'", Colors.YELLOW)
        else:
            print_colored("✗ Ollama server returned error", Colors.RED)
    except Exception as e:
        print_colored(f"✗ Could not connect to Ollama: {str(e)}", Colors.RED)
    
    model = input("\nEnter Ollama model name (e.g., llama3.2:latest): ").strip()
    if not model:
        print_colored("No model specified. You'll need to configure it later.", Colors.YELLOW)
        model = ""
    
    return base_url, model


def create_configuration(config_dir):
    """Create configuration file"""
    print_colored("\nCreating configuration...", Colors.YELLOW)
    
    config = {
        "settings": {
            "openai_api_key": "YOUR_OPENAI_API_KEY_OR_NULL",
            "ollama_base_url": "http://localhost:11434"
        },
        "shortcuts": []
    }
    
    # Choose backends
    use_openai = input("\nConfigure OpenAI backend? (Y/n): ").lower() != 'n'
    use_ollama = input("Configure Ollama backend? (Y/n): ").lower() != 'n'
    
    if not use_openai and not use_ollama:
        print_colored("At least one backend must be configured!", Colors.RED)
        return None
    
    # Configure chosen backends
    if use_openai:
        api_key = configure_openai()
        if api_key != "USE_ENV_VARIABLE":
            config["settings"]["openai_api_key"] = api_key
    
    ollama_model = ""
    if use_ollama:
        base_url, ollama_model = configure_ollama()
        config["settings"]["ollama_base_url"] = base_url
    
    # Configure default shortcut
    print_colored("\nConfiguring default shortcut...", Colors.BLUE)
    
    # Choose backend for default shortcut
    if use_openai and use_ollama:
        backend_choice = input("Use OpenAI or Ollama for default shortcut? (openai/ollama) [ollama]: ").lower()
        default_backend = "openai" if backend_choice == "openai" else "ollama"
    elif use_openai:
        default_backend = "openai"
    else:
        default_backend = "ollama"
    
    default_model = "gpt-4o" if default_backend == "openai" else ollama_model
    
    # Shortcut customization
    default_keys = "ctrl+e"
    custom_keys = input(f"Enter shortcut keys (default: {default_keys}): ").strip()
    if custom_keys:
        default_keys = custom_keys
    
    default_prompt = (
        "You are a helpful assistant that corrects grammar and spelling errors "
        "in the provided text without changing its meaning or tone. "
        "Only fix errors - don't rewrite the text. Output ONLY the corrected text."
    )
    
    print("\nDefault prompt:")
    print(default_prompt)
    custom_prompt = input("\nEnter custom prompt (or press Enter to use default): ").strip()
    if custom_prompt:
        default_prompt = custom_prompt
    
    # Add default shortcut
    config["shortcuts"].append({
        "id": "grammar_correct",
        "keys": default_keys,
        "backend": default_backend,
        "model": default_model,
        "prompt": default_prompt,
        "ollama_options": {"temperature": 0.3},
        "openai_options": {"temperature": 0.3}
    })
    
    # Save configuration
    config_file = config_dir / "config.json"
    with open(config_file, 'w') as f:
        json.dump(config, f, indent=2)
    
    print_colored(f"\n✓ Configuration saved to {config_file}", Colors.GREEN)
    return config_file


def create_batch_file(config_dir):
    """Create batch file for easy launching"""
    batch_content = f"""@echo off
cd /d "{config_dir}"
python "{config_dir / 'bettrwrite_windows.py'}"
pause
"""
    
    batch_file = config_dir / "bettrwrite.bat"
    with open(batch_file, 'w') as f:
        f.write(batch_content)
    
    # Also create a VBS file for silent launch
    vbs_content = f'''Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "{batch_file}" & Chr(34), 0
Set WshShell = Nothing
'''
    
    vbs_file = config_dir / "bettrwrite_silent.vbs"
    with open(vbs_file, 'w') as f:
        f.write(vbs_content)
    
    return batch_file, vbs_file


def create_desktop_shortcut(vbs_file):
    """Create desktop shortcut"""
    try:
        import win32com.client
        
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        
        shortcut = shell.CreateShortcut(os.path.join(desktop, "bettrWrite.lnk"))
        shortcut.TargetPath = str(vbs_file)
        shortcut.WorkingDirectory = str(vbs_file.parent)
        shortcut.IconLocation = "notepad.exe"
        shortcut.Description = "bettrWrite - AI Text Processing"
        shortcut.save()
        
        return True
    except ImportError:
        print_colored("pywin32 not installed. Skipping desktop shortcut.", Colors.YELLOW)
        print("Install with: pip install pywin32")
        return False
    except Exception as e:
        print_colored(f"Failed to create desktop shortcut: {e}", Colors.YELLOW)
        return False


def main():
    """Main installer function"""
    print_colored("\n=== bettrWrite Windows Installer ===", Colors.GREEN)
    print("This will install bettrWrite with AI text processing capabilities.")
    
    # Check Python version
    if sys.version_info < (3, 7):
        print_colored("Python 3.7 or higher is required!", Colors.RED)
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        print_colored("\nFailed to install dependencies. Please install manually:", Colors.RED)
        print("pip install keyboard pyperclip pyautogui requests plyer")
        sys.exit(1)
    
    # Create configuration
    config_dir = create_config_directory()
    config_file = create_configuration(config_dir)
    
    if not config_file:
        print_colored("Configuration failed!", Colors.RED)
        sys.exit(1)
    
    # Copy main script
    print_colored("\nInstalling main script...", Colors.YELLOW)
    main_script = Path("bettrwrite_windows.py")
    if main_script.exists():
        import shutil
        shutil.copy(main_script, config_dir / "bettrwrite_windows.py")
        print_colored("✓ Main script installed", Colors.GREEN)
    else:
        print_colored("✗ bettrwrite_windows.py not found in current directory!", Colors.RED)
        print("Please ensure bettrwrite_windows.py is in the same folder as this installer.")
        sys.exit(1)
    
    # Create launch files
    batch_file, vbs_file = create_batch_file(config_dir)
    print_colored(f"✓ Launch files created", Colors.GREEN)
    
    # Create desktop shortcut
    if create_desktop_shortcut(vbs_file):
        print_colored("✓ Desktop shortcut created", Colors.GREEN)
    
    # Final instructions
    print_colored("\n=== Installation Complete! ===", Colors.GREEN)
    print("\nTo run bettrWrite:")
    print(f"1. Double-click the desktop shortcut, OR")
    print(f"2. Run: {batch_file}")
    print(f"\nConfiguration file: {config_file}")
    
    # Load config to show the default shortcut
    try:
        with open(config_file, 'r') as f:
            saved_config = json.load(f)
            default_shortcut = saved_config.get('shortcuts', [{}])[0].get('keys', 'ctrl+e')
            print(f"\nDefault shortcut: {default_shortcut}")
    except:
        print("\nDefault shortcut: ctrl+e")
    
    print("\nPress Ctrl+Q to quit bettrWrite when running.")
    
    print_colored("\nIMPORTANT:", Colors.YELLOW)
    print("- Run bettrWrite as Administrator if you need to process text in elevated applications")
    print("- Some antivirus software may flag keyboard monitoring - this is normal")
    print("- Check logs at: ~/bettrwrite_logs/bettrwrite.log if you encounter issues")
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()