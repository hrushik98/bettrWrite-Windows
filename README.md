# bettrWrite for Windows

A Windows port of the bettrWrite text processing tool that lets you correct, summarize, or transform selected text using AI models with configurable keyboard shortcuts.

1.  Open your terminal (Command Prompt or PowerShell).
2.  Navigate to the directory where `bettrwrite.bat` is located (e.g., using `cd
C:\Users\hrushik\.config\bettrwrite`).
3.  Then, simply type `bettrwrite.bat` and press Enter.

## Features

- **Global Hotkeys**: Process text from any Windows application
- **Multiple AI Backends**: OpenAI (GPT-4) and local Ollama support
- **Customizable Shortcuts**: Define multiple shortcuts for different tasks
- **System Tray Notifications**: Visual feedback for all operations
- **Clipboard Safety**: Preserves your original clipboard content
- **Easy Configuration**: JSON-based configuration with installer

## Requirements

- Windows 10/11
- Python 3.7 or higher
- Administrator privileges (recommended for some applications)

## Installation

1. **Download the files**:
   - `bettrwrite_windows.py` (main application)
   - `install_windows.py` (installer)

2. **Run the installer**:
   ```bash
   python install_windows.py
   ```

3. **Follow the prompts** to:
   - Install Python dependencies automatically
   - Configure OpenAI and/or Ollama backends
   - Set up your first keyboard shortcut
   - Create desktop shortcut

## Usage

### Starting bettrWrite

- **Option 1**: Double-click the desktop shortcut
- **Option 2**: Run `~/.config/bettrwrite/bettrwrite.bat`
- **Option 3**: Run directly: `python ~/.config/bettrwrite/bettrwrite_windows.py`

### Using bettrWrite

1. Select text in any application
2. Press your configured hotkey (default: `Ctrl+E`)
3. Wait for the notification
4. The text is automatically replaced with the processed version

### Stopping bettrWrite

Press `Ctrl+Q` while bettrWrite is running

## Configuration

Configuration is stored in `~/.config/bettrwrite/config.json`

### Example Configuration

```json
{
  "settings": {
    "openai_api_key": "sk-...",
    "ollama_base_url": "http://localhost:11434"
  },
  "shortcuts": [
    {
      "id": "grammar_correct",
      "keys": "ctrl+e",
      "backend": "openai",
      "model": "gpt-4o",
      "prompt": "Fix grammar and spelling errors only. Output ONLY the corrected text.",
      "openai_options": {
        "temperature": 0.3
      }
    },
    {
      "id": "summarize",
      "keys": "ctrl+shift+s",
      "backend": "ollama",
      "model": "llama3.2:latest",
      "prompt": "Summarize this text in one sentence. Output ONLY the summary.",
      "ollama_options": {
        "temperature": 0.5
      }
    }
  ]
}
```

### Adding New Shortcuts

1. Edit `~/.config/bettrwrite/config.json`
2. Add a new entry to the `shortcuts` array
3. Restart bettrWrite

### Shortcut Format

- `keys`: Use format like `"ctrl+e"`, `"ctrl+shift+s"`, `"alt+g"`
- Available modifiers: `ctrl`, `alt`, `shift`, `win`
- Keys: Any letter, number, or F1-F12

## Differences from Linux Version

### What's Different:
- Uses Python instead of bash scripting
- Native Windows notifications instead of notify-send
- Windows clipboard API instead of xclip
- Python's `keyboard` module instead of xbindkeys
- Batch/VBS files for launching instead of shell scripts

### What's the Same:
- Core functionality and workflow
- Configuration file format
- API backend support (OpenAI/Ollama)
- Customizable shortcuts and prompts

## Troubleshooting

### Common Issues

1. **"Access Denied" errors**:
   - Run as Administrator for elevated applications
   - Some apps (like some games) block automation

2. **Hotkey not working**:
   - Check if another application is using the same hotkey
   - Ensure bettrWrite is running (check system tray)
   - Try running as Administrator

3. **Text not being replaced**:
   - Some applications don't support standard clipboard operations
   - Try with a different application first (like Notepad)
   - Check logs at `~/bettrwrite_logs/bettrwrite.log`

4. **Antivirus warnings**:
   - Keyboard monitoring may trigger warnings
   - Add bettrWrite to your antivirus exceptions

### Debug Mode

Check logs at: `C:\Users\YourName\bettrwrite_logs\bettrwrite.log`

## Advanced Usage

### Environment Variables

Instead of storing your OpenAI key in the config file:
```bash
set OPENAI_API_KEY=sk-your-key-here
```

### Custom Ollama Models

1. Install Ollama from https://ollama.com
2. Pull models: `ollama pull llama3.2`
3. Configure in bettrWrite

### Running on Startup

1. Press `Win+R`, type `shell:startup`
2. Copy the bettrWrite shortcut there
3. It will start automatically on login

## Security Notes

- API keys are stored in plain text in the config file
- Consider using environment variables for sensitive keys
- The tool requires keyboard monitoring permissions
- Only processes text when you explicitly trigger it

## Limitations

- Cannot interact with some protected applications
- Windows Secure Desktop (UAC prompts) is not accessible
- Some modern UWP apps may not work properly
- Gaming anti-cheat systems may block functionality

## Uninstallation

1. Delete the configuration folder: `~/.config/bettrwrite`
2. Delete the desktop shortcut
3. Remove from startup folder if added
4. Uninstall Python packages if desired:
   ```bash
   pip uninstall keyboard pyperclip pyautogui plyer
   ```

## Contributing

Feel free to submit issues or pull requests for:
- Additional backend support
- New features
- Bug fixes
- Documentation improvements

## License

Same as the original bettrWrite project
