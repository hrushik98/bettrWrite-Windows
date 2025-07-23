#!/usr/bin/env python3
"""
bettrWrite for Windows
A tool to process selected text using AI models via keyboard shortcuts
"""

import json
import time
import sys
import os
from pathlib import Path
from typing import Dict, Any, Optional
import logging
from datetime import datetime

import keyboard
import pyperclip
import pyautogui
import requests
from plyer import notification

# Configure logging
LOG_DIR = Path.home() / "bettrwrite_logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "bettrwrite.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration paths
CONFIG_DIR = Path.home() / ".config" / "bettrwrite"
CONFIG_FILE = CONFIG_DIR / "config.json"

# Fail-safe to prevent infinite loops
pyautogui.FAILSAFE = True


class BettrWriteError(Exception):
    """Custom exception for BettrWrite errors"""
    pass


class BettrWrite:
    def __init__(self):
        self.config = {}
        self.original_clipboard = ""
        self.load_config()
        self.setup_hotkeys()
        
    def load_config(self):
        """Load configuration from JSON file"""
        if not CONFIG_FILE.exists():
            logger.error(f"Configuration file not found at {CONFIG_FILE}")
            self.show_notification("Error", "Configuration file not found. Please run the installer.")
            sys.exit(1)
            
        try:
            with open(CONFIG_FILE, 'r') as f:
                self.config = json.load(f)
            logger.info("Configuration loaded successfully")
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            self.show_notification("Error", f"Failed to load config: {str(e)}")
            sys.exit(1)
    
    def show_notification(self, title: str, message: str, timeout: int = 5):
        """Show Windows notification"""
        try:
            notification.notify(
                title=f"bettrWrite - {title}",
                message=message,
                timeout=timeout,
                app_name="bettrWrite"
            )
        except Exception as e:
            logger.error(f"Failed to show notification: {e}")
    
    def save_clipboard(self):
        """Save current clipboard content"""
        try:
            self.original_clipboard = pyperclip.paste()
        except Exception:
            self.original_clipboard = ""
    
    def restore_clipboard(self):
        """Restore original clipboard content"""
        try:
            if self.original_clipboard:
                pyperclip.copy(self.original_clipboard)
        except Exception as e:
            logger.error(f"Failed to restore clipboard: {e}")
    
    def get_selected_text(self) -> Optional[str]:
        """Get currently selected text via clipboard"""
        try:
            # Save current clipboard
            self.save_clipboard()
            
            # Clear clipboard to detect if copy worked
            pyperclip.copy("")
            
            # Simulate Ctrl+C to copy selected text
            pyautogui.hotkey('ctrl', 'c')
            
            # Wait for clipboard to update
            time.sleep(0.1)
            
            # Get the selected text
            selected_text = pyperclip.paste()
            
            # If clipboard is still empty, no text was selected
            if not selected_text:
                self.restore_clipboard()
                return None
                
            return selected_text
            
        except Exception as e:
            logger.error(f"Failed to get selected text: {e}")
            self.restore_clipboard()
            return None
    
    def replace_selected_text(self, new_text: str):
        """Replace selected text with new text"""
        try:
            # Copy new text to clipboard
            pyperclip.copy(new_text)
            
            # Small delay to ensure clipboard is ready
            time.sleep(0.05)
            
            # Paste the new text
            pyautogui.hotkey('ctrl', 'v')
            
            # Restore original clipboard after a delay
            time.sleep(0.1)
            self.restore_clipboard()
            
        except Exception as e:
            logger.error(f"Failed to replace text: {e}")
            self.restore_clipboard()
            raise BettrWriteError(f"Failed to replace text: {str(e)}")
    
    def call_openai_api(self, text: str, config: Dict[str, Any]) -> str:
        """Call OpenAI API"""
        api_key = os.getenv("OPENAI_API_KEY") or self.config.get("settings", {}).get("openai_api_key")
        
        if not api_key or api_key == "YOUR_OPENAI_API_KEY_OR_NULL":
            raise BettrWriteError("OpenAI API key not configured")
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        payload = {
            "model": config.get("model", "gpt-4o"),
            "messages": [
                {"role": "system", "content": config.get("prompt", "")},
                {"role": "user", "content": text}
            ],
            **config.get("openai_options", {})
        }
        
        try:
            response = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            
            result = response.json()
            return result["choices"][0]["message"]["content"].strip()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"OpenAI API request failed: {e}")
            if hasattr(e, 'response') and e.response:
                logger.error(f"Response: {e.response.text}")
            raise BettrWriteError(f"OpenAI API error: {str(e)}")
    
    def call_ollama_api(self, text: str, config: Dict[str, Any]) -> str:
        """Call Ollama API"""
        base_url = self.config.get("settings", {}).get("ollama_base_url", "http://localhost:11434")
        api_url = f"{base_url}/api/generate"
        
        # Combine system prompt and user text
        full_prompt = f"{config.get('prompt', '')}\n\nText to process:\n{text}"
        
        payload = {
            "model": config.get("model", ""),
            "prompt": full_prompt,
            "stream": False,
            **config.get("ollama_options", {})
        }
        
        try:
            response = requests.post(api_url, json=payload, timeout=60)
            response.raise_for_status()
            
            result = response.json()
            if "error" in result:
                raise BettrWriteError(f"Ollama error: {result['error']}")
                
            return result.get("response", "").strip()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Ollama API request failed: {e}")
            raise BettrWriteError(f"Ollama API error: {str(e)}")
    
    def process_text(self, shortcut_id: str):
        """Main text processing function"""
        logger.info(f"Processing text for shortcut: {shortcut_id}")
        
        # Find shortcut configuration
        shortcut_config = None
        for shortcut in self.config.get("shortcuts", []):
            if shortcut.get("id") == shortcut_id:
                shortcut_config = shortcut
                break
        
        if not shortcut_config:
            logger.error(f"Shortcut configuration not found: {shortcut_id}")
            self.show_notification("Error", f"Shortcut '{shortcut_id}' not configured")
            return
        
        try:
            # Get selected text
            selected_text = self.get_selected_text()
            if not selected_text:
                self.show_notification("Info", "No text selected")
                return
            
            logger.info(f"Selected text length: {len(selected_text)}")
            
            # Show processing notification
            self.show_notification("Processing", "Processing text...", timeout=2)
            
            # Call appropriate API
            backend = shortcut_config.get("backend", "openai")
            
            if backend == "openai":
                processed_text = self.call_openai_api(selected_text, shortcut_config)
            elif backend == "ollama":
                processed_text = self.call_ollama_api(selected_text, shortcut_config)
            else:
                raise BettrWriteError(f"Unknown backend: {backend}")
            
            # Replace the text
            self.replace_selected_text(processed_text)
            
            # Show success notification
            self.show_notification("Success", "Text processed and replaced")
            logger.info("Text processing completed successfully")
            
        except BettrWriteError as e:
            logger.error(f"Processing error: {e}")
            self.show_notification("Error", str(e))
        except Exception as e:
            logger.error(f"Unexpected error: {e}", exc_info=True)
            self.show_notification("Error", f"Unexpected error: {str(e)}")
    
    def setup_hotkeys(self):
        """Set up keyboard shortcuts"""
        shortcuts = self.config.get("shortcuts", [])
        
        if not shortcuts:
            logger.error("No shortcuts configured")
            self.show_notification("Error", "No shortcuts configured")
            return
        
        for shortcut in shortcuts:
            hotkey = shortcut.get("keys", "").replace(" ", "")
            shortcut_id = shortcut.get("id")
            
            try:
                # Register the hotkey
                keyboard.add_hotkey(
                    hotkey,
                    lambda sid=shortcut_id: self.process_text(sid),
                    suppress=False
                )
                logger.info(f"Registered hotkey: {hotkey} -> {shortcut_id}")
                
            except Exception as e:
                logger.error(f"Failed to register hotkey {hotkey}: {e}")
                self.show_notification("Error", f"Failed to register hotkey: {hotkey}")
    
    def run(self):
        """Run the application"""
        logger.info("BettrWrite started")
        self.show_notification("Started", "BettrWrite is running. Press Ctrl+Q to quit.")
        
        try:
            # Add quit hotkey
            keyboard.add_hotkey('ctrl+q', self.quit)
            
            # Keep the script running
            keyboard.wait()
            
        except KeyboardInterrupt:
            self.quit()
    
    def quit(self):
        """Quit the application"""
        logger.info("BettrWrite shutting down")
        self.show_notification("Stopped", "BettrWrite stopped")
        os._exit(0)


def main():
    """Main entry point"""
    try:
        app = BettrWrite()
        app.run()
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()