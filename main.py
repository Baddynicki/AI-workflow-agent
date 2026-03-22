# main.py
import os
os.add_dll_directory(r'C:\Windows\System32')
os.add_dll_directory(r'C:\Windows\SysWOW64')

import json
import sys
import threading

from PySide6.QtWidgets import QApplication

from listener.keyboard_listener import KeyboardListener
from listener.clipboard_listener import ClipboardListener
from parser.command_parser import CommandParser
from ai.openai_handler import OpenAIHandler
from executor.excel_executor import ExcelExecutor
from executor.word_executor import WordExecutor
from executor.ppt_executor import PPTExecutor
from utils.app_launcher import AppLauncher
from utils.command_buffer import CommandBuffer
from utils.logger import setup_logger
from ui.main_window import MainWindow


def normalize_action_payload(action):
    """
    Normalizes various openai response shapes into one clean action dict.
    Handles:
      {"action": "create_table", "rows": 5}
      {"action": "create_table", "parameters": {"rows": 5}}
      {"action": "create_table {\"rows\": 5}"}
    """
    if not isinstance(action, dict):
        return {}

    normalized = dict(action)

    # Flatten nested "parameters" key
    params = normalized.get("parameters")
    if isinstance(params, dict):
        normalized.update(params)

    # Handle embedded JSON in action string
    raw_action = normalized.get("action")
    if isinstance(raw_action, str):
        action_text = raw_action.strip()
        if "{" in action_text and action_text.endswith("}"):
            name, _, json_part = action_text.partition("{")
            normalized["action"] = name.strip()
            try:
                embedded = json.loads("{" + json_part)
                if isinstance(embedded, dict):
                    normalized.update(embedded)
            except json.JSONDecodeError:
                pass
        else:
            normalized["action"] = action_text

    return normalized


def extract_actions(payload):
    """
    Extracts a list of executable action dicts from openai's response.
    Handles single action, list of actions, or container keys
    like "actions", "commands", "steps".
    """
    if isinstance(payload, list):
        return [normalize_action_payload(a) for a in payload if a]

    if isinstance(payload, dict):
        # Single action
        if "action" in payload:
            return [normalize_action_payload(payload)]
        # Container key
        for key in ["actions", "commands", "steps"]:
            if key in payload and isinstance(payload[key], list):
                return [normalize_action_payload(a) for a in payload[key]]

    return []


def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # Keep agent alive when window is closed

    # ── Core modules ──────────────────────────────────────────────────────
    parser   = CommandParser()
    openai   = OpenAIHandler()
    launcher = AppLauncher()
    cmd_buf  = CommandBuffer()

    executors = {
        "excel":      ExcelExecutor(launcher),
        "word":       WordExecutor(launcher),
        "powerpoint": PPTExecutor(launcher),
        "ppt":        PPTExecutor(launcher),
    }

    # ── UI ────────────────────────────────────────────────────────────────
    # Window is created before logger so the logger can attach to it
    window = MainWindow(on_command_callback=lambda text: on_command_detected(text))
    logger = setup_logger(ui_callback=window.log)

    logger.info("✅ Office Agent started")
    logger.info("Type or paste  agent: <app>: <command>  then press Enter")

    # ── Command handler ───────────────────────────────────────────────────
    def on_command_detected(raw_text):
        # Step 1: Parse trigger
        parsed = parser.parse(raw_text)
        if not parsed:
            logger.warning(f"Could not parse: {raw_text}")
            window.record_command(raw_text, "❌")
            return

        app_name = parsed["app"]
        command  = parsed["command"]
        logger.info(f"App: {app_name} | Command: {command}")

        # Step 2: Get executor
        executor = executors.get(app_name)
        if not executor:
            logger.warning(f"Unknown app: {app_name}")
            window.record_command(f"{app_name}: {command}", "❌")
            return

        # Step 3: Ask openai for structured action
        action_payload = openai.interpret(app_name, command)
        if not action_payload:
            logger.error("openai returned no action")
            window.record_command(f"{app_name}: {command}", "❌")
            return

        # Step 4: Extract and normalize actions
        actions = extract_actions(action_payload)
        if not actions:
            logger.error(f"No executable actions in: {action_payload}")
            window.record_command(f"{app_name}: {command}", "❌")
            return

        # Step 5: Execute each action
        success = True
        for action in actions:
            try:
                executor.execute(action)
            except Exception as e:
                logger.error(f"Execution failed: {e}")
                success = False

        window.record_command(f"{app_name}: {command}", "✅" if success else "❌")

    # ── Background listeners ──────────────────────────────────────────────
    # Clipboard: stores agent: commands as candidates (never executes directly)
    clipboard_listener = ClipboardListener(cmd_buf)
    threading.Thread(
        target=clipboard_listener.start,
        daemon=True,
        name="ClipboardListener"
    ).start()
    logger.info("📋 Clipboard listener started")

    # Keyboard: executes on Enter key only
    kb_listener = KeyboardListener(on_command_detected, cmd_buf)
    threading.Thread(
        target=kb_listener.start,
        daemon=True,
        name="KeyboardListener"
    ).start()
    logger.info("⌨️  Keyboard listener started")

    # ── Launch UI ─────────────────────────────────────────────────────────
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
