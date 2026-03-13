---
name: windows-desktop-bridge
description: Build, configure, or use a Windows desktop automation bridge for local GUI control. Use when the user wants desktop-level automation for Windows apps such as PyCharm, wants the agent to click/type/send hotkeys/focus windows/take screenshots, or wants an MCP-like local control service for GUI workflows.
---

# Windows Desktop Bridge

Create or use a local Windows desktop automation bridge for GUI tasks.

## Quick workflow

1. Prefer a local HTTP bridge that exposes a small set of actions: health, screenshot, windows, activate, launch, type, hotkey, click.
2. Keep the first version dependency-light. Prefer PowerShell/Win32/standard library over large frameworks.
3. Store bridge code in `scripts/` and keep operational notes in `references/`.
4. Test the bridge immediately after creation with a health check and at least one visible side effect.
5. For fragile GUI actions, require an explicit target window title or executable path.

## Default implementation

Use `scripts/desktop_bridge.py` as the local HTTP server.

Capabilities of the minimum viable bridge:
- `GET /health`
- `GET /windows`
- `POST /launch`
- `POST /activate`
- `POST /hotkey`
- `POST /type`
- `POST /click`
- `POST /screenshot`

## Safety

1. Keep the bind address on `127.0.0.1` unless the user explicitly wants remote exposure.
2. Prefer explicit window matching over blind coordinate clicks.
3. Announce destructive or high-risk GUI actions before doing them.

## References

Read `references/usage.md` when you need local run and test commands.
