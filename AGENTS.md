# Repository Guidelines

## Project Structure & Module Organization
- `addon/manifest.ini` holds NVDA add-on metadata (name, version, min NVDA version).
- `addon/globalPlugins/accessMenu/__init__.py` contains the global plugin implementation with three dialog classes (AccessMenuDialog, AppsMenuDialog, PowerMenuDialog) and a settings panel.
- `build.sh` packages the add-on into `accessMenu.nvda-addon` at the repo root.

## Build, Test, and Development Commands
- `./build.sh`
  - Packages the add-on by zipping the contents of `addon/` into `accessMenu.nvda-addon`.
  - Example output: `accessMenu.nvda-addon` in the repository root.
- There are no automated tests or dev servers in this repository yet.

## Coding Style & Naming Conventions
- Language: Python (NVDA add-on API).
- Indentation: 4 spaces, no tabs.
- Naming:
  - Classes: `PascalCase` (e.g., `AccessMenuSettingsPanel`).
  - Functions/methods: `snake_case` (e.g., `_populate_power_menu`).
  - Constants: `UPPER_SNAKE_CASE` (e.g., `APP_EXTENSIONS`).
- Keep UI strings short and user-facing labels configurable where practical.

## Testing Guidelines
- No formal test framework configured.
- Manual checks recommended:
  - Install the built add-on in NVDA and restart NVDA.
  - Verify Input Gestures shows **Access Menu** > **Open the Access Menu**.
  - Press NVDA+Shift+M to open the main Access Menu dialog.
  - Test navigation with Up/Down arrows through Apps (136 items) and Power (3 items) dialogs.
  - Validate dialog announcements, app launching, and power confirmations.
  - Check that Escape key closes dialogs properly.

## Commit & Pull Request Guidelines
- No Git history or conventions are present in this repository.
- If adding versioned changes, update `addon/manifest.ini` `version` and rebuild.
- Suggested commit format (if you initialize Git): `type: short summary` (e.g., `feat: add search dialog`).

## Configuration & Deployment Notes
- NVDA loads add-ons from `C:\Users\<User>\AppData\Roaming\NVDA\addons`.
- After rebuilding, reinstall the `.nvda-addon` and restart NVDA to pick up changes.
- Add-on uses dialog-based UI (wx.Dialog with wx.ListCtrl) for screen reader accessibility.
- Dialog announcements are handled via ui.message() with wx.CallAfter() for proper timing.
- Gesture: NVDA+Shift+M opens the main Access Menu dialog.
