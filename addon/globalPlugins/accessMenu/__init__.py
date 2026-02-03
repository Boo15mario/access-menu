# NVDA add-on: Access Menu
# Provides a start menu replacement with Apps and Power menus.

import os
import subprocess
import sys
import tempfile
import uuid
import threading

import wx
import addonHandler
try:
    import comtypes
    import comtypes.client as comtypes_client
except Exception:
    comtypes = None
    comtypes_client = None
import gui
import ui
import config
import globalPluginHandler
from gui import guiHelper
from gui.settingsDialogs import NVDASettingsDialog, SettingsPanel
from scriptHandler import script
from logHandler import log

try:
    addonHandler.initTranslation()
except addonHandler.AddonError:
    log.warning("Unable to init translations. This may be because the addon is running from NVDA scratchpad.")

try:
    curAddon = addonHandler.getCodeAddon()
    ADDON_SUMMARY = curAddon.manifest.get("summary", "Access Menu")
except:
    ADDON_SUMMARY = "Access Menu"

def _(text):
    """Fallback translation function"""
    return text

try:
    _ = addonHandler.getTranslationFunction()
except:
    pass

log.info("Access Menu add-on: module imported")


APP_EXTENSIONS = {".lnk", ".url", ".appref-ms"}

GOD_MODE_CLSID = "ED7BA470-8E54-465E-825C-99712043E01C"
GOD_MODE_SHELL = f"shell:::{GOD_MODE_CLSID}"
_HELPER_EXE_MISSING_LOGGED = False

CONFIG_SECTION = "accessMenu"
CONFIG_SPEC = {
    "appsLabel": "string(default='Apps')",
    "powerLabel": "string(default='Power')",
    "settingsLabel": "string(default='Settings')",
    "folderPrefix": "string(default='Folder: ')",
    "searchLabel": "string(default='Search...')",
    "allAppsLabel": "string(default='All Apps (A-Z)')",
    "categoriesLabel": "string(default='Categories')",
    "browseFoldersLabel": "string(default='Browse Folders')",
    "signOutLabel": "string(default='Sign out')",
    "powerOffLabel": "string(default='Power off')",
    "rebootLabel": "string(default='Reboot')",
    "confirmTitle": "string(default='Confirm')",
    "confirmSignOut": "string(default='Sign out of Windows?')",
    "confirmPowerOff": "string(default='Power off the PC?')",
    "confirmReboot": "string(default='Restart the PC?')",
    "searchDialogTitle": "string(default='Search Apps')",
    "searchHint": "string(default='Type to filter apps')",
}


def _unique_name(name, existing):
    if name not in existing:
        return name
    i = 2
    while f"{name} ({i})" in existing:
        i += 1
    return f"{name} ({i})"


def _start_menu_roots():
    roots = []
    program_data = os.environ.get("PROGRAMDATA")
    app_data = os.environ.get("APPDATA")
    if program_data:
        roots.append(os.path.join(program_data, "Microsoft", "Windows", "Start Menu", "Programs"))
    if app_data:
        roots.append(os.path.join(app_data, "Microsoft", "Windows", "Start Menu", "Programs"))
    return [r for r in roots if os.path.isdir(r)]


def _build_tree():
    tree = {}
    for root in _start_menu_roots():
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            parts = [] if rel == "." else rel.split(os.sep)
            node = tree
            for part in parts:
                node = node.setdefault(part, {})
            for filename in filenames:
                base, ext = os.path.splitext(filename)
                if ext.lower() not in APP_EXTENSIONS:
                    continue
                path = os.path.join(dirpath, filename)
                name = _unique_name(base, node)
                node[name] = path
    return tree


def _sorted_items(mapping):
    # mapping: name -> dict (folder) or path (file)
    return sorted(mapping.items(), key=lambda kv: kv[0].casefold())


def _flatten_apps(tree, prefix=""):
    apps = []
    for name, value in _sorted_items(tree):
        if isinstance(value, dict):
            next_prefix = f"{prefix}{name}\\"
            apps.extend(_flatten_apps(value, next_prefix))
        else:
            display = name
            if prefix:
                stripped_prefix = prefix.rstrip("\\")
                display = f"{name} ({stripped_prefix})"
            apps.append((display, value))
    return apps


def _top_level_categories(tree):
    # Return list of (categoryName, categoryTree)
    return [(name, value) for name, value in _sorted_items(tree) if isinstance(value, dict)]


def _ensure_config():
    if CONFIG_SECTION not in config.conf.spec:
        config.conf.spec[CONFIG_SECTION] = CONFIG_SPEC


def _get_cfg(key):
    return config.conf[CONFIG_SECTION].get(key, CONFIG_SPEC[key].split("default='", 1)[1].split("'")[0])


def _open_god_mode_folder():
    try:
        subprocess.Popen(["explorer.exe", GOD_MODE_SHELL])
    except OSError:
        try:
            os.startfile(GOD_MODE_SHELL)
        except OSError:
            subprocess.Popen(["explorer.exe"])


def _get_powershell_candidates():
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    candidates = []
    if sys.maxsize <= 2**32:
        candidates.append(os.path.join(system_root, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe"))
    candidates.append(os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"))
    candidates.append("powershell")
    return candidates


def _get_helper_exe():
    bin_dir = os.path.join(os.path.dirname(__file__), "bin")
    exe_path = os.path.join(bin_dir, "godmode_helper.exe")
    if os.path.isfile(exe_path):
        return exe_path
    return None


def _get_cscript_candidates():
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    candidates = []
    if sys.maxsize <= 2**32:
        candidates.append(os.path.join(system_root, "Sysnative", "cscript.exe"))
    candidates.append(os.path.join(system_root, "System32", "cscript.exe"))
    candidates.append("cscript")
    return candidates


def _get_god_mode_items_via_powershell():
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_file:
            temp_path = temp_file.name
    except Exception:
        log.exception("Access Menu: failed to create temp file for PowerShell output.")
        return []

    escaped_path = temp_path.replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject Shell.Application
$folder = $shell.NameSpace('shell:::{GOD_MODE_CLSID}')
if ($null -eq $folder) {{
    [System.IO.File]::WriteAllText('{escaped_path}', '', (New-Object System.Text.UTF8Encoding($false)))
    exit
}}
$items = @()
foreach ($item in $folder.Items()) {{
    if ($null -ne $item -and $item.Name) {{
        $items += $item.Name
    }}
}}
[System.IO.File]::WriteAllLines('{escaped_path}', $items, (New-Object System.Text.UTF8Encoding($false)))
"""
    result = None
    ps_exe = None
    try:
        for candidate in _get_powershell_candidates():
            try:
                ps_exe = candidate
                result = subprocess.run(
                    [
                        candidate,
                        "-NoLogo",
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-STA",
                        "-Command",
                        script,
                    ],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
                break
            except FileNotFoundError:
                continue
        if result is None:
            log.warning("Access Menu: PowerShell executable not found.")
            return []
    except Exception:
        log.exception("Access Menu: failed to run PowerShell for God Mode enumeration.")
        return []

    if result.returncode != 0:
        log.warning("Access Menu: PowerShell God Mode enumeration failed: %s", result.stderr.strip())
        return []

    payload = ""
    try:
        if temp_path and os.path.exists(temp_path):
            with open(temp_path, "r", encoding="utf-8", errors="replace") as handle:
                payload = handle.read()
    except Exception:
        log.exception("Access Menu: failed to read PowerShell output file.")
        payload = ""
    finally:
        try:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass

    if not payload.strip():
        log.warning(
            "Access Menu: PowerShell God Mode enumeration returned empty output. exe=%s stderr=%s",
            ps_exe,
            (result.stderr or "").strip(),
        )
        return []
    results = []
    for line in payload.splitlines():
        name = line.strip()
        if name:
            results.append((name, name))
    results.sort(key=lambda kv: kv[0].casefold())
    return results


def _get_god_mode_items_via_helper():
    global _HELPER_EXE_MISSING_LOGGED
    exe_path = _get_helper_exe()
    if not exe_path:
        if not _HELPER_EXE_MISSING_LOGGED:
            log.warning("Access Menu: helper exe not found in addon bin directory.")
            _HELPER_EXE_MISSING_LOGGED = True
        return []
    try:
        result = subprocess.run(
            [exe_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except OSError as exc:
        log.warning("Access Menu: helper exe failed to launch: %s", exc)
        return []

    if result.returncode != 0:
        log.warning("Access Menu: helper exe failed: %s", (result.stderr or "").strip())
        return []

    lines = []
    for line in (result.stdout or "").splitlines():
        name = line.strip()
        if name:
            lines.append(name)
    if not lines:
        log.warning("Access Menu: helper exe returned no items.")
        return []

    # Deduplicate while preserving order.
    seen = set()
    results = []
    for name in lines:
        key = name.casefold()
        if key in seen:
            continue
        seen.add(key)
        results.append((name, name))
    results.sort(key=lambda kv: kv[0].casefold())
    log.info("Access Menu: helper exe returned %d items.", len(results))
    return results


def _get_god_mode_items_via_cscript():
    temp_path = None
    script_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_file:
            temp_path = temp_file.name
        script_name = f"godmode-{uuid.uuid4().hex}.vbs"
        script_path = os.path.join(tempfile.gettempdir(), script_name)
        escaped_path = temp_path.replace('"', '""')
        script = (
            'On Error Resume Next\n'
            'Dim shell, folder, items, fso, outFile\n'
            'Set shell = CreateObject("Shell.Application")\n'
            f'Set folder = shell.NameSpace("shell:::{GOD_MODE_CLSID}")\n'
            'If folder Is Nothing Then WScript.Quit 0\n'
            'Set items = folder.Items()\n'
            'Set fso = CreateObject("Scripting.FileSystemObject")\n'
            f'Set outFile = fso.CreateTextFile("{escaped_path}", True, True)\n'
            'For Each item In items\n'
            '  If Not item Is Nothing Then\n'
            '    If Len(item.Name) > 0 Then outFile.WriteLine item.Name\n'
            '  End If\n'
            'Next\n'
            'outFile.Close\n'
        )
        with open(script_path, "w", encoding="utf-8") as handle:
            handle.write(script)

        result = None
        for candidate in _get_cscript_candidates():
            try:
                result = subprocess.run(
                    [candidate, "//NoLogo", script_path],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
                break
            except FileNotFoundError:
                continue
        if result is None:
            log.warning("Access Menu: cscript executable not found.")
            return []
        if result.returncode != 0:
            log.warning("Access Menu: cscript God Mode enumeration failed: %s", (result.stderr or "").strip())
            return []

        if not temp_path or not os.path.exists(temp_path):
            log.warning("Access Menu: cscript output file not created.")
            return []
        try:
            with open(temp_path, "r", encoding="utf-16", errors="replace") as handle:
                payload = handle.read()
        except Exception:
            log.exception("Access Menu: failed to read cscript output file.")
            return []
        if not payload.strip():
            log.warning("Access Menu: cscript God Mode enumeration returned empty output.")
            return []
        results = []
        for line in payload.splitlines():
            name = line.strip()
            if name:
                results.append((name, name))
        results.sort(key=lambda kv: kv[0].casefold())
        return results
    finally:
        try:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
        try:
            if script_path and os.path.exists(script_path):
                os.remove(script_path)
        except Exception:
            pass


def _run_in_sta_thread(action, name="AccessMenuSTA"):
    result = {"value": None, "error": None}
    done = threading.Event()

    def runner():
        try:
            if comtypes is not None:
                comtypes.CoInitialize()
            result["value"] = action()
        except Exception as exc:
            result["error"] = exc
        finally:
            if comtypes is not None:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass
            done.set()

    thread = threading.Thread(target=runner, name=name)
    thread.daemon = True
    thread.start()
    done.wait()
    if result["error"] is not None:
        log.exception("Access Menu: STA thread action failed.", exc_info=result["error"])
    return result["value"]


def _get_god_mode_items_via_com_sta():
    if comtypes_client is None:
        return []

    def action():
        shell = comtypes_client.CreateObject("Shell.Application")
        folder = shell.NameSpace(GOD_MODE_SHELL)
        if not folder:
            return []
        try:
            items = folder.Items()
        except Exception:
            return []
        results = []
        for item in items:
            try:
                name = item.Name
            except Exception:
                continue
            if name:
                results.append((name, name))
        results.sort(key=lambda kv: kv[0].casefold())
        return results

    return _run_in_sta_thread(action, name="AccessMenuSTA-Enum") or []


def _invoke_god_mode_item_by_name(name):
    if not name:
        return

    exe_path = _get_helper_exe()
    if exe_path:
        try:
            result = subprocess.run(
                [exe_path, "--invoke", name],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )
            if result.returncode == 0:
                return
            log.warning(
                "Access Menu: helper exe failed to invoke item (%s): %s",
                name,
                (result.stderr or "").strip(),
            )
        except OSError as exc:
            log.warning("Access Menu: helper exe failed to launch for invoke: %s", exc)

    def action():
        if comtypes_client is None:
            return False
        shell = comtypes_client.CreateObject("Shell.Application")
        folder = shell.NameSpace(GOD_MODE_SHELL)
        if not folder:
            return False
        item = folder.ParseName(name)
        if not item:
            return False
        try:
            item.InvokeVerb()
            return True
        except Exception:
            return False

    if _run_in_sta_thread(action, name="AccessMenuSTA-Invoke"):
        return

    safe_name = name.replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject Shell.Application
$folder = $shell.NameSpace('shell:::{GOD_MODE_CLSID}')
if ($null -eq $folder) {{ exit 2 }}
$item = $folder.ParseName('{safe_name}')
if ($null -eq $item) {{ exit 3 }}
$item.InvokeVerb()
"""
    result = None
    ps_exe = None
    try:
        for candidate in _get_powershell_candidates():
            try:
                ps_exe = candidate
                result = subprocess.run(
                    [
                        candidate,
                        "-NoLogo",
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-STA",
                        "-Command",
                        script,
                    ],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
                break
            except FileNotFoundError:
                continue
        if result is None:
            log.warning("Access Menu: PowerShell executable not found for invoke.")
            _invoke_god_mode_item_by_name_via_cscript(name)
            return
    except Exception:
        log.exception("Access Menu: failed to invoke God Mode item via PowerShell.")
        _invoke_god_mode_item_by_name_via_cscript(name)
        return

    if result.returncode != 0:
        log.warning(
            "Access Menu: PowerShell failed to invoke God Mode item (%s): %s",
            name,
            (result.stderr or "").strip(),
        )
        _invoke_god_mode_item_by_name_via_cscript(name)


def _invoke_god_mode_item_by_name_via_cscript(name):
    if not name:
        return
    script_path = None
    try:
        script_name = f"godmode-invoke-{uuid.uuid4().hex}.vbs"
        script_path = os.path.join(tempfile.gettempdir(), script_name)
        escaped_name = name.replace('"', '""')
        script = (
            'On Error Resume Next\n'
            'Dim shell, folder, item\n'
            'Set shell = CreateObject("Shell.Application")\n'
            f'Set folder = shell.NameSpace("shell:::{GOD_MODE_CLSID}")\n'
            'If folder Is Nothing Then WScript.Quit 2\n'
            f'Set item = folder.ParseName("{escaped_name}")\n'
            'If item Is Nothing Then WScript.Quit 3\n'
            'item.InvokeVerb\n'
        )
        with open(script_path, "w", encoding="utf-8") as handle:
            handle.write(script)

        result = None
        for candidate in _get_cscript_candidates():
            try:
                result = subprocess.run(
                    [candidate, "//NoLogo", script_path],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
                break
            except FileNotFoundError:
                continue
        if result is None:
            log.warning("Access Menu: cscript executable not found for invoke.")
            return
        if result.returncode != 0:
            log.warning(
                "Access Menu: cscript failed to invoke God Mode item (%s): %s",
                name,
                (result.stderr or "").strip(),
            )
    finally:
        try:
            if script_path and os.path.exists(script_path):
                os.remove(script_path)
        except Exception:
            pass

def _get_god_mode_items():
    results = _get_god_mode_items_via_helper()
    shell = None
    if results:
        log.info("Access Menu: God Mode enumeration succeeded via helper exe.")
    else:
        results = _get_god_mode_items_via_com_sta()
        if results:
            log.info("Access Menu: God Mode enumeration succeeded via COM STA.")

    if not results:
        results = _get_god_mode_items_via_powershell()
        if results:
            log.info("Access Menu: God Mode enumeration succeeded via PowerShell.")
        else:
            results = _get_god_mode_items_via_cscript()
            if results:
                log.info("Access Menu: God Mode enumeration succeeded via cscript.")
            else:
                log.warning("Access Menu: God Mode enumeration returned no items.")
    else:
        results.sort(key=lambda kv: kv[0].casefold())

    return shell, results


class AccessMenuSettingsPanel(SettingsPanel):
    title = "Access Menu"

    def makeSettings(self, settingsSizer):
        helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        helper.addItem(wx.StaticText(self, label="Menu labels"))
        self.appsLabelCtrl = helper.addLabeledControl("Apps label", wx.TextCtrl)
        self.appsLabelCtrl.SetValue(_get_cfg("appsLabel"))
        self.powerLabelCtrl = helper.addLabeledControl("Power label", wx.TextCtrl)
        self.powerLabelCtrl.SetValue(_get_cfg("powerLabel"))
        self.settingsLabelCtrl = helper.addLabeledControl("Settings label", wx.TextCtrl)
        self.settingsLabelCtrl.SetValue(_get_cfg("settingsLabel"))
        self.folderPrefixCtrl = helper.addLabeledControl("Folder prefix", wx.TextCtrl)
        self.folderPrefixCtrl.SetValue(_get_cfg("folderPrefix"))
        self.searchLabelCtrl = helper.addLabeledControl("Search label", wx.TextCtrl)
        self.searchLabelCtrl.SetValue(_get_cfg("searchLabel"))
        self.allAppsLabelCtrl = helper.addLabeledControl("All apps label", wx.TextCtrl)
        self.allAppsLabelCtrl.SetValue(_get_cfg("allAppsLabel"))
        self.categoriesLabelCtrl = helper.addLabeledControl("Categories label", wx.TextCtrl)
        self.categoriesLabelCtrl.SetValue(_get_cfg("categoriesLabel"))
        self.browseFoldersLabelCtrl = helper.addLabeledControl("Browse folders label", wx.TextCtrl)
        self.browseFoldersLabelCtrl.SetValue(_get_cfg("browseFoldersLabel"))

        helper.addItem(wx.StaticText(self, label="Power labels"))
        self.signOutLabelCtrl = helper.addLabeledControl("Sign out label", wx.TextCtrl)
        self.signOutLabelCtrl.SetValue(_get_cfg("signOutLabel"))
        self.powerOffLabelCtrl = helper.addLabeledControl("Power off label", wx.TextCtrl)
        self.powerOffLabelCtrl.SetValue(_get_cfg("powerOffLabel"))
        self.rebootLabelCtrl = helper.addLabeledControl("Reboot label", wx.TextCtrl)
        self.rebootLabelCtrl.SetValue(_get_cfg("rebootLabel"))

        helper.addItem(wx.StaticText(self, label="Confirm prompts"))
        self.confirmTitleCtrl = helper.addLabeledControl("Confirm title", wx.TextCtrl)
        self.confirmTitleCtrl.SetValue(_get_cfg("confirmTitle"))
        self.confirmSignOutCtrl = helper.addLabeledControl("Sign out prompt", wx.TextCtrl)
        self.confirmSignOutCtrl.SetValue(_get_cfg("confirmSignOut"))
        self.confirmPowerOffCtrl = helper.addLabeledControl("Power off prompt", wx.TextCtrl)
        self.confirmPowerOffCtrl.SetValue(_get_cfg("confirmPowerOff"))
        self.confirmRebootCtrl = helper.addLabeledControl("Reboot prompt", wx.TextCtrl)
        self.confirmRebootCtrl.SetValue(_get_cfg("confirmReboot"))

        helper.addItem(wx.StaticText(self, label="Search"))
        self.searchDialogTitleCtrl = helper.addLabeledControl("Search dialog title", wx.TextCtrl)
        self.searchDialogTitleCtrl.SetValue(_get_cfg("searchDialogTitle"))
        self.searchHintCtrl = helper.addLabeledControl("Search hint", wx.TextCtrl)
        self.searchHintCtrl.SetValue(_get_cfg("searchHint"))

    def onSave(self):
        conf = config.conf[CONFIG_SECTION]
        conf["appsLabel"] = self.appsLabelCtrl.GetValue()
        conf["powerLabel"] = self.powerLabelCtrl.GetValue()
        conf["settingsLabel"] = self.settingsLabelCtrl.GetValue()
        conf["folderPrefix"] = self.folderPrefixCtrl.GetValue()
        conf["searchLabel"] = self.searchLabelCtrl.GetValue()
        conf["allAppsLabel"] = self.allAppsLabelCtrl.GetValue()
        conf["categoriesLabel"] = self.categoriesLabelCtrl.GetValue()
        conf["browseFoldersLabel"] = self.browseFoldersLabelCtrl.GetValue()
        conf["signOutLabel"] = self.signOutLabelCtrl.GetValue()
        conf["powerOffLabel"] = self.powerOffLabelCtrl.GetValue()
        conf["rebootLabel"] = self.rebootLabelCtrl.GetValue()
        conf["confirmTitle"] = self.confirmTitleCtrl.GetValue()
        conf["confirmSignOut"] = self.confirmSignOutCtrl.GetValue()
        conf["confirmPowerOff"] = self.confirmPowerOffCtrl.GetValue()
        conf["confirmReboot"] = self.confirmRebootCtrl.GetValue()
        conf["searchDialogTitle"] = self.searchDialogTitleCtrl.GetValue()
        conf["searchHint"] = self.searchHintCtrl.GetValue()


class AccessMenuSearchDialog(wx.Dialog):
    def __init__(self, parent, apps):
        title = _get_cfg("searchDialogTitle")
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.apps = apps
        self.filtered = apps

        sizer = wx.BoxSizer(wx.VERTICAL)
        hint = _get_cfg("searchHint")
        self.searchCtrl = wx.TextCtrl(self)
        self.searchCtrl.SetHint(hint)
        sizer.Add(self.searchCtrl, 0, wx.ALL | wx.EXPAND, 8)

        self.listBox = wx.ListBox(self)
        sizer.Add(self.listBox, 1, wx.ALL | wx.EXPAND, 8)

        buttonSizer = wx.StdDialogButtonSizer()
        self.launchButton = wx.Button(self, wx.ID_OK, "Launch")
        self.closeButton = wx.Button(self, wx.ID_CANCEL, "Close")
        buttonSizer.AddButton(self.launchButton)
        buttonSizer.AddButton(self.closeButton)
        buttonSizer.Realize()
        sizer.Add(buttonSizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        self.SetSizer(sizer)
        self.SetSize((520, 420))

        self._refresh_list("")

        self.searchCtrl.Bind(wx.EVT_TEXT, self._on_filter)
        self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self._on_launch)
        self.Bind(wx.EVT_BUTTON, self._on_launch, id=wx.ID_OK)
        self.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CANCEL), id=wx.ID_CANCEL)
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

    def _on_filter(self, event):
        self._refresh_list(self.searchCtrl.GetValue())

    def _refresh_list(self, query):
        query = query.strip().casefold()
        if not query:
            self.filtered = self.apps
        else:
            self.filtered = [item for item in self.apps if query in item[0].casefold()]
        self.listBox.Clear()
        for display, _path in self.filtered:
            self.listBox.Append(display)
        if self.filtered:
            self.listBox.SetSelection(0)

    def _on_launch(self, event):
        idx = self.listBox.GetSelection()
        if idx == wx.NOT_FOUND or idx >= len(self.filtered):
            return
        _display, path = self.filtered[idx]
        try:
            os.startfile(path)
        except OSError:
            subprocess.Popen(["cmd", "/c", "start", "", path])
        self.EndModal(wx.ID_OK)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        event.Skip()


class AccessMenuDialog(wx.Dialog):
    """Dialog for displaying the Access Menu"""
    
    def __init__(self, parent, tree, plugin):
        super().__init__(parent, title="Access Menu", style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.tree = tree
        self.plugin = plugin
        
        # Create main sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Create list control
        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_NO_HEADER)
        self.listCtrl.InsertColumn(0, "Item")
        self.listCtrl.SetMinSize((600, 400))
        
        # Populate with categories
        self.items = []
        self._populate_list()
        
        mainSizer.Add(self.listCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
        
        # Add buttons
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.okButton = wx.Button(self, wx.ID_OK, "OK")
        self.cancelButton = wx.Button(self, wx.ID_CANCEL, "Cancel")
        buttonSizer.Add(self.okButton, flag=wx.RIGHT, border=5)
        buttonSizer.Add(self.cancelButton)
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)
        
        self.SetSizer(mainSizer)
        self.Fit()
        self.Centre()
        
        # Bind events
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnActivate)
        self.okButton.Bind(wx.EVT_BUTTON, self.OnOK)
        self.cancelButton.Bind(wx.EVT_BUTTON, self.OnCancel)
        
        # Set focus to list and announce
        wx.CallAfter(self._announce_dialog)
    
    def _announce_dialog(self):
        """Announce the dialog and set focus"""
        self.listCtrl.SetFocus()
        if self.listCtrl.GetItemCount() > 0:
            self.listCtrl.Select(0)
            self.listCtrl.Focus(0)
        
        # Announce dialog title and first item
        first_item = self.listCtrl.GetItemText(0) if self.listCtrl.GetItemCount() > 0 else ""
        ui.message(f"Access Menu dialog. {first_item}")
    
    def _populate_list(self):
        """Populate the list with menu items"""
        # Add Apps section
        apps_label = _get_cfg("appsLabel")
        idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), f"📁 {apps_label}")
        self.items.append(("category", "apps", self.tree))
        
        # Add Power section
        power_label = _get_cfg("powerLabel")
        idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), f"⚡ {power_label}")
        self.items.append(("category", "power", None))

        # Add Settings (God Mode)
        settings_label = _get_cfg("settingsLabel")
        idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), f"⚙️ {settings_label}")
        self.items.append(("category", "settings", None))
        
        # Auto-size column
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)
    
    def OnActivate(self, event):
        """Handle double-click or Enter on an item"""
        self.OnOK(event)
    
    def OnOK(self, event):
        """Handle OK button or Enter key"""
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return
        
        item_type, item_data, extra = self.items[selected]
        
        if item_type == "category":
            if item_data == "apps":
                # Show apps submenu
                dlg = AppsMenuDialog(self, self.tree, self.plugin)
                dlg.ShowModal()
                dlg.Destroy()
            elif item_data == "power":
                # Show power submenu
                dlg = PowerMenuDialog(self, self.plugin)
                dlg.ShowModal()
                dlg.Destroy()
            elif item_data == "settings":
                dlg = SettingsMenuDialog(self, self.plugin)
                dlg.ShowModal()
                dlg.Destroy()
        
        # Don't close the main dialog, let user choose again or press Cancel
    
    def OnCancel(self, event):
        """Handle Cancel button or Escape key"""
        self.EndModal(wx.ID_CANCEL)


class AppsMenuDialog(wx.Dialog):
    """Dialog for Apps submenu"""
    
    def __init__(self, parent, tree, plugin):
        super().__init__(parent, title=_get_cfg("appsLabel"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.tree = tree
        self.plugin = plugin
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Create list control
        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_NO_HEADER)
        self.listCtrl.InsertColumn(0, "Item")
        self.listCtrl.SetMinSize((600, 400))
        
        self.items = []
        self._populate_list()
        
        mainSizer.Add(self.listCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
        
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.okButton = wx.Button(self, wx.ID_OK, "OK")
        self.cancelButton = wx.Button(self, wx.ID_CANCEL, "Cancel")
        buttonSizer.Add(self.okButton, flag=wx.RIGHT, border=5)
        buttonSizer.Add(self.cancelButton)
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)
        
        self.SetSizer(mainSizer)
        self.Fit()
        self.Centre()
        
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnActivate)
        self.okButton.Bind(wx.EVT_BUTTON, self.OnOK)
        self.cancelButton.Bind(wx.EVT_BUTTON, self.OnCancel)
        
        wx.CallAfter(self._announce_dialog)
    
    def _announce_dialog(self):
        """Announce the dialog and set focus"""
        self.listCtrl.SetFocus()
        if self.listCtrl.GetItemCount() > 0:
            self.listCtrl.Select(0)
            self.listCtrl.Focus(0)
        
        count = self.listCtrl.GetItemCount()
        first_item = self.listCtrl.GetItemText(0) if count > 0 else ""
        ui.message(f"{_get_cfg('appsLabel')} dialog. {count} apps. {first_item}")
    
    def _populate_list(self):
        """Populate with all apps sorted A-Z"""
        apps = _flatten_apps(self.tree)
        
        for display_name, path in apps:
            idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), display_name)
            self.items.append(("app", path))
        
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)
    
    def OnActivate(self, event):
        self.OnOK(event)
    
    def OnOK(self, event):
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return
        
        item_type, path = self.items[selected]
        
        if item_type == "app":
            # Launch the app
            try:
                os.startfile(path)
            except OSError:
                subprocess.Popen(["cmd", "/c", "start", "", path])
            
            # Close both dialogs
            self.EndModal(wx.ID_OK)
            if self.GetParent():
                self.GetParent().EndModal(wx.ID_OK)
    
    def OnCancel(self, event):
        self.EndModal(wx.ID_CANCEL)


class SettingsMenuDialog(wx.Dialog):
    """Dialog for Settings (God Mode) submenu"""

    def __init__(self, parent, plugin):
        super().__init__(parent, title=_get_cfg("settingsLabel"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.plugin = plugin
        self._shell, self.items = _get_god_mode_items()
        self._empty_item = None
        if not self.items:
            self._empty_item = _("No settings items available")
            self.items = [(self._empty_item, None)]

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_NO_HEADER)
        self.listCtrl.InsertColumn(0, "Item")
        self.listCtrl.SetMinSize((600, 400))

        self._populate_list()

        mainSizer.Add(self.listCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)

        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.okButton = wx.Button(self, wx.ID_OK, "OK")
        self.cancelButton = wx.Button(self, wx.ID_CANCEL, "Cancel")
        buttonSizer.Add(self.okButton, flag=wx.RIGHT, border=5)
        buttonSizer.Add(self.cancelButton)
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)

        self.SetSizer(mainSizer)
        self.Fit()
        self.Centre()

        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnActivate)
        self.okButton.Bind(wx.EVT_BUTTON, self.OnOK)
        self.cancelButton.Bind(wx.EVT_BUTTON, self.OnCancel)

        wx.CallAfter(self._announce_dialog)

    def _announce_dialog(self):
        """Announce the dialog and set focus"""
        self.listCtrl.SetFocus()
        count = self.listCtrl.GetItemCount()
        if count > 0:
            self.listCtrl.Select(0)
            self.listCtrl.Focus(0)
        first_item = self.listCtrl.GetItemText(0) if count > 0 else ""
        if count == 0:
            ui.message(f"{_get_cfg('settingsLabel')} dialog. No items available.")
        elif self._empty_item and count == 1:
            ui.message(f"{_get_cfg('settingsLabel')} dialog. {first_item}")
        else:
            ui.message(f"{_get_cfg('settingsLabel')} dialog. {count} items. {first_item}")

    def _populate_list(self):
        for name, _item in self.items:
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), name)
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)

    def OnActivate(self, event):
        self.OnOK(event)

    def OnOK(self, event):
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return

        if selected >= len(self.items):
            return

        _label, item = self.items[selected]
        if item is None:
            ui.message(_("No settings items available."))
            return
        if isinstance(item, str):
            _invoke_god_mode_item_by_name(item)
            self.EndModal(wx.ID_OK)
            if self.GetParent():
                self.GetParent().EndModal(wx.ID_OK)
            return
        try:
            item.InvokeVerb()
        except Exception:
            path = getattr(item, "Path", "")
            if path:
                try:
                    os.startfile(path)
                except OSError:
                    subprocess.Popen(["explorer.exe", path])
            else:
                _open_god_mode_folder()

        self.EndModal(wx.ID_OK)
        if self.GetParent():
            self.GetParent().EndModal(wx.ID_OK)

    def OnCancel(self, event):
        self.EndModal(wx.ID_CANCEL)


class PowerMenuDialog(wx.Dialog):
    """Dialog for Power submenu"""
    
    def __init__(self, parent, plugin):
        super().__init__(parent, title=_get_cfg("powerLabel"), style=wx.DEFAULT_DIALOG_STYLE)
        self.plugin = plugin
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_NO_HEADER)
        self.listCtrl.InsertColumn(0, "Item")
        self.listCtrl.SetMinSize((400, 200))
        
        self.items = []
        self._populate_list()
        
        mainSizer.Add(self.listCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
        
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.okButton = wx.Button(self, wx.ID_OK, "OK")
        self.cancelButton = wx.Button(self, wx.ID_CANCEL, "Cancel")
        buttonSizer.Add(self.okButton, flag=wx.RIGHT, border=5)
        buttonSizer.Add(self.cancelButton)
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)
        
        self.SetSizer(mainSizer)
        self.Fit()
        self.Centre()
        
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnActivate)
        self.okButton.Bind(wx.EVT_BUTTON, self.OnOK)
        self.cancelButton.Bind(wx.EVT_BUTTON, self.OnCancel)
        
        wx.CallAfter(self._announce_dialog)
    
    def _announce_dialog(self):
        """Announce the dialog and set focus"""
        self.listCtrl.SetFocus()
        if self.listCtrl.GetItemCount() > 0:
            self.listCtrl.Select(0)
            self.listCtrl.Focus(0)
        
        first_item = self.listCtrl.GetItemText(0) if self.listCtrl.GetItemCount() > 0 else ""
        ui.message(f"{_get_cfg('powerLabel')} dialog. {first_item}")
    
    def _populate_list(self):
        """Populate with power options"""
        options = [
            (_get_cfg("signOutLabel"), "signout"),
            (_get_cfg("powerOffLabel"), "poweroff"),
            (_get_cfg("rebootLabel"), "reboot"),
        ]
        
        for label, action in options:
            idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), label)
            self.items.append(("power", action))
        
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)
    
    def OnActivate(self, event):
        self.OnOK(event)
    
    def OnOK(self, event):
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return
        
        item_type, action = self.items[selected]
        
        if item_type == "power":
            # Show confirmation dialog
            if action == "signout":
                msg = _get_cfg("confirmSignOut")
            elif action == "poweroff":
                msg = _get_cfg("confirmPowerOff")
            elif action == "reboot":
                msg = _get_cfg("confirmReboot")
            
            result = wx.MessageBox(msg, _get_cfg("confirmTitle"), wx.YES_NO | wx.ICON_QUESTION)
            
            if result == wx.YES:
                if action == "signout":
                    subprocess.Popen(["shutdown", "/l"])
                elif action == "poweroff":
                    subprocess.Popen(["shutdown", "/s", "/t", "0"])
                elif action == "reboot":
                    subprocess.Popen(["shutdown", "/r", "/t", "0"])
                
                # Close all dialogs
                self.EndModal(wx.ID_OK)
                if self.GetParent():
                    self.GetParent().EndModal(wx.ID_OK)
    
    def OnCancel(self, event):
        self.EndModal(wx.ID_CANCEL)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = ADDON_SUMMARY

    __gestures = {
        "kb:NVDA+shift+m": "openAccessMenu",
    }

    def __init__(self):
        super().__init__()
        _ensure_config()
        if AccessMenuSettingsPanel not in NVDASettingsDialog.categoryClasses:
            NVDASettingsDialog.categoryClasses.append(AccessMenuSettingsPanel)
        self._menu_actions = {}
        self.bindGestures(self.__gestures)
        log.info("Access Menu add-on: GlobalPlugin initialized")

    @script(description="Open the Access Menu")
    def script_openAccessMenu(self, gesture):
        wx.CallAfter(self._show_dialog)

    def _show_dialog(self):
        """Show the Access Menu as a dialog"""
        
        # Build the menu tree
        tree = _build_tree()
        
        # Create and show the dialog
        gui.mainFrame.prePopup()
        dlg = AccessMenuDialog(gui.mainFrame, tree, self)
        try:
            dlg.ShowModal()
        finally:
            dlg.Destroy()
            gui.mainFrame.postPopup()

    def _populate_apps_menu(self, menu):
        tree = _build_tree()
        search_id = wx.NewIdRef()
        menu.Append(search_id, _get_cfg("searchLabel"))
        menu.Bind(wx.EVT_MENU, lambda evt: self._open_search_dialog(tree), id=search_id)

        menu.AppendSeparator()

        all_apps_menu = wx.Menu()
        self._populate_all_apps_menu(all_apps_menu, tree)
        menu.AppendSubMenu(all_apps_menu, _get_cfg("allAppsLabel"))

        categories_menu = wx.Menu()
        self._populate_categories_menu(categories_menu, tree)
        menu.AppendSubMenu(categories_menu, _get_cfg("categoriesLabel"))

        browse_menu = wx.Menu()
        self._add_tree_to_menu(browse_menu, tree)
        menu.AppendSubMenu(browse_menu, _get_cfg("browseFoldersLabel"))

    def _populate_all_apps_menu(self, menu, tree):
        apps = _flatten_apps(tree)
        for display, path in apps:
            item_id = wx.NewIdRef()
            menu.Append(item_id, display)
            self._menu_actions[item_id] = path
            menu.Bind(wx.EVT_MENU, self._on_launch_app, id=item_id)

    def _populate_categories_menu(self, menu, tree):
        for name, subtree in _top_level_categories(tree):
            sub_menu = wx.Menu()
            self._add_tree_to_menu(sub_menu, subtree)
            menu.AppendSubMenu(sub_menu, name)

    def _add_tree_to_menu(self, menu, tree):
        for name, value in _sorted_items(tree):
            if isinstance(value, dict):
                sub_menu = wx.Menu()
                self._add_tree_to_menu(sub_menu, value)
                menu.AppendSubMenu(sub_menu, f"{_get_cfg('folderPrefix')}{name}")
            else:
                item_id = wx.NewIdRef()
                menu.Append(item_id, name)
                self._menu_actions[item_id] = value
                menu.Bind(wx.EVT_MENU, self._on_launch_app, id=item_id)

    def _populate_power_menu(self, menu):
        self._add_power_item(menu, _get_cfg("signOutLabel"), ["shutdown", "/l"], _get_cfg("confirmSignOut"))
        self._add_power_item(menu, _get_cfg("powerOffLabel"), ["shutdown", "/s", "/t", "0"], _get_cfg("confirmPowerOff"))
        self._add_power_item(menu, _get_cfg("rebootLabel"), ["shutdown", "/r", "/t", "0"], _get_cfg("confirmReboot"))

    def _add_power_item(self, menu, label, command, prompt):
        item_id = wx.NewIdRef()
        menu.Append(item_id, label)
        menu.Bind(wx.EVT_MENU, lambda evt: self._confirm_and_run(command, prompt), id=item_id)

    def _confirm_and_run(self, command, prompt):
        result = wx.MessageBox(prompt, _get_cfg("confirmTitle"), style=wx.YES_NO | wx.ICON_QUESTION | wx.CENTER)
        if result == wx.YES:
            subprocess.Popen(command)

    def _open_search_dialog(self, tree):
        apps = _flatten_apps(tree)
        dlg = AccessMenuSearchDialog(gui.mainFrame, apps)
        try:
            dlg.ShowModal()
        finally:
            dlg.Destroy()

    def _on_launch_app(self, event):
        path = self._menu_actions.get(event.GetId())
        if not path:
            return
        try:
            os.startfile(path)
        except OSError:
            # If startfile fails, try shell execute via cmd.
            subprocess.Popen(["cmd", "/c", "start", "", path])
