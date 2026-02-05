# NVDA add-on: Access Menu
# Provides a start menu replacement with Apps and Power menus.

import os
import subprocess

import wx
import addonHandler
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

CONFIG_SECTION = "accessMenu"
CONFIG_SPEC = {
    "appsLabel": "string(default='Apps')",
    "powerLabel": "string(default='Power')",
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
    path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    return [path] if os.path.isdir(path) else []


def _build_tree():
    tree = {}
    # Track app names that exist in subfolders
    subfolder_apps = set()
    
    # First pass: collect all apps in subfolders
    for root in _start_menu_roots():
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            # Skip root level in first pass
            if rel == ".":
                continue
            for filename in filenames:
                base, ext = os.path.splitext(filename)
                if ext.lower() in APP_EXTENSIONS:
                    subfolder_apps.add(base.lower())
    
    # Second pass: build tree, skip root duplicates
    for root in _start_menu_roots():
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            parts = [] if rel == "." else rel.split(os.sep)
            node = tree
            for part in parts:
                node = node.setdefault(part, {})
            
            is_root = (rel == ".")
            
            for filename in filenames:
                base, ext = os.path.splitext(filename)
                if ext.lower() not in APP_EXTENSIONS:
                    continue
                
                # Skip root-level apps that exist in subfolders
                if is_root and base.lower() in subfolder_apps:
                    log.debug(f"Skipping root duplicate: {base}")
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


class AccessMenuSettingsPanel(SettingsPanel):
    title = "Access Menu"

    def makeSettings(self, settingsSizer):
        helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        helper.addItem(wx.StaticText(self, label="Menu labels"))
        self.appsLabelCtrl = helper.addLabeledControl("Apps label", wx.TextCtrl)
        self.appsLabelCtrl.SetValue(_get_cfg("appsLabel"))
        self.powerLabelCtrl = helper.addLabeledControl("Power label", wx.TextCtrl)
        self.powerLabelCtrl.SetValue(_get_cfg("powerLabel"))
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
        
        # Don't close the main dialog, let user choose again or press Cancel
    
    def OnCancel(self, event):
        """Handle Cancel button or Escape key"""
        self.EndModal(wx.ID_CANCEL)


class AppsMenuDialog(wx.Dialog):
    """Dialog for Apps submenu"""
    
    def __init__(self, parent, tree, plugin, breadcrumb=""):
        title = _get_cfg("appsLabel") if not breadcrumb else breadcrumb
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.tree = tree
        self.plugin = plugin
        self.breadcrumb = breadcrumb
        
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
        title = self.breadcrumb if self.breadcrumb else _get_cfg('appsLabel')
        ui.message(f"{title} dialog. {count} items. {first_item}")
    
    def _populate_list(self):
        """Populate with folders first, then root-level apps"""
        folders = []
        apps = []
        
        # Separate folders and apps
        for name, value in _sorted_items(self.tree):
            if isinstance(value, dict):
                folders.append((name, value))
            else:
                apps.append((name, value))
        
        # Add folders first
        for name, subtree in folders:
            idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), f"📁 {name}")
            self.items.append(("folder", name, subtree))
        
        # Add root-level apps
        for name, path in apps:
            idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), name)
            self.items.append(("app", path, None))
        
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)
    
    def OnActivate(self, event):
        self.OnOK(event)
    
    def OnOK(self, event):
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return
        
        item_type, data, extra = self.items[selected]
        
        if item_type == "folder":
            # Open subdialog for this folder
            folder_name = data
            subtree = extra
            new_breadcrumb = f"{self.breadcrumb} > {folder_name}" if self.breadcrumb else folder_name
            dlg = AppsMenuDialog(self, subtree, self.plugin, new_breadcrumb)
            dlg.ShowModal()
            dlg.Destroy()
        elif item_type == "app":
            # Launch the app
            path = data
            log.info(f"Attempting to launch: {path}")
            try:
                # Use explorer.exe to launch shortcuts - more reliable than os.startfile
                subprocess.Popen(["explorer.exe", path], creationflags=subprocess.CREATE_NO_WINDOW)
                log.info(f"Launch command sent via explorer.exe")
            except Exception as e:
                log.error(f"Launch failed: {e}")
                ui.message(f"Failed to launch application")
            
            # Close all dialogs by propagating up
            self.EndModal(wx.ID_OK)
            parent = self.GetParent()
            while parent and isinstance(parent, (AppsMenuDialog, AccessMenuDialog)):
                parent.EndModal(wx.ID_OK)
                parent = parent.GetParent()
    
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
        dlg = AccessMenuDialog(gui.mainFrame, tree, self)
        dlg.ShowModal()
        dlg.Destroy()

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
