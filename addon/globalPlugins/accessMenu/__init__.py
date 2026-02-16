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
    "searchLabel": "string(default='Search')",
    "favoritesLabel": "string(default='Favorites')",
    "appsLabel": "string(default='Apps')",
    "powerLabel": "string(default='Power')",
    "aboutLabel": "string(default='About')",
    "winutilsLabel": "string(default='WinUtils')",
    "folderPrefix": "string(default='Folder: ')",
    "signOutLabel": "string(default='Sign out')",
    "powerOffLabel": "string(default='Power off')",
    "rebootLabel": "string(default='Reboot')",
    "confirmTitle": "string(default='Confirm')",
    "confirmSignOut": "string(default='Sign out of Windows?')",
    "confirmPowerOff": "string(default='Power off the PC?')",
    "confirmReboot": "string(default='Restart the PC?')",
    "searchDialogTitle": "string(default='Search Apps')",
    "searchHint": "string(default='Type to filter apps')",
    "favorites": "list(default=list())",
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


def _ensure_config():
    if CONFIG_SECTION not in config.conf.spec:
        config.conf.spec[CONFIG_SECTION] = CONFIG_SPEC


def _get_cfg(key):
    return config.conf[CONFIG_SECTION].get(key, CONFIG_SPEC[key].split("default='", 1)[1].split("'")[0])


def _get_favorites():
    """Get list of favorite app paths from config"""
    try:
        favorites = list(config.conf[CONFIG_SECTION]["favorites"])
        # Validate that favorites exist
        valid_favorites = [path for path in favorites if os.path.exists(path)]
        # Update config if we removed any broken favorites
        if len(valid_favorites) != len(favorites):
            config.conf[CONFIG_SECTION]["favorites"] = valid_favorites
        return valid_favorites
    except:
        return []


def _save_favorites(favorites_list):
    """Save favorites list to config"""
    config.conf[CONFIG_SECTION]["favorites"] = favorites_list


def _add_favorite(path):
    """Add a path to favorites"""
    favorites = _get_favorites()
    if path not in favorites:
        favorites.append(path)
        _save_favorites(favorites)
        return True
    return False


def _remove_favorite(path):
    """Remove a path from favorites"""
    favorites = _get_favorites()
    if path in favorites:
        favorites.remove(path)
        _save_favorites(favorites)
        return True
    return False


def _get_favorite_apps():
    """Get list of (display_name, path) tuples for favorites"""
    favorites = _get_favorites()
    apps = []
    for path in favorites:
        if os.path.exists(path):
            # Extract display name from path
            basename = os.path.basename(path)
            name, ext = os.path.splitext(basename)
            apps.append((name, path))
    return apps


def _show_modal_dialog(dialog):
    """Show and always destroy a dialog."""
    try:
        return dialog.ShowModal()
    finally:
        dialog.Destroy()


def _launch_via_explorer(path, attempt_log_label=None):
    """Launch an app path through explorer.exe for shortcut reliability."""
    if attempt_log_label:
        log.info(f"{attempt_log_label}: {path}")
    try:
        subprocess.Popen(["explorer.exe", path], creationflags=subprocess.CREATE_NO_WINDOW)
        if attempt_log_label:
            log.info("Launch command sent via explorer.exe")
        return True
    except Exception as e:
        log.error(f"Launch failed: {e}")
        ui.message("Failed to launch application")
        return False


def _close_parent_dialogs(dialog, dialog_types):
    """Close parent dialogs while they match the allowed dialog types."""
    parent = dialog.GetParent()
    while parent and isinstance(parent, dialog_types):
        parent.EndModal(wx.ID_OK)
        parent = parent.GetParent()


class AccessMenuSettingsPanel(SettingsPanel):
    title = "Access Menu"

    def makeSettings(self, settingsSizer):
        helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        # Favorites management section
        helper.addItem(wx.StaticText(self, label="Manage Favorites"))
        self.favoritesListBox = wx.ListBox(self, style=wx.LB_SINGLE)
        self.favoritesListBox.SetMinSize((400, 200))
        helper.addItem(self.favoritesListBox)
        
        # Favorites buttons
        buttonsSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.addFavButton = wx.Button(self, label="Browse and Add...")
        self.addFavButton.Bind(wx.EVT_BUTTON, self.onAddFavorite)
        buttonsSizer.Add(self.addFavButton, flag=wx.RIGHT, border=5)
        
        self.removeFavButton = wx.Button(self, label="Remove")
        self.removeFavButton.Bind(wx.EVT_BUTTON, self.onRemoveFavorite)
        buttonsSizer.Add(self.removeFavButton, flag=wx.RIGHT, border=5)
        
        self.moveUpButton = wx.Button(self, label="Move Up")
        self.moveUpButton.Bind(wx.EVT_BUTTON, self.onMoveUp)
        buttonsSizer.Add(self.moveUpButton, flag=wx.RIGHT, border=5)
        
        self.moveDownButton = wx.Button(self, label="Move Down")
        self.moveDownButton.Bind(wx.EVT_BUTTON, self.onMoveDown)
        buttonsSizer.Add(self.moveDownButton)
        
        helper.addItem(buttonsSizer)
        
        # About button
        helper.addItem(wx.StaticText(self, label=""))
        self.aboutButton = wx.Button(self, label="About Access Menu")
        self.aboutButton.Bind(wx.EVT_BUTTON, self.onAbout)
        helper.addItem(self.aboutButton)
        
        # Load favorites into list
        self._refresh_favorites_list()

    def _refresh_favorites_list(self):
        """Refresh the favorites list box"""
        self.favoritesListBox.Clear()
        favorites = _get_favorite_apps()
        for name, path in favorites:
            self.favoritesListBox.Append(name, path)
    
    def onAddFavorite(self, event):
        """Show app picker dialog to add favorite"""
        tree = _build_tree()
        apps = _flatten_apps(tree)
        dlg = AccessMenuSearchDialog(self, apps, picker_mode=True)
        dlg.SetTitle("Select App to Add to Favorites")
        if dlg.ShowModal() == wx.ID_OK:
            # Get selected app from search dialog
            idx = dlg.listBox.GetSelection()
            if idx != wx.NOT_FOUND and idx < len(dlg.filtered):
                _display, path = dlg.filtered[idx]
                if _add_favorite(path):
                    self._refresh_favorites_list()
                    ui.message(f"Added {_display} to favorites")
                else:
                    ui.message(f"{_display} is already in favorites")
        dlg.Destroy()
    
    def onRemoveFavorite(self, event):
        """Remove selected favorite"""
        idx = self.favoritesListBox.GetSelection()
        if idx == wx.NOT_FOUND:
            ui.message("No favorite selected")
            return
        path = self.favoritesListBox.GetClientData(idx)
        name = self.favoritesListBox.GetString(idx)
        if _remove_favorite(path):
            self._refresh_favorites_list()
            ui.message(f"Removed {name} from favorites")
    
    def onMoveUp(self, event):
        """Move selected favorite up"""
        idx = self.favoritesListBox.GetSelection()
        if idx == wx.NOT_FOUND or idx == 0:
            return
        favorites = _get_favorites()
        favorites[idx], favorites[idx-1] = favorites[idx-1], favorites[idx]
        _save_favorites(favorites)
        self._refresh_favorites_list()
        self.favoritesListBox.SetSelection(idx-1)
    
    def onMoveDown(self, event):
        """Move selected favorite down"""
        idx = self.favoritesListBox.GetSelection()
        favorites = _get_favorites()
        if idx == wx.NOT_FOUND or idx >= len(favorites) - 1:
            return
        favorites[idx], favorites[idx+1] = favorites[idx+1], favorites[idx]
        _save_favorites(favorites)
        self._refresh_favorites_list()
        self.favoritesListBox.SetSelection(idx+1)
    
    def onAbout(self, event):
        """Show about dialog"""
        _show_modal_dialog(AboutDialog(self))

    def onSave(self):
        # Favorites are saved immediately on add/remove/reorder
        # No other settings to save
        pass


class AccessMenuSearchDialog(wx.Dialog):
    def __init__(self, parent, apps, picker_mode=False):
        title = _get_cfg("searchDialogTitle")
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.apps = apps
        self.filtered = apps
        self.picker_mode = picker_mode  # If True, don't launch on select

        sizer = wx.BoxSizer(wx.VERTICAL)
        hint = _get_cfg("searchHint")
        self.searchCtrl = wx.TextCtrl(self)
        self.searchCtrl.SetHint(hint)
        sizer.Add(self.searchCtrl, 0, wx.ALL | wx.EXPAND, 8)

        self.listBox = wx.ListBox(self)
        sizer.Add(self.listBox, 1, wx.ALL | wx.EXPAND, 8)

        buttonSizer = wx.StdDialogButtonSizer()
        # Change button label based on mode
        button_label = "Add" if picker_mode else "Launch"
        self.launchButton = wx.Button(self, wx.ID_OK, button_label)
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
        
        if self.picker_mode:
            # Just close dialog, don't launch
            self.EndModal(wx.ID_OK)
            return

        _launch_via_explorer(path)
        self.EndModal(wx.ID_OK)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        elif key == wx.WXK_RETURN:
            self._on_launch(event)
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
        categories = [
            (_get_cfg("searchLabel"), "search", self.tree),
            (_get_cfg("favoritesLabel"), "favorites", None),
            (_get_cfg("appsLabel"), "apps", self.tree),
            (_get_cfg("winutilsLabel"), "winutils", None),
            (_get_cfg("powerLabel"), "power", None),
            (_get_cfg("aboutLabel"), "about", None),
        ]

        for label, category_name, extra in categories:
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), label)
            self.items.append(("category", category_name, extra))
        
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
        
        item_type, item_data, _extra = self.items[selected]

        if item_type == "category":
            if item_data == "winutils":
                self.EndModal(wx.ID_OK)
                subprocess.Popen(
                    ["powershell", "-Command", "iwr -useb https://christitus.com/win | iex"],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                return
            dialog_builders = {
                "search": lambda: AccessMenuSearchDialog(self, _flatten_apps(self.tree)),
                "favorites": lambda: FavoritesMenuDialog(self, self.plugin),
                "apps": lambda: AppsMenuDialog(self, self.tree, self.plugin),
                "power": lambda: PowerMenuDialog(self, self.plugin),
                "about": lambda: AboutDialog(self),
            }
            builder = dialog_builders.get(item_data)
            if builder:
                _show_modal_dialog(builder())
        
        # Don't close the main dialog, let user choose again or press Cancel
    
    def OnCancel(self, event):
        """Handle Cancel button or Escape key"""
        self.EndModal(wx.ID_CANCEL)


class AboutDialog(wx.Dialog):
    """About dialog showing add-on information"""
    
    def __init__(self, parent):
        super().__init__(parent, title="About", style=wx.DEFAULT_DIALOG_STYLE)
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Get add-on info from manifest
        try:
            addon = addonHandler.getCodeAddon()
            name = addon.manifest.get("summary", "Access Menu")
            version = addon.manifest.get("version", "Unknown")
            author = addon.manifest.get("author", "Unknown")
            description = addon.manifest.get("description", "Start menu replacement with Apps and Power menus.")
        except:
            name = "Access Menu"
            version = "Unknown"
            author = "Unknown"
            description = "Start menu replacement with Apps and Power menus."
        
        # Add info text
        info_text = f"{name}\n\nVersion: {version}\nAuthor: {author}\n\n{description}"
        infoLabel = wx.StaticText(self, label=info_text)
        mainSizer.Add(infoLabel, flag=wx.ALL, border=20)
        
        # Add OK button
        buttonSizer = wx.StdDialogButtonSizer()
        okButton = wx.Button(self, wx.ID_OK, "OK")
        buttonSizer.AddButton(okButton)
        buttonSizer.Realize()
        mainSizer.Add(buttonSizer, flag=wx.ALL | wx.ALIGN_CENTER, border=10)
        
        self.SetSizer(mainSizer)
        self.Fit()
        self.Centre()
        
        self.Bind(wx.EVT_BUTTON, self.OnOK, id=wx.ID_OK)
        
        wx.CallAfter(self._announce_dialog)
    
    def _announce_dialog(self):
        """Announce the dialog"""
        try:
            addon = addonHandler.getCodeAddon()
            name = addon.manifest.get("summary", "Access Menu")
            version = addon.manifest.get("version", "Unknown")
        except:
            name = "Access Menu"
            version = "Unknown"
        ui.message(f"About {name} version {version}")
    
    def OnOK(self, event):
        self.EndModal(wx.ID_OK)


class FavoritesMenuDialog(wx.Dialog):
    """Dialog for Favorites submenu"""
    
    def __init__(self, parent, plugin):
        super().__init__(parent, title=_get_cfg("favoritesLabel"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.plugin = plugin
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
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
        if count == 0:
            ui.message(f"{_get_cfg('favoritesLabel')} dialog. No favorites. Press Escape to close.")
        else:
            first_item = self.listCtrl.GetItemText(0) if count > 0 else ""
            ui.message(f"{_get_cfg('favoritesLabel')} dialog. {count} favorites. {first_item}")
    
    def _populate_list(self):
        """Populate with favorite apps"""
        favorite_apps = _get_favorite_apps()
        
        if not favorite_apps:
            # Show empty message
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), "(No favorites)")
            self.items.append(("empty", None, None))
        else:
            for name, path in favorite_apps:
                self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), name)
                self.items.append(("app", path, None))
        
        self.listCtrl.SetColumnWidth(0, wx.LIST_AUTOSIZE)
    
    def OnActivate(self, event):
        self.OnOK(event)
    
    def OnOK(self, event):
        selected = self.listCtrl.GetFirstSelected()
        if selected < 0:
            self.EndModal(wx.ID_CANCEL)
            return
        
        item_type, path, _ = self.items[selected]
        
        if item_type == "empty":
            # Nothing to launch
            self.EndModal(wx.ID_CANCEL)
            return
        
        if item_type == "app":
            # Launch the app
            _launch_via_explorer(path, "Launching favorite")
            
            # Close all dialogs
            self.EndModal(wx.ID_OK)
            _close_parent_dialogs(self, (AccessMenuDialog,))
    
    def OnCancel(self, event):
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
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), name)
            self.items.append(("folder", name, subtree))
        
        # Add root-level apps
        for name, path in apps:
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), name)
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
            _show_modal_dialog(AppsMenuDialog(self, subtree, self.plugin, new_breadcrumb))
        elif item_type == "app":
            # Launch the app
            path = data
            _launch_via_explorer(path, "Attempting to launch")
            
            # Close all dialogs by propagating up
            self.EndModal(wx.ID_OK)
            _close_parent_dialogs(self, (AppsMenuDialog, AccessMenuDialog))
    
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
            self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), label)
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
            action_config = {
                "signout": (_get_cfg("confirmSignOut"), ["shutdown", "/l"]),
                "poweroff": (_get_cfg("confirmPowerOff"), ["shutdown", "/s", "/t", "0"]),
                "reboot": (_get_cfg("confirmReboot"), ["shutdown", "/r", "/t", "0"]),
            }
            configured_action = action_config.get(action)
            if not configured_action:
                return
            msg, command = configured_action

            result = wx.MessageBox(msg, _get_cfg("confirmTitle"), wx.YES_NO | wx.ICON_QUESTION)
            
            if result == wx.YES:
                subprocess.Popen(command)
                
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
        _show_modal_dialog(AccessMenuDialog(gui.mainFrame, tree, self))
