# -*- coding: utf-8 -*-
"""Toggle automatic project saving (Riyan)."""

import os
from pyrevit import EXEC_PARAMS, script, forms
from pyrevit.forms import WPFWindow
import clr
clr.AddReference("System.Windows.Presentation")
import System

STATE_VAR = 'RIYAN_AUTOSAVE_ENABLED'
INTERVAL_VAR = 'RIYAN_AUTOSAVE_INTERVAL'
LAST_SAVES_VAR = 'RIYAN_AUTOSAVE_LAST_SAVES'

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    # Retrieve current active state of AutoSave
    state = script.get_envvar(STATE_VAR) or False
    
    # Try to set the correct icon on startup
    try:
        on_icon = script_cmp.get_bundle_file('on.png')
        off_icon = script_cmp.get_bundle_file('off.png')
        if state:
            ui_button_cmp.set_icon(on_icon)
        else:
            ui_button_cmp.set_icon(off_icon)
    except Exception:
        pass


class AutoSaveSettingsWindow(WPFWindow):
    def __init__(self, xaml_file_name, current_interval):
        WPFWindow.__init__(self, xaml_file_name)
        self.IntervalInput.Text = str(current_interval)
        self.result = None

    def TitleBar_MouseDown(self, sender, args):
        if args.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
            self.DragMove()

    def SaveButton_Click(self, sender, args):
        new_val = self.IntervalInput.Text
        try:
            val = int(new_val)
            if val > 0:
                self.result = val
                self.DialogResult = True
                self.Close()
            else:
                forms.alert("Please enter a positive number of minutes.", title="Invalid Input")
        except ValueError:
            forms.alert("Please enter a valid whole number.", title="Invalid Input")

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


if __name__ == '__main__':
    if EXEC_PARAMS.config_mode:
        # Shift-Click detected: Configuration mode
        current_interval = script.get_envvar(INTERVAL_VAR) or 10
        xaml_file = os.path.join(os.path.dirname(__file__), "settings.xaml")
        
        if os.path.exists(xaml_file):
            try:
                window = AutoSaveSettingsWindow(xaml_file, current_interval)
                window.ShowDialog()
                if window.DialogResult and window.result is not None:
                    script.set_envvar(INTERVAL_VAR, window.result)
                    forms.toast("AutoSave interval updated to {} minutes.".format(window.result), title="AutoSave Settings")
            except Exception as e:
                forms.alert("Could not load settings window:\n{}".format(e))
        else:
            # Fallback if xaml file not found
            new_interval = forms.ask_for_string(
                default=str(current_interval),
                title="AutoSave Settings",
                prompt="Enter auto-save interval in minutes (positive integer):"
            )
            if new_interval:
                try:
                    val = int(new_interval)
                    if val > 0:
                        script.set_envvar(INTERVAL_VAR, val)
                        forms.toast("AutoSave interval updated to {} minutes.".format(val), title="AutoSave Settings")
                    else:
                        forms.alert("Please enter a positive integer.", title="Error")
                except ValueError:
                    forms.alert("Invalid input.", title="Error")
    else:
        # Regular Click: Toggle active state
        is_active = script.get_envvar(STATE_VAR) or False
        new_state = not is_active
        
        script.set_envvar(STATE_VAR, new_state)
        script.toggle_icon(new_state)
        
        if new_state:
            # Initialize or reset last saves dictionary
            script.set_envvar(LAST_SAVES_VAR, {})
            # Read current interval for confirmation message
            current_interval = script.get_envvar(INTERVAL_VAR) or 10
            forms.toast("AutoSave is now ENABLED (Saving every {} minutes).".format(current_interval), title="AutoSave")
        else:
            forms.toast("AutoSave is now DISABLED.", title="AutoSave")
