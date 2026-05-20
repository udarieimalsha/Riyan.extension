# -*- coding: utf-8 -*-
"""AutoSave hook on view activation (Riyan)."""

import time
from pyrevit import EXEC_PARAMS, script, forms

STATE_VAR = 'RIYAN_AUTOSAVE_ENABLED'
INTERVAL_VAR = 'RIYAN_AUTOSAVE_INTERVAL'
LAST_SAVES_VAR = 'RIYAN_AUTOSAVE_LAST_SAVES'

def run():
    # 1. Check if AutoSave is enabled
    enabled = script.get_envvar(STATE_VAR) or False
    if not enabled:
        return

    # 2. Get event arguments and active document
    try:
        args = EXEC_PARAMS.event_args
        if not args:
            return
        doc = args.Document
        if not doc:
            return
    except Exception:
        return

    # 3. Check document properties (must not be read-only, family, or unsaved)
    if doc.IsReadOnly or doc.IsFamilyDocument:
        return

    doc_path = doc.PathName
    if not doc_path:
        # Unsaved new project (e.g. "Project1"). Do not auto-save to avoid prompting Save As dialog.
        return

    # 4. Check if document has actually been modified
    try:
        if not doc.IsModified:
            return
    except Exception:
        pass

    # 5. Check if the interval has passed
    interval = script.get_envvar(INTERVAL_VAR) or 10  # default 10 minutes
    last_saves = script.get_envvar(LAST_SAVES_VAR) or {}
    
    current_time = time.time()
    last_save_time = last_saves.get(doc_path, 0)

    # Initialize last save time if not present
    if last_save_time == 0:
        last_saves[doc_path] = current_time
        script.set_envvar(LAST_SAVES_VAR, last_saves)
        return

    elapsed_minutes = (current_time - last_save_time) / 60.0

    if elapsed_minutes >= interval:
        try:
            # Perform save operation
            doc.Save()
            # Update last save time
            last_saves[doc_path] = current_time
            script.set_envvar(LAST_SAVES_VAR, last_saves)
            # Show visual feedback
            forms.toast("Project auto-saved successfully!", title="AutoSave")
        except Exception as e:
            # Silence error or log to output if needed
            print("AutoSave failed: {}".format(e))


if __name__ == '__main__':
    run()
