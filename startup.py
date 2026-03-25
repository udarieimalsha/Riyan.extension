# -*- coding: utf-8 -*-
"""
Auto-updater script for Riyan.extension.
Executed by PyRevit when loading the extension.
"""

import os
import subprocess

def run_auto_update():
    extension_dir = os.path.dirname(__file__)
    git_dir = os.path.join(extension_dir, '.git')
    
    if not os.path.exists(git_dir):
        return
        
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        # Temporarily stash any local unsaved user changes (like local test scripts)
        subprocess.call(["git", "stash"], cwd=extension_dir, startupinfo=startupinfo)
        # Pull the absolute latest code from GitHub
        subprocess.call(["git", "pull"], cwd=extension_dir, startupinfo=startupinfo)
        # Restore any local changes that were stashed
        subprocess.call(["git", "stash", "pop"], cwd=extension_dir, startupinfo=startupinfo)
    except Exception as e:
        pass

try:
    if __name__ == '__main__':
        run_auto_update()
    else:
        run_auto_update()
except:
    pass
