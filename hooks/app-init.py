# -*- coding: utf-8 -*-
"""
Auto-updater script for Riyan.extension.
Executed by pyRevit at startup.
Non-blocking version for Revit 2025.
"""

import os
import subprocess
import threading
import time

def run_git_pull_update(extension_dir):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        env = os.environ.copy()
        subprocess.call(["git", "stash"], cwd=extension_dir, startupinfo=startupinfo, env=env)
        subprocess.call(["git", "pull"], cwd=extension_dir, startupinfo=startupinfo, env=env)
        subprocess.call(["git", "stash", "pop"], cwd=extension_dir, startupinfo=startupinfo, env=env)
    except:
        pass

def notify_update(version):
    # Use pyRevit toast if possible, it's safer than WPF during startup
    try:
        from pyrevit import forms
        forms.toast(
            "Riyan Update v{} is available!".format(version),
            title="Riyan Extension",
            click_command="About"
        )
    except:
        pass

def run_exe_update_checker(extension_dir):
    try:
        import System.Net as Net
        import json
        
        v_file = os.path.join(extension_dir, 'version.txt')
        if not os.path.exists(v_file): return
        with open(v_file, 'r') as f:
            local_v = f.read().strip()
            
        # Add delay to not interfere with Revit startup splash
        time.sleep(5)
        
        client = Net.WebClient()
        client.Headers.Add("Cache-Control", "no-cache")
        import System.Net
        System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12
        cache_buster = "?v=" + str(int(time.time()))
        json_str = client.DownloadString("https://raw.githubusercontent.com/udarieimalsha/Riyan.extension/main/update.json" + cache_buster)
        data = json.loads(json_str)
        
        remote_v = data.get("version", "")
        if not remote_v: return
        
        def v_to_tuple(v): 
            try: return tuple(map(int, str(v).split('.')))
            except: return (0,0,0)
        
        if v_to_tuple(remote_v) > v_to_tuple(local_v):
            notify_update(remote_v)
    except:
        pass

def main():
    hooks_dir = os.path.dirname(__file__)
    extension_dir = os.path.dirname(hooks_dir)
    git_dir = os.path.join(extension_dir, '.git')
    
    if os.path.exists(git_dir):
        # Update thread for git users
        t = threading.Thread(target=run_git_pull_update, args=(extension_dir,))
        t.daemon = True
        t.start()
    else:
        # Update thread for end users (EXE)
        t = threading.Thread(target=run_exe_update_checker, args=(extension_dir,))
        t.daemon = True
        t.start()

if __name__ == '__main__':
    main()
