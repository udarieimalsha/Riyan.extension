# -*- coding: utf-8 -*-
<<<<<<< Updated upstream
"""
Auto-updater script for Riyan.extension.
Handles:
- Git pull for developers (if .git folder exists)
- Silent EXE update check for end-users
"""

import os
import subprocess
import threading
import clr
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
=======
import os
import subprocess
import datetime
import traceback
>>>>>>> Stashed changes

def run_git_pull_update(extension_dir):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        # Stash local changes before pulling
        subprocess.call(["git", "stash"], cwd=extension_dir, startupinfo=startupinfo)
        subprocess.call(["git", "pull"], cwd=extension_dir, startupinfo=startupinfo)
        subprocess.call(["git", "stash", "pop"], cwd=extension_dir, startupinfo=startupinfo)
    except:
        pass

def show_updater_window(version, url):
    def thread_proc():
        import System.Windows as WPF
        from System.Windows.Markup import XamlReader
        from System import Uri
        import System.Net as Net
        import System.ComponentModel
        XAML = """
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Update Available" Width="400" Height="220" 
                WindowStartupLocation="CenterScreen" WindowStyle="None" 
                AllowsTransparency="True" Background="Transparent"
                Topmost="True">
            <Border Background="#000000" BorderBrush="#802F2D" BorderThickness="2" CornerRadius="10">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Riyan Extension Update" Foreground="#802F2D" FontSize="18" FontWeight="Bold" Grid.Row="0"/>
                    
                    <TextBlock x:Name="MsgTxt" Foreground="White" FontSize="13" TextWrapping="Wrap" Margin="0,15,0,0" Grid.Row="1">
                        A new version (vVERSION_STR) is available! Click Update Now to install it seamlessly.
                    </TextBlock>
                    
                    <ProgressBar x:Name="ProgBar" Height="10" Margin="0,15,0,0" Grid.Row="2" Visibility="Hidden" Foreground="#802F2D" Background="#333333" BorderThickness="0"/>
                    
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Grid.Row="3" Margin="0,15,0,0">
                        <Button x:Name="Btn_Cancel" Content="Later" Width="80" Height="30" Background="#333333" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                        <Button x:Name="Btn_Update" Content="Update Now" Width="100" Height="30" Background="#802F2D" Foreground="White" BorderThickness="0"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Window>
        """.replace("VERSION_STR", version)
        
        window = XamlReader.Parse(XAML)
        btn_update = window.FindName("Btn_Update")
        btn_cancel = window.FindName("Btn_Cancel")
        prog_bar = window.FindName("ProgBar")
        msg_txt = window.FindName("MsgTxt")
        
        from System import Action
        
        def close_win(s, e): window.Close()
        btn_cancel.Click += close_win
        
        state = {"downloading": False}
        
        def start_update(s, e):
            if state["downloading"]: return
            state["downloading"] = True
            
            btn_update.IsEnabled = False
            btn_cancel.IsEnabled = False
            prog_bar.Visibility = WPF.Visibility.Visible
            msg_txt.Text = "Downloading update... Please wait."
            
            temp_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "RiyanSetup_Update.exe")
            from System.Net import ServicePointManager, SecurityProtocolType, WebClient, Uri
            client = WebClient()
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12
            
            def set_prog(val): prog_bar.Value = val
            def on_progress(sender, ev):
                window.Dispatcher.Invoke(Action[int](set_prog), ev.ProgressPercentage)
                
            def finish_dl():
                window.Close()
                import subprocess
                # Seamless Silent Install
                subprocess.Popen([temp_path, '/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART'])
                
            def on_complete(sender, ev):
                if ev.Error is None and not ev.Cancelled:
                    window.Dispatcher.Invoke(Action(finish_dl))
                else:
                    window.Dispatcher.Invoke(Action(window.Close))
                    
            client.DownloadProgressChanged += on_progress
            client.DownloadFileCompleted += on_complete
            client.DownloadFileAsync(Uri(url), temp_path)
            
        btn_update.Click += start_update
        window.ShowDialog()
        
    import System.Threading as Threading
    t = Threading.Thread(Threading.ThreadStart(thread_proc))
    t.SetApartmentState(Threading.ApartmentState.STA)
    t.Start()

def run_exe_update_checker(extension_dir):
    try:
        import System.Net as Net
        import json
        
        v_file = os.path.join(extension_dir, 'version.txt')
        if not os.path.exists(v_file): return
        with open(v_file, 'r') as f:
            local_v = f.read().strip()
            
        client = Net.WebClient()
        client.Headers.Add("Cache-Control", "no-cache")
        import System.Net
        System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12
        import time
        cache_buster = "?v=" + str(int(time.time()))
        json_str = client.DownloadString("https://raw.githubusercontent.com/udarieimalsha/Riyan.extension/main/update.json" + cache_buster)
        data = json.loads(json_str)
        
        remote_v = data.get("version", "")
        dl_url = data.get("download_url", "")
        if not remote_v or not dl_url: return
        
        def v_to_tuple(v): 
            try: return tuple(map(int, str(v).split('.')))
            except: return (0,0,0)
        
        if v_to_tuple(remote_v) > v_to_tuple(local_v):
            show_updater_window(remote_v, dl_url)
    except:
        pass

def main():
    extension_dir = os.path.dirname(__file__)
    git_dir = os.path.join(extension_dir, '.git')
    log_file = os.path.join(extension_dir, 'update_log.txt')
    
<<<<<<< Updated upstream
    if not os.path.exists(git_dir):
        run_exe_update_checker(extension_dir)
    else:
        # For developers, just do a silent pull in background
        import threading
        t = threading.Thread(target=run_git_pull_update, args=(extension_dir,))
        t.start()

if __name__ == '__main__':
    main()
=======
    with open(log_file, "a") as f:
        f.write("\n\n--- Update triggered at {} ---\n".format(datetime.datetime.now()))
        
        if not os.path.exists(git_dir):
            f.write("No .git folder found. Aborting.\n")
            return
            
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            # Setup environment variables to ensure git finds its path correctly
            envs = os.environ.copy()
            
            f.write("Running git stash...\n")
            p1 = subprocess.Popen(["git", "stash"], cwd=extension_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo, env=envs)
            out1, err1 = p1.communicate()
            f.write("Stash Output:\n{}\nError:\n{}\n".format(out1, err1))
            
            f.write("Running git pull...\n")
            p2 = subprocess.Popen(["git", "pull"], cwd=extension_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo, env=envs)
            out2, err2 = p2.communicate()
            f.write("Pull Output:\n{}\nError:\n{}\n".format(out2, err2))
            
            f.write("Running git stash pop...\n")
            p3 = subprocess.Popen(["git", "stash", "pop"], cwd=extension_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo, env=envs)
            out3, err3 = p3.communicate()
            f.write("Pop Output:\n{}\nError:\n{}\n".format(out3, err3))
            
        except Exception as e:
            f.write("Exception occurred:\n{}\n".format(traceback.format_exc()))

try:
    run_auto_update()
except Exception as e:
    pass
>>>>>>> Stashed changes
