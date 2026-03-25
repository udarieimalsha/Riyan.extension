# -*- coding: utf-8 -*-
"""
Auto-updater script for Riyan.extension.
Runs silently when Revit starts up.
- For Git collaborators: Performs a git pull to fetch latest changes.
- For EXE users: Checks GitHub for a new version and shows an update popup.
"""

import os
import subprocess
import threading
import clr
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

def run_git_pull_update(extension_dir):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        process = subprocess.Popen(
            ["git", "pull"],
            cwd=extension_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo
        )
        process.communicate()
    except:
        pass

def show_updater_window(version, url):
    def thread_proc():
        import System.Windows as WPF
        from System.Windows.Markup import XamlReader
        from System import Uri
        import System.Net as Net
        
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
                        A new version (vVERSION_STR) is available! Click Update Now to download and install it automatically.
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
        
        # Action delegate wrapper for IronPython 2.7
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
            client = Net.WebClient()
            Net.ServicePointManager.SecurityProtocol |= Net.SecurityProtocolType.Tls12
            
            def set_prog(val): prog_bar.Value = val
            def on_progress(sender, ev):
                window.Dispatcher.Invoke(Action[int](set_prog), ev.ProgressPercentage)
                
            def finish_dl():
                window.Close()
                import subprocess
                subprocess.Popen([temp_path, '/SILENT'])
                
            def on_complete(sender, ev):
                if ev.Error is None and not ev.Cancelled:
                    window.Dispatcher.Invoke(Action(finish_dl))
                else:
                    window.Dispatcher.Invoke(Action(window.Close))
                    
            client.DownloadProgressChanged += Net.DownloadProgressChangedEventHandler(on_progress)
            client.DownloadFileCompleted += Net.ComponentModel.AsyncCompletedEventHandler(on_complete)
            
        btn_update.Click += start_update
        window.ShowDialog()
        
    # Run the window synchronously so PyRevit engine does not collect the scope
    thread_proc()

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
        Net.ServicePointManager.SecurityProtocol |= Net.SecurityProtocolType.Tls12
        json_str = client.DownloadString("https://raw.githubusercontent.com/udarieimalsha/Riyan.extension/main/update.json")
        data = json.loads(json_str)
        
        remote_v = data.get("version", "")
        dl_url = data.get("download_url", "")
        if not remote_v or not dl_url: return
        
        def v_to_tuple(v): return tuple(map(int, v.split('.')))
        
        if v_to_tuple(remote_v) > v_to_tuple(local_v):
            show_updater_window(remote_v, dl_url)
    except Exception as e:
        import clr
        clr.AddReference('System.Windows.Forms')
        import System.Windows.Forms as Forms
        Forms.MessageBox.Show("Network/Update Error: " + str(e))

def main():
    curr_dir = os.path.dirname(__file__)
    extension_dir = os.path.dirname(curr_dir)
    git_dir = os.path.join(extension_dir, '.git')
    
    if os.path.exists(git_dir):
        run_git_pull_update(extension_dir)
    else:
        # Check updates synchronously so the window can block Revit startup
        run_exe_update_checker(extension_dir)

try:
    main()
except:
    pass
