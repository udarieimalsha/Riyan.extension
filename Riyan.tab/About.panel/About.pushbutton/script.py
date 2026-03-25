# -*- coding: utf-8 -*-
import os
import System
import subprocess
from pyrevit import forms
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
import clr
clr.AddReference("System")
import System.Net as Net
import json
import webbrowser

# Custom WebClient with Timeout
class WebClientWithTimeout(Net.WebClient):
    def GetWebRequest(self, address):
        wr = Net.WebClient.GetWebRequest(self, address)
        wr.Timeout = 5000 # 5 seconds
        return wr

def get_version():
    try:
        # Search for version.txt in parent directories
        curr = os.path.dirname(__file__)
        for _ in range(5): # Check up to 5 levels up
            v_file = os.path.join(curr, "version.txt")
            if os.path.exists(v_file):
                with open(v_file, "r") as f:
                    return f.read().strip()
            curr = os.path.dirname(curr)
    except:
        pass
    return "1.0.5" # Default fallback for this version

VERSION = get_version()

def show_about_dialog():
    plugin_dir = os.path.dirname(__file__)
    # Robust logo path
    curr = os.path.dirname(__file__)
    ext_dir = None
    for _ in range(5):
        if os.path.exists(os.path.join(curr, "version.txt")):
            ext_dir = curr
            break
        curr = os.path.dirname(curr)
    
    if ext_dir:
        logo_path = os.path.join(ext_dir, "Riyan.tab", "Coordination.panel", "Copy from Link.pushbutton", "logo.png")
    else:
        logo_path = os.path.join(plugin_dir, "icon.png")

    xaml_str = """
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="About Riyan Plugin"
        Width="450" Height="550"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        FontFamily="Segoe UI"
        Background="Black"
        WindowStyle="ToolWindow">

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Header bar -->
            <Border Grid.Row="0" Background="#111111" BorderBrush="#7B2C2C" BorderThickness="0,0,0,2" Padding="20,15">
                <StackPanel HorizontalAlignment="Center">
                    <Image x:Name="BigLogo" Height="90" Width="200" 
                           Margin="0,0,0,10"
                           RenderOptions.BitmapScalingMode="HighQuality"/>
                    <TextBlock Text="Riyan Revit Plugin Suite"
                               FontSize="22" FontWeight="Bold"
                               Foreground="White" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>

            <!-- Content Area -->
            <StackPanel Grid.Row="1" Margin="35,25,35,25" VerticalAlignment="Top">
                <Border BorderBrush="#222222" BorderThickness="0,0,0,1" Margin="0,0,0,15" Padding="0,0,0,10">
                    <Grid>
                        <TextBlock Text="Version" Foreground="#888888" HorizontalAlignment="Left" FontSize="14"/>
                        <TextBlock Text="v{version}" Foreground="White" FontWeight="Bold" HorizontalAlignment="Right" FontSize="14"/>
                    </Grid>
                </Border>

                <Border BorderBrush="#222222" BorderThickness="0,0,0,1" Margin="0,0,0,15" Padding="0,0,0,10">
                    <Grid>
                        <TextBlock Text="Developer" Foreground="#888888" HorizontalAlignment="Left" FontSize="14"/>
                        <TextBlock Text="Udarie &amp; Chalana" Foreground="White" FontWeight="Bold" HorizontalAlignment="Right" FontSize="14"/>
                    </Grid>
                </Border>

                <TextBlock Text="Professional Revit automation tools for link management and coordination." 
                           Foreground="#A0A0A0" TextWrapping="Wrap" Margin="0,10,0,30" TextAlignment="Center" FontStyle="Italic" FontSize="13"/>

                <Button x:Name="UpdateBtn" Content="Check for Updates" Margin="0,0,0,15" Cursor="Hand" Background="#1A1A1A" BorderBrush="#7B2C2C" BorderThickness="1" Foreground="White" Padding="12">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                
                <!-- Update Status Area -->
                <Border x:Name="StatusBorder" Background="#337B2C2C" BorderBrush="#7B2C2C" BorderThickness="1" CornerRadius="5" Padding="12" Margin="0,0,0,10" Visibility="Collapsed">
                    <StackPanel>
                        <TextBlock x:Name="StatusText" Foreground="White" TextAlignment="Center" TextWrapping="Wrap" FontSize="14" FontWeight="SemiBold"/>
                        <ProgressBar x:Name="DownloadBar" Height="8" Margin="0,10,0,0" Foreground="#7B2C2C" Background="#222222" BorderThickness="0" Visibility="Collapsed"/>
                    </StackPanel>
                </Border>

                <Button x:Name="CloseBtn" Content="Close" Background="#7B2C2C" Foreground="White" FontWeight="Bold"
                        Padding="25,10" HorizontalAlignment="Center" Margin="0,30,0,0" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" CornerRadius="25" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </StackPanel>
        </Grid>
    </Window>
    """.replace("{version}", VERSION)

    window = XamlReader.Parse(xaml_str)
    
    # Load Logo
    logo_img = window.FindName("BigLogo")
    if os.path.exists(logo_path):
        uri = Uri(logo_path, UriKind.Absolute)
        logo_img.Source = BitmapImage(uri)
    
    # Setup Buttons
    close_btn = window.FindName("CloseBtn")
    close_btn.Click += lambda s, e: window.Close()

    def show_branded_message(title, message):
        try:
            # Manual XML escaping for robustness
            safe_title = str(title).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            safe_message = str(message).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "&#x0a;")
            
            msg_xaml = """
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    Title="{title}" Width="350" Height="220" 
                    WindowStartupLocation="CenterScreen"
                    Background="Black" WindowStyle="ToolWindow" Topmost="True">
                <Border BorderBrush="#7B2C2C" BorderThickness="3" Margin="5">
                    <StackPanel VerticalAlignment="Center" Margin="20">
                        <TextBlock Text="{title}" Foreground="#7B2C2C" FontWeight="Bold" FontSize="20" Margin="0,0,0,15" HorizontalAlignment="Center"/>
                        <TextBlock Text="{message}" Foreground="White" FontSize="15" TextWrapping="Wrap" HorizontalAlignment="Center" TextAlignment="Center"/>
                        <Button x:Name="OkBtn" Content="OK" Margin="0,25,0,0" Width="90" Height="35" Background="#7B2C2C" Foreground="White" FontWeight="Bold" Cursor="Hand">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" CornerRadius="5">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                    </StackPanel>
                </Border>
            </Window>
            """.replace("{title}", safe_title).replace("{message}", safe_message)
            msg_win = XamlReader.Parse(msg_xaml)
            msg_win.FindName("OkBtn").Click += lambda s, e: msg_win.Close()
            msg_win.ShowDialog()
        except:
            forms.alert(str(message), title=str(title))

    update_btn = window.FindName("UpdateBtn")
    status_border = window.FindName("StatusBorder")
    status_text = window.FindName("StatusText")
    download_bar = window.FindName("DownloadBar")

    def on_update(sender, args):
        try:
            status_border.Visibility = System.Windows.Visibility.Collapsed
            download_bar.Visibility = System.Windows.Visibility.Collapsed
            update_btn.IsEnabled = False
            update_btn.Content = "Checking..."
            
            # Force UI update
            from System.Windows.Threading import DispatcherPriority
            from System import Action
            window.Dispatcher.Invoke(DispatcherPriority.Background, Action(lambda: None))

            Net.ServicePointManager.SecurityProtocol |= Net.SecurityProtocolType.Tls12
            
            client = WebClientWithTimeout()
            client.Headers.Add("Cache-Control", "no-cache")
            
            # Use cache buster for more reliable checking
            import time
            url = "https://raw.githubusercontent.com/udarieimalsha/Riyan.extension/main/update.json?v=" + str(int(time.time()))
            json_str = client.DownloadString(url)
            data = json.loads(json_str)
            
            # Detect if this is a Git-based installation
            curr = os.path.dirname(__file__)
            ext_dir = None
            for _ in range(5):
                if os.path.exists(os.path.join(curr, "version.txt")):
                    ext_dir = curr
                    break
                curr = os.path.dirname(curr)
            
            is_git = False
            if ext_dir and os.path.exists(os.path.join(ext_dir, ".git")):
                is_git = True

            if is_git:
                update_btn.Content = "Pulling Git..."
                try:
                    # Sync UI
                    window.Dispatcher.Invoke(DispatcherPriority.Background, Action(lambda: None))
                    
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    env = os.environ.copy()
                    subprocess.call(["git", "stash"], cwd=ext_dir, startupinfo=startupinfo, env=env)
                    res = subprocess.call(["git", "pull"], cwd=ext_dir, startupinfo=startupinfo, env=env)
                    subprocess.call(["git", "stash", "pop"], cwd=ext_dir, startupinfo=startupinfo, env=env)
                    
                    if res == 0:
                        status_text.Text = "Successfully updated via Git! Please restart Revit."
                    else:
                        status_text.Text = "Git update failed. Please check for conflicts manually."
                    status_border.Visibility = System.Windows.Visibility.Visible
                except Exception as ex:
                    status_text.Text = "Git Error: " + str(ex)
                    status_border.Visibility = System.Windows.Visibility.Visible
                return

            remote_v = data.get("version", "")
            dl_url = data.get("download_url", "")
            
            def v_to_tuple(v): 
                try: return tuple(map(int, str(v).split('.')))
                except: return (0,0,0)
            
            if v_to_tuple(remote_v) > v_to_tuple(VERSION):
                update_btn.Content = "Update Available!"
                res = forms.alert("A new version (%s) is available!\n\nWould you like to install it now?" % remote_v, 
                                  title="Update Available", yes=True, no=True)
                if res:
                    # Direct Download Mode
                    status_border.Visibility = System.Windows.Visibility.Visible
                    download_bar.Visibility = System.Windows.Visibility.Visible
                    status_text.Text = "Downloading v%s..." % remote_v
                    
                    temp_exe = os.path.join(os.environ["TEMP"], "RiyanSetup_Latest.exe")
                    
                    # Async Download for UI responsiveness
                    def dl_progress(s, e):
                        window.Dispatcher.Invoke(Action(lambda: setattr(download_bar, "Value", e.ProgressPercentage)))
                    
                    def dl_complete(s, e):
                        if e.Error:
                            window.Dispatcher.Invoke(Action(lambda: forms.alert("Download failed: " + str(e.Error))))
                        else:
                            # Use PowerShell to wait for Revit and then run installer
                            try:
                                ps_cmd = (
                                    "powershell -Command \"while (Get-Process Revit -ErrorAction SilentlyContinue) "
                                    "{ Start-Sleep -Seconds 2 }; "
                                    "Start-Process '{0}' -ArgumentList '/SILENT','/SUPPRESSMSGBOXES','/NORESTART'\""
                                ).format(temp_exe)
                                
                                import subprocess
                                si = subprocess.STARTUPINFO()
                                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                                si.wShowWindow = 0 # SW_HIDE
                                subprocess.Popen(ps_cmd, startupinfo=si)
                                
                                from System import Action
                                window.Dispatcher.Invoke(Action(window.Close))
                                forms.alert("Update downloaded! It will install automatically as soon as you close Revit.", title="Riyan Update")
                            except Exception as ex:
                                from System import Action
                                window.Dispatcher.Invoke(Action(window.Close))
                                forms.alert("Ready to install. Please close Revit and run manually:\n" + temp_exe, title="Riyan Update")
                    
                    dl_client = Net.WebClient()
                    dl_client.DownloadProgressChanged += dl_progress
                    dl_client.DownloadFileCompleted += dl_complete
                    dl_client.DownloadFileAsync(Uri(dl_url), temp_exe)
                    return # Exit so we don't reset button state yet
            else:
                status_text.Text = "You are up to date! (v%s)" % VERSION
                status_border.Visibility = System.Windows.Visibility.Visible
                
        except Exception as e:
            status_text.Text = "Check Failed: " + str(e)
            status_border.Visibility = System.Windows.Visibility.Visible
        finally:
            if update_btn.Content == "Checking...":
                update_btn.Content = "Check for Updates"
            if download_bar.Visibility != System.Windows.Visibility.Visible:
                update_btn.IsEnabled = True

    update_btn.Click += on_update


    window.ShowDialog()

if __name__ == "__main__":
    show_about_dialog()
