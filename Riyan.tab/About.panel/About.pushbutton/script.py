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
        Width="450" Height="480"
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
            <Border Grid.Row="0" Background="#111111" BorderBrush="#7B2C2C" BorderThickness="0,0,0,2" Padding="20,10">
                <StackPanel HorizontalAlignment="Center">
                    <Image x:Name="BigLogo" Height="80" Width="180" 
                           Margin="0,0,0,5"
                           RenderOptions.BitmapScalingMode="HighQuality"/>
                    <TextBlock Text="Riyan Revit Plugin Suite"
                               FontSize="20" FontWeight="Bold"
                               Foreground="White" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>

            <!-- Content Area -->
            <StackPanel Grid.Row="1" Margin="30,20,30,20" VerticalAlignment="Top">
                <Border BorderBrush="#222222" BorderThickness="0,0,0,1" Margin="0,0,0,15" Padding="0,0,0,10">
                    <Grid>
                        <TextBlock Text="Version" Foreground="#888888" HorizontalAlignment="Left"/>
                        <TextBlock Text="v{version}" Foreground="White" FontWeight="Bold" HorizontalAlignment="Right"/>
                    </Grid>
                </Border>

                <Border BorderBrush="#222222" BorderThickness="0,0,0,1" Margin="0,0,0,15" Padding="0,0,0,10">
                    <Grid>
                        <TextBlock Text="Developer" Foreground="#888888" HorizontalAlignment="Left"/>
                        <TextBlock Text="Udarie &amp; Chalana" Foreground="White" FontWeight="Bold" HorizontalAlignment="Right"/>
                    </Grid>
                </Border>

                <TextBlock Text="Professional Revit automation tools for link management and coordination." 
                           Foreground="#A0A0A0" TextWrapping="Wrap" Margin="0,0,0,25" TextAlignment="Center" FontStyle="Italic"/>

                <Button x:Name="UpdateBtn" Content="Check for Updates" Margin="0,0,0,10" Cursor="Hand" Background="#1A1A1A" BorderBrush="#7B2C2C" BorderThickness="1" Foreground="White" Padding="10">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Button.Template>
                </Button>

                <Button x:Name="CloseBtn" Content="Close" Background="#7B2C2C" Foreground="White" FontWeight="Bold"
                        Padding="15,8" HorizontalAlignment="Center" Margin="0,20,0,0" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" CornerRadius="20" Padding="{TemplateBinding Padding}">
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
            # Escape strings for XML to prevent Xaml parsing errors
            import cgi
            safe_title = cgi.escape(str(title))
            safe_message = cgi.escape(str(message))
            
            msg_xaml = """
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    Title="{title}" Width="350" Height="180" 
                    WindowStartupLocation="CenterScreen" Background="#111111" Topmost="True">
                <Border BorderBrush="#7B2C2C" BorderThickness="2">
                    <StackPanel VerticalAlignment="Center" Margin="20">
                        <TextBlock Text="{title}" Foreground="#7B2C2C" FontWeight="Bold" FontSize="16" Margin="0,0,0,10" HorizontalAlignment="Center"/>
                        <TextBlock Text="{message}" Foreground="White" FontSize="14" TextWrapping="Wrap" HorizontalAlignment="Center" TextAlignment="Center"/>
                        <Button x:Name="OkBtn" Content="OK" Margin="0,20,0,0" Width="80" Height="30" Background="#7B2C2C" Foreground="White" FontWeight="Bold" Cursor="Hand">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" CornerRadius="15">
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
        except Exception as e:
            # Fallback to standard alert if custom XAML fails
            forms.alert(str(message), title=str(title))

    update_btn = window.FindName("UpdateBtn")

    def on_update(sender, args):
        try:
            update_btn.IsEnabled = False
            update_btn.Content = "Checking..."
            
            # Use Dispatcher to force UI update so "Checking..." is seen
            from System.Windows.Threading import DispatcherPriority
            from System import Action
            try:
                window.Dispatcher.Invoke(DispatcherPriority.Background, Action(lambda: None))
            except: pass

            try:
                Net.ServicePointManager.SecurityProtocol |= Net.SecurityProtocolType.Tls12
            except: pass
            
            client = WebClientWithTimeout()
            client.Headers.Add("Cache-Control", "no-cache")
            
            url = "https://raw.githubusercontent.com/udarieimalsha/Riyan.extension/main/update.json"
            json_str = client.DownloadString(url)
            data = json.loads(json_str)
            
            remote_v = data.get("version", "")
            dl_url = data.get("download_url", "")
            
            def v_to_tuple(v): 
                try: return tuple(map(int, str(v).split('.')))
                except: return (0,0,0)
            
            if v_to_tuple(remote_v) > v_to_tuple(VERSION):
                res = forms.alert("A new version (%s) is available!\n\nWould you like to download it?" % remote_v, 
                                  title="Update Available", yes=True, no=True)
                if res:
                    webbrowser.open(dl_url)
                    window.Close()
            else:
                forms.alert("You are up to date! (v%s)" % VERSION, title="Riyan Tool", warn_icon=False)
                
        except Exception as e:
            forms.alert("Update Check Failed: " + str(e), title="Error")
        finally:
            update_btn.Content = "Check for Updates"
            update_btn.IsEnabled = True

    update_btn.Click += on_update


    window.ShowDialog()

if __name__ == "__main__":
    show_about_dialog()
