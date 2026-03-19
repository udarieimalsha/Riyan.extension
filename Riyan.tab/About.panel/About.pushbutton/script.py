# -*- coding: utf-8 -*-
import os
import System
import subprocess
from pyrevit import forms
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

VERSION = "1.0.1"

def show_about_dialog():
    plugin_dir = os.path.dirname(__file__)
    # Find logo in the other button folder (standard practice in pyRevit suites)
    # Riyan.tab/Link Tools.panel/Copy from Link.pushbutton/logo.png
    logo_path = os.path.join(os.path.dirname(os.path.dirname(plugin_dir)), 
                             "Link Tools.panel", 
                             "Copy from Link.pushbutton", 
                             "logo.png")

    xaml_str = """
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="About Riyan Plugin"
        Width="450" Height="420"
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
                    <Image x:Name="BigLogo" Height="100" Width="180" 
                           Margin="0,0,0,10"
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
                        <TextBlock Text="Udarie Imalsha" Foreground="White" FontWeight="Bold" HorizontalAlignment="Right"/>
                    </Grid>
                </Border>

                <TextBlock Text="Professional Revit automation tools for link management and coordination." 
                           Foreground="#A0A0A0" TextWrapping="Wrap" Margin="0,0,0,25" TextAlignment="Center" FontStyle="Italic"/>

                <Button x:Name="UpdateBtn" Content="Check for Updates" Margin="0,0,0,10" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="Transparent" BorderBrush="#7B2C2C" BorderThickness="1" CornerRadius="5" Padding="10">
                                <TextBlock Text="{TemplateBinding Content}" Foreground="#7B2C2C" HorizontalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Button.Template>
                </Button>

                <Button x:Name="RepoBtn" Content="Visit GitHub Repository" Margin="0,0,0,10" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="Transparent" BorderBrush="#7B2C2C" BorderThickness="1" CornerRadius="5" Padding="10">
                                <TextBlock Text="{TemplateBinding Content}" Foreground="#7B2C2C" HorizontalAlignment="Center"/>
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
    def on_close(sender, args):
        window.Close()
    close_btn.Click += on_close

    repo_btn = window.FindName("RepoBtn")
    def on_repo(sender, args):
        import webbrowser
        webbrowser.open("https://github.com/udarieimalsha/Riyan.extension")
    repo_btn.Click += on_repo

    update_btn = window.FindName("UpdateBtn")
    def on_update(sender, args):
        try:
            # Run git pull in the plugin directory
            process = subprocess.Popen(['git', 'pull', 'origin', 'main'], 
                                     cwd=plugin_dir,
                                     stdout=subprocess.PIPE, 
                                     stderr=subprocess.PIPE,
                                     shell=True)
            stdout, stderr = process.communicate()
            
            if process.returncode == 0:
                if "Already up to date" in stdout:
                    forms.alert("Your plugin is already up to date!", title="Update Check")
                else:
                    forms.alert("Successfully updated to the latest version!\n\nPlease Reload pyRevit to apply changes.", 
                                title="Update Success")
            else:
                forms.alert("Failed to update: " + stderr, title="Update Error")
        except Exception as e:
            forms.alert("Error running update: " + str(e), title="Update Error")
            
    update_btn.Click += on_update

    window.ShowDialog()

if __name__ == "__main__":
    show_about_dialog()
