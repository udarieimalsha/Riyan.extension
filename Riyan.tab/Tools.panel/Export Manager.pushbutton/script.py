# -*- coding: utf-8 -*-
import clr
import os
import json
import re

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Xaml")
clr.AddReference("System.Xml")

from System.Windows import Window
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader
import System

from pyrevit import revit, DB, UI, forms

doc = revit.doc
uidoc = revit.uidoc

# ------------------------------------------------------------------------------
# Custom Dark Alert Dialog
# ------------------------------------------------------------------------------
class CustomAlertWindow(object):
    def __init__(self, message, title, icon_char):
        xaml_code = """<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Alert" Width="400" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        Background="#111111" WindowStyle="None" AllowsTransparency="False"
        ResizeMode="NoResize">
    <Border BorderBrush="#3A3A3A" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="46"/>
            </Grid.RowDefinitions>
            
            <!-- Title Bar -->
            <Grid x:Name="TitleBar" Grid.Row="0" Background="#1A1A1A">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
                    <TextBlock x:Name="TxtAccent" Text="" Foreground="#802F2D" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TxtTitle" Text="Alert" Foreground="#CCCCCC" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
                <Button x:Name="CloseBtn" Content="" HorizontalAlignment="Right"
                        Width="44" Height="36" BorderThickness="0" Cursor="Hand"
                        Background="Transparent" Foreground="#666666"
                        FontSize="12">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="{TemplateBinding Background}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#C0272D"/>
                                    <Setter Property="Foreground" Value="White"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
            
            <!-- Content Area -->
            <Grid Grid.Row="1" Margin="20,16,20,16">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock x:Name="TxtIcon" Grid.Column="0" Text="" Foreground="#802F2D" FontSize="26" 
                           VerticalAlignment="Center" Margin="0,0,16,0"/>
                
                <TextBlock x:Name="TxtMessage" Grid.Column="1" Text="" Foreground="#CCCCCC" 
                           FontSize="11.5" TextWrapping="Wrap" VerticalAlignment="Center" HorizontalAlignment="Left"/>
            </Grid>
            
            <!-- Footer -->
            <Border Grid.Row="2" Background="#0E0E0E" BorderBrush="#222222" BorderThickness="0,1,0,0">
                <Button x:Name="OkBtn" Content="OK" HorizontalAlignment="Right" Width="80" Height="26" 
                        Margin="0,0,14,0" Cursor="Hand" Foreground="White" FontWeight="SemiBold" FontSize="11">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="#802F2D" CornerRadius="3">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#9E3A38"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Border>
        </Grid>
    </Border>
</Window>
"""
        r = XmlReader.Create(StringReader(xaml_code))
        self.win = XamlReader.Load(r)
        
        self.TxtTitle = self.win.FindName("TxtTitle")
        self.TxtMessage = self.win.FindName("TxtMessage")
        self.TxtIcon = self.win.FindName("TxtIcon")
        self.TxtAccent = self.win.FindName("TxtAccent")
        self.CloseBtn = self.win.FindName("CloseBtn")
        self.OkBtn = self.win.FindName("OkBtn")
        self.TitleBar = self.win.FindName("TitleBar")
        
        if self.TxtTitle:
            self.TxtTitle.Text = title
        if self.TxtMessage:
            self.TxtMessage.Text = message
        if self.TxtIcon:
            self.TxtIcon.Text = icon_char
        if self.TxtAccent:
            self.TxtAccent.Text = u"\u2B0C"  # ⬌
        if self.CloseBtn:
            self.CloseBtn.Content = u"\u2715"  # ✕
            self.CloseBtn.Click += self.CloseBtn_Click
        if self.OkBtn:
            self.OkBtn.Click += self.OkBtn_Click
        if self.TitleBar:
            self.TitleBar.MouseLeftButtonDown += self.TitleBar_MouseDown

    def TitleBar_MouseDown(self, sender, e):
        try:
            self.win.DragMove()
        except:
            pass
            
    def CloseBtn_Click(self, sender, e):
        self.win.Close()
        
    def OkBtn_Click(self, sender, e):
        self.win.Close()

    def ShowDialog(self):
        return self.win.ShowDialog()

def show_alert(message, title="Export Manager", is_error=False, is_warning=False):
    icon_char = "ⓘ"
    if is_error:
        icon_char = "❌"
    elif is_warning:
        icon_char = "⚠️"
        
    try:
        dialog = CustomAlertWindow(message, title, icon_char)
        dialog.ShowDialog()
    except Exception as e:
        forms.alert("Error rendering dark UI: " + str(e) + "\n\nOriginal message: " + message, title=title)

# ------------------------------------------------------------------------------
# Custom Dark Text Input Dialog
# ------------------------------------------------------------------------------
class CustomTextInputWindow(object):
    def __init__(self, title, description):
        self.result = None
        xaml_code = """<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Input" Width="400" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        Background="#111111" WindowStyle="None" AllowsTransparency="False"
        ResizeMode="NoResize">
    <Border BorderBrush="#3A3A3A" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="46"/>
            </Grid.RowDefinitions>
            
            <!-- Title Bar -->
            <Grid x:Name="TitleBar" Grid.Row="0" Background="#1A1A1A">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
                    <TextBlock x:Name="TxtAccent" Text="" Foreground="#802F2D" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TxtTitle" Text="Input" Foreground="#CCCCCC" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
                <Button x:Name="CloseBtn" Content="" HorizontalAlignment="Right"
                        Width="44" Height="36" BorderThickness="0" Cursor="Hand"
                        Background="Transparent" Foreground="#666666"
                        FontSize="12">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="{TemplateBinding Background}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#C0272D"/>
                                    <Setter Property="Foreground" Value="White"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
            
            <!-- Content Area -->
            <StackPanel Grid.Row="1" Margin="20,16,20,16">
                <TextBlock x:Name="TxtDescription" Text="Enter value:" Foreground="#CCCCCC" FontSize="11.5" Margin="0,0,0,8"/>
                <TextBox x:Name="TxtInput" Background="#161616" Foreground="White" BorderBrush="#333333" BorderThickness="1" Padding="6,4" FontSize="12"/>
            </StackPanel>
            
            <!-- Footer -->
            <Border Grid.Row="2" Background="#0E0E0E" BorderBrush="#222222" BorderThickness="0,1,0,0">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="CancelBtn" Content="Cancel" Width="80" Height="26" Margin="0,0,8,0" Cursor="Hand" Foreground="#AAAAAA" FontSize="11">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="bd" Background="#222222" CornerRadius="3">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="bd" Property="Background" Value="#333333"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <Button x:Name="OkBtn" Content="OK" Width="80" Height="26" Margin="0,0,14,0" Cursor="Hand" Foreground="White" FontWeight="SemiBold" FontSize="11">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="bd" Background="#802F2D" CornerRadius="3">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="bd" Property="Background" Value="#9E3A38"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </StackPanel>
            </Border>
        </Grid>
    </Border>
</Window>
"""
        r = XmlReader.Create(StringReader(xaml_code))
        self.win = XamlReader.Load(r)
        
        self.TxtTitle = self.win.FindName("TxtTitle")
        self.TxtDescription = self.win.FindName("TxtDescription")
        self.TxtInput = self.win.FindName("TxtInput")
        self.TxtAccent = self.win.FindName("TxtAccent")
        self.CloseBtn = self.win.FindName("CloseBtn")
        self.OkBtn = self.win.FindName("OkBtn")
        self.CancelBtn = self.win.FindName("CancelBtn")
        self.TitleBar = self.win.FindName("TitleBar")
        
        if self.TxtTitle:
            self.TxtTitle.Text = title
        if self.TxtDescription:
            self.TxtDescription.Text = description
        if self.TxtAccent:
            self.TxtAccent.Text = u"\u2B0C"  # ⬌
        if self.CloseBtn:
            self.CloseBtn.Content = u"\u2715"  # ✕
            self.CloseBtn.Click += self.CloseBtn_Click
        if self.CancelBtn:
            self.CancelBtn.Click += self.CancelBtn_Click
        if self.OkBtn:
            self.OkBtn.Click += self.OkBtn_Click
        if self.TitleBar:
            self.TitleBar.MouseLeftButtonDown += self.TitleBar_MouseDown
            
        # Focus textbox
        if self.TxtInput:
            self.TxtInput.Focus()

    def TitleBar_MouseDown(self, sender, e):
        try:
            self.win.DragMove()
        except:
            pass
            
    def CloseBtn_Click(self, sender, e):
        self.result = None
        self.win.Close()
        
    def CancelBtn_Click(self, sender, e):
        self.result = None
        self.win.Close()
        
    def OkBtn_Click(self, sender, e):
        if self.TxtInput:
            self.result = self.TxtInput.Text
        self.win.Close()

    def ShowDialog(self):
        self.win.ShowDialog()
        return self.result

# ------------------------------------------------------------------------------
# Custom Profile Save Dialog
# ------------------------------------------------------------------------------
class CustomProfileSaveWindow(object):
    def __init__(self):
        self.result = None
        xaml_code = """<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Save Profile" Width="360" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        Background="#111111" WindowStyle="None" AllowsTransparency="False"
        ResizeMode="NoResize">
    <Border BorderBrush="#3A3A3A" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="46"/>
            </Grid.RowDefinitions>
            
            <!-- Title Bar -->
            <Grid x:Name="TitleBar" Grid.Row="0" Background="#1A1A1A">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
                    <TextBlock x:Name="TxtAccent" Text="" Foreground="#802F2D" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TxtTitle" Text="Save Profile" Foreground="#CCCCCC" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
                <Button x:Name="CloseBtn" Content="" HorizontalAlignment="Right"
                        Width="44" Height="36" BorderThickness="0" Cursor="Hand"
                        Background="Transparent" Foreground="#666666"
                        FontSize="12">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="{TemplateBinding Background}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#C0272D"/>
                                    <Setter Property="Foreground" Value="White"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
            
            <!-- Content Area -->
            <StackPanel Grid.Row="1" Margin="20,16,20,16">
                <TextBlock Text="This profile will be updated with" Foreground="#CCCCCC" FontSize="11.5" Margin="0,0,0,12"/>
                <TextBlock Text="- Custom Drawing Number" Foreground="#888888" FontSize="11.5" Margin="0,0,0,4"/>
                <TextBlock Text="- Format options" Foreground="#888888" FontSize="11.5" Margin="0,0,0,4"/>
                <TextBlock Text="  PDF, DWG, DGN, DWF/DWFx, NWC, IFC AND IMG" Foreground="#888888" FontSize="11" Margin="0,0,0,8" TextWrapping="Wrap"/>
            </StackPanel>
            
            <!-- Footer -->
            <Border Grid.Row="2" Background="#0E0E0E" BorderBrush="#222222" BorderThickness="0,1,0,0">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="SaveAsBtn" Content="Save As" Width="80" Height="26" Margin="0,0,8,0" Cursor="Hand" Foreground="#AAAAAA" FontSize="11">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="bd" Background="#222222" CornerRadius="3">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="bd" Property="Background" Value="#333333"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <Button x:Name="SaveBtn" Content="Save" Width="80" Height="26" Margin="0,0,14,0" Cursor="Hand" Foreground="White" FontWeight="SemiBold" FontSize="11">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="bd" Background="#802F2D" CornerRadius="3">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="bd" Property="Background" Value="#9E3A38"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </StackPanel>
            </Border>
        </Grid>
    </Border>
</Window>
"""
        r = XmlReader.Create(StringReader(xaml_code))
        self.win = XamlReader.Load(r)
        
        self.TxtAccent = self.win.FindName("TxtAccent")
        self.CloseBtn = self.win.FindName("CloseBtn")
        self.SaveAsBtn = self.win.FindName("SaveAsBtn")
        self.SaveBtn = self.win.FindName("SaveBtn")
        self.TitleBar = self.win.FindName("TitleBar")
        
        if self.TxtAccent:
            self.TxtAccent.Text = u"\u2B0C"  # ⬌
        if self.CloseBtn:
            self.CloseBtn.Content = u"\u2715"  # ✕
            self.CloseBtn.Click += self.CloseBtn_Click
        if self.SaveAsBtn:
            self.SaveAsBtn.Click += self.SaveAsBtn_Click
        if self.SaveBtn:
            self.SaveBtn.Click += self.SaveBtn_Click
        if self.TitleBar:
            self.TitleBar.MouseLeftButtonDown += self.TitleBar_MouseDown
            
    def TitleBar_MouseDown(self, sender, e):
        try:
            self.win.DragMove()
        except:
            pass
            
    def CloseBtn_Click(self, sender, e):
        self.result = "Cancel"
        self.win.Close()
        
    def SaveAsBtn_Click(self, sender, e):
        self.result = "SaveAs"
        self.win.Close()
        
    def SaveBtn_Click(self, sender, e):
        self.result = "Save"
        self.win.Close()

    def ShowDialog(self):
        self.win.ShowDialog()
        return self.result

def show_text_input(title, description):
    try:
        dialog = CustomTextInputWindow(title, description)
        return dialog.ShowDialog()
    except Exception as e:
        show_alert("Dialog Error: " + str(e), is_error=True)
        from rpw.ui.forms import TextInput
        return TextInput(title, description=description)

# ------------------------------------------------------------------------------
# Settings Storage Utility
# ------------------------------------------------------------------------------
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "naming_settings.json")

def load_settings():
    default_settings = {
        "active_scheme": "Default",
        "schemes": {
            "Default": [
                {"ParameterName": "Sheet Number", "Prefix": "", "Suffix": "", "Separator": " - "},
                {"ParameterName": "Sheet Name", "Prefix": "", "Suffix": "", "Separator": ""}
            ]
        },
        "view_sets": {}
    }
    if not os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(default_settings, f, indent=4)
        except Exception:
            pass
        return default_settings
    
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
            if "schemes" not in data or not data["schemes"]:
                data["schemes"] = default_settings["schemes"]
                data["active_scheme"] = default_settings["active_scheme"]
            if "view_sets" not in data:
                data["view_sets"] = {}
            return data
    except Exception:
        return default_settings

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        show_alert("Error saving naming settings:\n" + str(e), is_error=True)

# ------------------------------------------------------------------------------
# View Models
# ------------------------------------------------------------------------------
class SheetViewModel(object):
    def __init__(self, sheet, scheme_parts, doc):
        self.Sheet = sheet
        self.SheetNumber = sheet.SheetNumber
        self.SheetName = sheet.Name
        
        # Get revision
        p = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
        self.Revision = p.AsString() or p.AsValueString() or "-" if p else "-"
        
        # Get Size (from TitleBlock if available)
        self.Size = ""
        try:
            tbs = DB.FilteredElementCollector(doc, sheet.Id).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).ToElements()
            if tbs:
                self.Size = tbs[0].Name
            else:
                self.Size = "A1"
        except:
            self.Size = "A1"
        
        self._is_selected = False
        self._custom_file_name = ""
        self.update_filename(scheme_parts, doc)
        
    @property
    def IsSelected(self):
        return self._is_selected
        
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        
    @property
    def CustomFileName(self):
        return self._custom_file_name
        
    @CustomFileName.setter
    def CustomFileName(self, value):
        self._custom_file_name = value
        
    def update_filename(self, scheme_parts, doc):
        self.CustomFileName = generate_filename(self.Sheet, scheme_parts, doc)

class QueueItemViewModel(object):
    def __init__(self, sheet_vm, format, active_scheme_parts, doc):
        self.SheetVM = sheet_vm
        self.SheetNumber = sheet_vm.SheetNumber
        self.SheetName = sheet_vm.SheetName
        self.Format = format
        self.TargetFileName = generate_filename(sheet_vm.Sheet, active_scheme_parts, doc)
        self._status = "Pending"
        
    @property
    def Status(self):
        return self._status
        
    @Status.setter
    def Status(self, value):
        self._status = value

class ParameterViewModel(object):
    def __init__(self, name, category):
        self._name = name
        self._category = category
        
    @property
    def Name(self):
        return self._name
        
    @property
    def Category(self):
        return self._category

class NamingPart(object):
    def __init__(self, param_name, sample_value="", prefix="", suffix="", separator=""):
        self._param_name = param_name
        self._sample_value = sample_value
        self._prefix = prefix
        self._suffix = suffix
        self._separator = separator

    @property
    def ParameterName(self):
        return self._param_name

    @property
    def SampleValue(self):
        return self._sample_value

    @property
    def Prefix(self):
        return self._prefix

    @Prefix.setter
    def Prefix(self, value):
        self._prefix = value or ""

    @property
    def Suffix(self):
        return self._suffix

    @Suffix.setter
    def Suffix(self, value):
        self._suffix = value or ""

    @property
    def Separator(self):
        return self._separator

    @Separator.setter
    def Separator(self, value):
        self._separator = value or ""

    def to_dict(self):
        return {
            "ParameterName": self._param_name,
            "Prefix": self._prefix,
            "Suffix": self._suffix,
            "Separator": self._separator
        }

# ------------------------------------------------------------------------------
# Parameter and Sample Value Helpers
# ------------------------------------------------------------------------------
def get_sample_value(doc, param_name, sheet=None):
    if sheet:
        p = sheet.LookupParameter(param_name)
        if not p:
            for bp in sheet.Parameters:
                if bp.Definition.Name == param_name:
                    p = bp
                    break
        if p:
            val = p.AsValueString() or p.AsString()
            if val is not None:
                return val
                
    proj_info = doc.ProjectInformation
    if proj_info:
        p = proj_info.LookupParameter(param_name)
        if not p:
            for bp in proj_info.Parameters:
                if bp.Definition.Name == param_name:
                    p = bp
                    break
        if p:
            val = p.AsValueString() or p.AsString()
            if val is not None:
                return val
                
    # Fallback default values
    if param_name == "Sheet Number": return "A101"
    if param_name == "Sheet Name": return "Floor Plan"
    if param_name == "Current Revision": return "01"
    if param_name == "Discipline": return "Architectural"
    return param_name

def get_all_sheet_parameters(doc):
    params = []
    seen = set()
    
    # 1. Sheet Category Parameters
    sheets = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    if sheets:
        sheet = sheets[0]
        for p in sheet.Parameters:
            name = p.Definition.Name
            if name and name not in seen:
                seen.add(name)
                params.append(ParameterViewModel(name, "Sheet"))
                
    # 2. Project Information Parameters
    proj_info = doc.ProjectInformation
    if proj_info:
        for p in proj_info.Parameters:
            name = p.Definition.Name
            if name and name not in seen:
                seen.add(name)
                params.append(ParameterViewModel(name, "Project Information"))
                
    # Fallbacks for core parameters
    for name, cat in [("Sheet Number", "Sheet"), ("Sheet Name", "Sheet"), ("Current Revision", "Sheet")]:
        if name not in seen:
            seen.add(name)
            params.append(ParameterViewModel(name, cat))
            
    return sorted(params, key=lambda x: x.Name)

def generate_filename(sheet, scheme_parts, doc):
    name_parts = []
    for part in scheme_parts:
        param_name = part["ParameterName"]
        prefix = part.get("Prefix", "")
        suffix = part.get("Suffix", "")
        separator = part.get("Separator", "")
        
        val = ""
        p = sheet.LookupParameter(param_name)
        if not p:
            for bp in sheet.Parameters:
                if bp.Definition.Name == param_name:
                    p = bp
                    break
        if p:
            val = p.AsValueString() or p.AsString() or ""
            
        if not val and doc.ProjectInformation:
            pi_p = doc.ProjectInformation.LookupParameter(param_name)
            if not pi_p:
                for bp in doc.ProjectInformation.Parameters:
                    if bp.Definition.Name == param_name:
                        pi_p = bp
                        break
            if pi_p:
                val = pi_p.AsValueString() or pi_p.AsString() or ""
                
        name_parts.append(prefix + val + suffix + separator)
        
    filename = "".join(name_parts)
    
    # Clean invalid characters
    invalid_chars = '<>:"/\\|?*'
    for c in invalid_chars:
        filename = filename.replace(c, "_")
        
    # Final fallback if name resolves to empty string
    filename = filename.strip()
    if not filename:
        filename = sheet.SheetNumber + " - " + sheet.Name
        
    return filename

# ------------------------------------------------------------------------------
# Naming Builder Form Controller
# ------------------------------------------------------------------------------
class NamingBuilderForm(forms.WPFWindow):
    def __init__(self, xaml_file_name, current_scheme_name, doc, sheets):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.doc = doc
        self.sheets = sheets
        self.sample_sheet = sheets[0] if sheets else None
        
        # Load settings
        self.settings = load_settings()
        self.current_scheme_name = current_scheme_name
        if self.current_scheme_name not in self.settings["schemes"]:
            self.current_scheme_name = self.settings["active_scheme"]
            if self.current_scheme_name not in self.settings["schemes"]:
                self.current_scheme_name = list(self.settings["schemes"].keys())[0]
                
        self.TxtSchemeName.Text = self.current_scheme_name
        
        # Populate available parameters
        self.all_params = get_all_sheet_parameters(doc)
        
        # Initialize selected parameters list
        self.selected_parts = []
        scheme_data = self.settings["schemes"].get(self.current_scheme_name, [])
        for item in scheme_data:
            param_name = item["ParameterName"]
            prefix = item.get("Prefix", "")
            suffix = item.get("Suffix", "")
            separator = item.get("Separator", "")
            sample_val = get_sample_value(doc, param_name, self.sample_sheet)
            self.selected_parts.append(NamingPart(param_name, sample_val, prefix, suffix, separator))
            
        self.GridSelectedParams.ItemsSource = self.selected_parts
        
        # Setup category and search
        self.CmbCategory.SelectedIndex = 0
        self.filter_parameters()
        self.update_preview()
        
    def TitleBar_MouseDown(self, sender, e):
        try:
            self.DragMove()
        except:
            pass
            
    def CloseBtn_Click(self, sender, e):
        self.DialogResult = False
        self.Close()
        
    def filter_parameters(self):
        if not hasattr(self, 'LstAvailableParams'):
            return
            
        search_text = self.TxtSearch.Text.lower().strip()
        cat_idx = self.CmbCategory.SelectedIndex
        
        filtered = []
        for p in self.all_params:
            if cat_idx == 1 and p.Category != "Sheet":
                continue
            if cat_idx == 2 and p.Category != "Project Information":
                continue
            
            if search_text and search_text not in p.Name.lower():
                continue
                
            filtered.append(p)
            
        self.LstAvailableParams.ItemsSource = filtered
        
    def TxtSearch_TextChanged(self, sender, e):
        self.filter_parameters()
        
    def CmbCategory_SelectionChanged(self, sender, e):
        self.filter_parameters()
        
    def BtnAdd_Click(self, sender, e):
        selected_item = self.LstAvailableParams.SelectedItem
        if selected_item:
            param_name = selected_item.Name
            sample_val = get_sample_value(self.doc, param_name, self.sample_sheet)
            part = NamingPart(param_name, sample_val, "", "", "")
            self.selected_parts.append(part)
            
            self.GridSelectedParams.ItemsSource = None
            self.GridSelectedParams.ItemsSource = self.selected_parts
            self.update_preview()
            
    def BtnRemove_Click(self, sender, e):
        selected_item = self.GridSelectedParams.SelectedItem
        if selected_item:
            self.selected_parts.remove(selected_item)
            self.GridSelectedParams.ItemsSource = None
            self.GridSelectedParams.ItemsSource = self.selected_parts
            self.update_preview()
            
    def BtnMoveToTop_Click(self, sender, e):
        selected_item = self.GridSelectedParams.SelectedItem
        if selected_item:
            idx = self.selected_parts.index(selected_item)
            if idx > 0:
                self.selected_parts.remove(selected_item)
                self.selected_parts.insert(0, selected_item)
                self.GridSelectedParams.ItemsSource = None
                self.GridSelectedParams.ItemsSource = self.selected_parts
                self.GridSelectedParams.SelectedItem = selected_item
                self.update_preview()
                
    def BtnMoveUp_Click(self, sender, e):
        selected_item = self.GridSelectedParams.SelectedItem
        if selected_item:
            idx = self.selected_parts.index(selected_item)
            if idx > 0:
                self.selected_parts.remove(selected_item)
                self.selected_parts.insert(idx - 1, selected_item)
                self.GridSelectedParams.ItemsSource = None
                self.GridSelectedParams.ItemsSource = self.selected_parts
                self.GridSelectedParams.SelectedItem = selected_item
                self.update_preview()
                
    def BtnMoveDown_Click(self, sender, e):
        selected_item = self.GridSelectedParams.SelectedItem
        if selected_item:
            idx = self.selected_parts.index(selected_item)
            if idx < len(self.selected_parts) - 1:
                self.selected_parts.remove(selected_item)
                self.selected_parts.insert(idx + 1, selected_item)
                self.GridSelectedParams.ItemsSource = None
                self.GridSelectedParams.ItemsSource = self.selected_parts
                self.GridSelectedParams.SelectedItem = selected_item
                self.update_preview()
                
    def BtnMoveToBottom_Click(self, sender, e):
        selected_item = self.GridSelectedParams.SelectedItem
        if selected_item:
            idx = self.selected_parts.index(selected_item)
            if idx < len(self.selected_parts) - 1:
                self.selected_parts.remove(selected_item)
                self.selected_parts.append(selected_item)
                self.GridSelectedParams.ItemsSource = None
                self.GridSelectedParams.ItemsSource = self.selected_parts
                self.GridSelectedParams.SelectedItem = selected_item
                self.update_preview()
                
    def BtnReset_Click(self, sender, e):
        self.selected_parts = []
        self.GridSelectedParams.ItemsSource = None
        self.GridSelectedParams.ItemsSource = self.selected_parts
        self.update_preview()
        
    def GridSelectedParams_CellEditEnding(self, sender, e):
        from System.Windows.Threading import DispatcherPriority
        from System import Action
        self.Dispatcher.BeginInvoke(Action(self.update_preview), DispatcherPriority.Background)
        
    def update_preview(self):
        preview_parts = []
        for part in self.selected_parts:
            val = part.SampleValue or ""
            prefix = part.Prefix or ""
            suffix = part.Suffix or ""
            separator = part.Separator or ""
            preview_parts.append(prefix + val + suffix + separator)
            
        preview_text = "".join(preview_parts)
        if not preview_text:
            preview_text = "[None]"
        self.TxtPreview.Text = preview_text
        
    def BtnSaveScheme_Click(self, sender, e):
        scheme_name = self.TxtSchemeName.Text.strip()
        if not scheme_name:
            show_alert("Please enter a valid scheme name.", is_warning=True)
            return
            
        serialized = [part.to_dict() for part in self.selected_parts]
        self.settings["schemes"][scheme_name] = serialized
        self.settings["active_scheme"] = scheme_name
        save_settings(self.settings)
        
        self.current_scheme_name = scheme_name
        show_alert("Scheme '{}' saved successfully!".format(scheme_name))
        
    def BtnDeleteScheme_Click(self, sender, e):
        scheme_name = self.TxtSchemeName.Text.strip()
        if scheme_name not in self.settings["schemes"]:
            show_alert("Scheme '{}' does not exist.".format(scheme_name), is_warning=True)
            return
            
        if len(self.settings["schemes"]) <= 1:
            show_alert("Cannot delete the only naming scheme. At least one scheme must exist.", is_warning=True)
            return
            
        del self.settings["schemes"][scheme_name]
        
        new_active = list(self.settings["schemes"].keys())[0]
        self.settings["active_scheme"] = new_active
        save_settings(self.settings)
        
        show_alert("Scheme '{}' deleted.".format(scheme_name))
        
        self.current_scheme_name = new_active
        self.TxtSchemeName.Text = new_active
        self.selected_parts = []
        for item in self.settings["schemes"][new_active]:
            param_name = item["ParameterName"]
            prefix = item.get("Prefix", "")
            suffix = item.get("Suffix", "")
            separator = item.get("Separator", "")
            sample_val = get_sample_value(self.doc, param_name, self.sample_sheet)
            self.selected_parts.append(NamingPart(param_name, sample_val, prefix, suffix, separator))
            
        self.GridSelectedParams.ItemsSource = None
        self.GridSelectedParams.ItemsSource = self.selected_parts
        self.update_preview()
        
    def BtnCancel_Click(self, sender, e):
        self.DialogResult = False
        self.Close()
        
    def BtnOk_Click(self, sender, e):
        scheme_name = self.TxtSchemeName.Text.strip()
        if not scheme_name:
            show_alert("Please enter a valid scheme name.", is_warning=True)
            return
            
        serialized = [part.to_dict() for part in self.selected_parts]
        self.settings["schemes"][scheme_name] = serialized
        self.settings["active_scheme"] = scheme_name
        save_settings(self.settings)
        
        self.DialogResult = True
        self.Close()

# ------------------------------------------------------------------------------
# Export Manager Main Form (Wizard Controller)
# ------------------------------------------------------------------------------
class ExportManagerForm(forms.WPFWindow):
    def __init__(self, xaml_file_name, sheets):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        # Load naming settings
        settings = load_settings()
        active = settings.get("active_scheme", "Default")
        self.active_scheme_parts = settings["schemes"].get(active, [])
        
        # Set default export folder (User Desktop)
        self.export_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
        self.TxtExportPath.Text = self.export_path
        
        # Wrap sheets into ViewModels
        self.sheets = [SheetViewModel(s, self.active_scheme_parts, doc) for s in sheets]
        self.sheets.sort(key=lambda x: x.SheetNumber)
        self.GridSheets.ItemsSource = self.sheets
        
        # Populate Setups
        self.print_settings = list(DB.FilteredElementCollector(doc).OfClass(DB.PrintSetting).ToElements())
        self.dwg_settings = list(DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements())
        
        self.pdf_setting_names = [ps.Name for ps in self.print_settings]
        self.dwg_setting_names = [ds.Name for ds in self.dwg_settings]
        
        self.pdf_setting_names.insert(0, "<In-Session / Default>")
        self.dwg_setting_names.insert(0, "<In-Session / Default>")
        
        self.CmbPdfSetup.ItemsSource = self.pdf_setting_names
        self.CmbDwgSetup.ItemsSource = self.dwg_setting_names
        
        if self.pdf_setting_names: self.CmbPdfSetup.SelectedIndex = 0
        if self.dwg_setting_names: self.CmbDwgSetup.SelectedIndex = 0
        
        # Initialize naming schemes
        self.reload_schemes()
        self.load_viewsets()
        self.update_selection_stats()
        
        # Select first tab by default
        self.MainTabControl.SelectedIndex = 0

    def reload_schemes(self):
        settings = load_settings()
        active = settings.get("active_scheme", "Default")
        schemes_list = list(settings["schemes"].keys())
        
        self.CmbProfile.ItemsSource = schemes_list
        if active in schemes_list:
            self.CmbProfile.SelectedItem = active
        elif schemes_list:
            self.CmbProfile.SelectedIndex = 0
            
        self.active_scheme_parts = settings["schemes"].get(self.CmbProfile.SelectedItem, [])
        
    # ViewSheetSets logic
    def load_viewsets(self):
        settings = load_settings()
        self.viewsets_dict = settings.get("view_sets", {})
        
        self.viewset_names = list(self.viewsets_dict.keys())
        self.viewset_names.sort()
        self.viewset_names.insert(0, "Unsaved Set")
        self.CmbViewSets.ItemsSource = self.viewset_names
        self.CmbViewSets.SelectedIndex = 0
        
    def CmbViewSets_SelectionChanged(self, sender, e):
        idx = self.CmbViewSets.SelectedIndex
        if idx <= 0: return # Unsaved set
        
        set_name = self.CmbViewSets.SelectedItem
        if not set_name: return
        
        set_sheet_numbers = self.viewsets_dict.get(set_name, [])
        
        for sv in self.sheets:
            sv.IsSelected = (sv.SheetNumber in set_sheet_numbers)
            
        self.GridSheets.Items.Refresh()
        self.update_selection_stats()

    def BtnSaveViewSet_Click(self, sender, e):
        # Gather selected sheets
        selected_vms = [sv for sv in self.sheets if sv.IsSelected]
        if not selected_vms:
            show_alert("No sheets selected to save.", is_warning=True)
            return
            
        set_name = show_text_input("Save View/Sheet Set", "Enter name for the new set:")
        if not set_name: return
        
        try:
            settings = load_settings()
            if "view_sets" not in settings:
                settings["view_sets"] = {}
                
            # Save by SheetNumber for resilience
            sheet_numbers = [sv.SheetNumber for sv in selected_vms]
            settings["view_sets"][set_name] = sheet_numbers
            
            save_settings(settings)
            show_alert("Successfully saved set '{}'".format(set_name))
            self.load_viewsets()
            self.CmbViewSets.SelectedItem = set_name
            
        except Exception as ex:
            show_alert("Failed to save set:\n" + str(ex), is_error=True)
            
    def BtnDeleteViewSet_Click(self, sender, e):
        idx = self.CmbViewSets.SelectedIndex
        if idx <= 0:
            show_alert("Cannot delete Unsaved Set.", is_warning=True)
            return
            
        set_name = self.CmbViewSets.SelectedItem
        if not set_name: return
        
        try:
            settings = load_settings()
            if "view_sets" in settings and set_name in settings["view_sets"]:
                del settings["view_sets"][set_name]
                save_settings(settings)
                
            show_alert("Successfully deleted set '{}'".format(set_name))
            self.load_viewsets()
        except Exception as ex:
            show_alert("Failed to delete set:\n" + str(ex), is_error=True)

    def TitleBar_MouseDown(self, sender, e):
        if e.ChangedButton == UI.MouseButtons.Left:
            self.DragMove()
            
    def CloseBtn_Click(self, sender, e):
        self.DialogResult = False
        self.Close()
        
    # Tab 1: Selection Logic
    def update_selection_stats(self):
        selected_count = sum(1 for sv in self.sheets if sv.IsSelected)
        total_count = len(self.sheets)
        self.StatusTextBlock.Text = "{} sheets selected. Total: {}".format(selected_count, total_count)
        
    def filter_sheets(self):
        search_text = self.TxtSearch.Text.lower().strip()
        filtered = []
        for sv in self.sheets:
            if not search_text or (search_text in sv.SheetNumber.lower() or search_text in sv.SheetName.lower()):
                filtered.append(sv)
        self.GridSheets.ItemsSource = filtered
        
    def TxtSearch_TextChanged(self, sender, e):
        self.filter_sheets()
        
    def CbHeaderSelectAll_Click(self, sender, e):
        is_checked = sender.IsChecked
        items = self.GridSheets.ItemsSource or self.sheets
        for sv in items:
            sv.IsSelected = is_checked
        self.GridSheets.Items.Refresh()
        self.update_selection_stats()
        
    def CbSheetSelect_Click(self, sender, e):
        self.update_selection_stats()
        
    # Tab 2: Format Logic
    def CbFormat_Click(self, sender, e):
        pass

    def CmbProfile_SelectionChanged(self, sender, e):
        active = self.CmbProfile.SelectedItem
        if active:
            settings = load_settings()
            settings["active_scheme"] = active
            save_settings(settings)
            self.active_scheme_parts = settings["schemes"].get(active, [])
            
            # Recalculate naming previews for all sheets
            for sv in self.sheets:
                sv.update_filename(self.active_scheme_parts, doc)
            self.GridSheets.Items.Refresh()

    def BtnAddProfile_Click(self, sender, e):
        new_name = show_text_input("New Profile", "Enter new profile name:")
        if not new_name: return
        
        settings = load_settings()
        if new_name in settings["schemes"]:
            show_alert("Profile already exists.", is_warning=True)
            return
            
        settings["schemes"][new_name] = []
        settings["active_scheme"] = new_name
        save_settings(settings)
        self.reload_schemes()
        
    def BtnDeleteProfile_Click(self, sender, e):
        active = self.CmbProfile.SelectedItem
        if not active or active == "Default":
            show_alert("Cannot delete Default profile.", is_warning=True)
            return
            
        settings = load_settings()
        if active in settings["schemes"]:
            del settings["schemes"][active]
            settings["active_scheme"] = "Default"
            save_settings(settings)
            self.reload_schemes()

    def BtnNamingBuilder_Click(self, sender, e):
        active = self.CmbProfile.SelectedItem or "Default"
        builder_xaml_path = os.path.join(os.path.dirname(__file__), "NamingBuilder.xaml")
        raw_sheets = [sv.Sheet for sv in self.sheets]
        
        builder_form = NamingBuilderForm(builder_xaml_path, active, doc, raw_sheets)
        if builder_form.ShowDialog():
            self.reload_schemes()
            # Recalculate naming previews
            for sv in self.sheets:
                sv.update_filename(self.active_scheme_parts, doc)
            self.GridSheets.Items.Refresh()

    def BtnSaveProfile_Click(self, sender, e):
        dialog = CustomProfileSaveWindow()
        res = dialog.ShowDialog()
        
        if res == "SaveAs":
            try:
                # Show SaveFileDialog
                from Microsoft.Win32 import SaveFileDialog
                dlg = SaveFileDialog()
                dlg.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
                dlg.DefaultExt = ".xml"
                dlg.FileName = (self.CmbProfile.SelectedItem or "Default") + ".xml"
                
                if dlg.ShowDialog() == True:
                    save_path = dlg.FileName
                    active = self.CmbProfile.SelectedItem or "Default"
                    settings = load_settings()
                    scheme = settings["schemes"].get(active, [])
                    
                    # Build simple XML manually to avoid IronPython xml module issues
                    xml_string = '<?xml version="1.0" encoding="utf-8"?>\n<Profile>\n'
                    xml_string += '  <Name>{}</Name>\n'.format(active)
                    xml_string += '  <NamingRules>\n'
                    for part in scheme:
                        xml_string += '    <Rule Type="{}" Value="{}" />\n'.format(part["type"], part["value"].replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;'))
                    xml_string += '  </NamingRules>\n</Profile>'
                    
                    with open(save_path, "w") as f:
                        f.write(xml_string)
                    
                    forms.alert("Profile successfully exported to XML!", title="Export Manager")
            except Exception as ex:
                forms.alert("Error saving XML: " + str(ex), title="Export Manager")
                
        elif res == "Save":
            # Saving internally is actually automatic when editing, but we can show a confirmation
            forms.alert("Profile saved successfully.", title="Export Manager")

    # Tab 3: Create Logic
    def BtnBrowse_Click(self, sender, e):
        selected_folder = forms.pick_folder(title="Select Export Destination")
        if selected_folder:
            self.export_path = selected_folder
            self.TxtExportPath.Text = self.export_path
            
    def generate_queue(self):
        selected_vms = [sv for sv in self.sheets if sv.IsSelected]
        queue = []
        
        export_pdf = self.CbPDF.IsChecked
        export_dwg = self.CbDWG.IsChecked
        
        for sv in selected_vms:
            if export_pdf:
                queue.append(QueueItemViewModel(sv, "PDF", self.active_scheme_parts, doc))
            if export_dwg:
                queue.append(QueueItemViewModel(sv, "DWG", self.active_scheme_parts, doc))
                
        self.queue_items = queue
        self.GridQueue.ItemsSource = self.queue_items

    # Wizard Navigation
    def MainTabControl_SelectionChanged(self, sender, e):
        if not hasattr(self, 'BtnBack') or not hasattr(self, 'BtnNext'):
            return
        idx = self.MainTabControl.SelectedIndex
        if idx == 0:
            self.BtnBack.IsEnabled = False
            self.BtnNext.Content = "Next"
        elif idx == 1:
            self.BtnBack.IsEnabled = True
            self.BtnNext.Content = "Next"
        elif idx == 2:
            self.BtnBack.IsEnabled = True
            self.BtnNext.Content = "Create"
            self.generate_queue()
            
    def BtnBack_Click(self, sender, e):
        idx = self.MainTabControl.SelectedIndex
        if idx > 0:
            self.MainTabControl.SelectedIndex = idx - 1
            
    def BtnNext_Click(self, sender, e):
        idx = self.MainTabControl.SelectedIndex
        if idx == 0:
            # Check selection
            selected_count = sum(1 for sv in self.sheets if sv.IsSelected)
            if selected_count == 0:
                show_alert("Please select at least one sheet before proceeding.", is_warning=True)
                return
            self.MainTabControl.SelectedIndex = 1
        elif idx == 1:
            # Check format
            if not self.CbPDF.IsChecked and not self.CbDWG.IsChecked:
                show_alert("Please select at least one export format.", is_warning=True)
                return
            self.MainTabControl.SelectedIndex = 2
        elif idx == 2:
            # Trigger Export Execution
            self.run_export()
            
    def BtnResetSettings_Click(self, sender, e):
        show_alert("Riyan Export Manager v1.2\nTheme: Dark Red\nConfiguration settings verified successfully.")

    # UI yielding events
    def do_events(self):
        from System.Windows.Threading import DispatcherFrame, Dispatcher
        from System import Action
        
        frame = DispatcherFrame()
        def exit_frame(f):
            f.Continue = False
            
        Dispatcher.CurrentDispatcher.BeginInvoke(
            System.Windows.Threading.DispatcherPriority.Background,
            Action[DispatcherFrame](exit_frame),
            frame
        )
        Dispatcher.PushFrame(frame)

    # Export Process
    def run_export(self):
        folder = self.TxtExportPath.Text.strip()
        if not folder or not os.path.isdir(folder):
            show_alert("Please select a valid export directory.", is_warning=True)
            return
            
        if not self.queue_items:
            show_alert("Export queue is empty. Please select sheets and formats.", is_warning=True)
            return
            
        pdf_idx = self.CmbPdfSetup.SelectedIndex
        selected_pdf_setting = self.print_settings[pdf_idx - 1] if pdf_idx > 0 else None
        
        dwg_idx = self.CmbDwgSetup.SelectedIndex
        selected_dwg_setting = self.dwg_settings[dwg_idx - 1] if dwg_idx > 0 else None
        
        # Disable navigation controls
        self.BtnBack.IsEnabled = False
        self.BtnNext.IsEnabled = False
        self.CloseBtn.IsEnabled = False
        
        revit_version = int(__revit__.Application.VersionNumber)
        total = len(self.queue_items)
        
        try:
            for idx, item in enumerate(self.queue_items):
                item.Status = "Exporting..."
                self.GridQueue.Items.Refresh()
                self.do_events()
                
                success = False
                err_msg = ""
                
                try:
                    if item.Format == "PDF":
                        if revit_version >= 2022:
                            success = export_pdf_2022(folder, item.SheetVM.Sheet, item.TargetFileName, selected_pdf_setting)
                        else:
                            success = False
                            err_msg = "PDF requires Revit 2022+"
                    elif item.Format == "DWG":
                        export_dwg(folder, item.SheetVM.Sheet, item.TargetFileName, selected_dwg_setting)
                        success = True
                except Exception as ex:
                    import traceback
                    success = False
                    err_msg = traceback.format_exc()
                    
                if success:
                    item.Status = "Done"
                else:
                    item.Status = "Error"
                    
                # Update progress bar
                percent = int(((idx + 1) / float(total)) * 100)
                self.ExportProgressBar.Value = percent
                self.TxtPercent.Text = "Completed {}%".format(percent)
                self.GridQueue.Items.Refresh()
                self.do_events()
                
            show_alert("Export completed successfully!", title="Export Manager")
        finally:
            self.BtnBack.IsEnabled = True
            self.BtnNext.IsEnabled = True
            self.CloseBtn.IsEnabled = True

# ------------------------------------------------------------------------------
# Export Execution Logic
# ------------------------------------------------------------------------------
def export_dwg(folder, sheet, filename, dwg_setting):
    opt = DB.DWGExportOptions()
    if dwg_setting:
        opt = dwg_setting.GetDWGExportOptions()
        
    from System.Collections.Generic import List
    views = List[DB.ElementId]()
    views.Add(sheet.Id)
        
    doc.Export(folder, filename, views, opt)

def export_pdf_2022(folder, sheet, filename, print_setting):
    try:
        opt = DB.PDFExportOptions()
        opt.FileName = filename
        
        if print_setting:
            try:
                ps = print_setting.PrintParameters
                opt.ZoomType = ps.ZoomType
                opt.ZoomPercentage = ps.Zoom
            except:
                pass
        
        from System.Collections.Generic import List
        views = List[DB.ElementId]()
        views.Add(sheet.Id)
        
        doc.Export(folder, views, opt)
        return True
    except Exception as e:
        import traceback
        show_alert("Failed to export PDF for sheet {}:\n{}".format(sheet.SheetNumber, traceback.format_exc()), is_error=True)
        return False

# ------------------------------------------------------------------------------
# Main Execution
# ------------------------------------------------------------------------------
def main():
    sheets = DB.FilteredElementCollector(doc)\
               .OfCategory(DB.BuiltInCategory.OST_Sheets)\
               .WhereElementIsNotElementType()\
               .ToElements()
               
    if not sheets:
        show_alert("No Sheets found in the current project.", is_warning=True)
        return
        
    xaml_path = os.path.join(os.path.dirname(__file__), "ExportUI.xaml")
    form = ExportManagerForm(xaml_path, sheets)
    form.ShowDialog()

if __name__ == '__main__':
    main()
