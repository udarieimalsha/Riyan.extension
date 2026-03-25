# -*- coding: utf-8 -*-
"""
Get Coordinates
Exact Screenshot UI Version
Author: Chalana Perera
"""

import clr
import os
# --- ASSEMBLY REFERENCES (MINIMIZED) ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
# --- END REFERENCES ---

from Autodesk.Revit.DB import (
    FilteredElementCollector, 
    BuiltInCategory, 
    BasePoint,
)
import System.Windows.Controls as Controls
import System.Windows.Media as Media
import System.Windows as WPF
import System.Drawing as Drawing
import System.Drawing.Imaging as Imaging
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
from pyrevit import revit, DB, forms

# Constants
doc = revit.doc
SCRIPT_DIR = os.path.dirname(__file__)
MAROON = "#802F2D"
DARK_BG = "#000000"
PARAMETER_NAMES = ["Coord_X", "Coord_Y", "Coord_Z"]

XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Get Coordinates - v3.2 (PREMIUM MASTER STABLE)"
    Width="450" Height="780"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    FontFamily="Segoe UI"
    Background="#000000"
    WindowStyle="None"
    AllowsTransparency="True">

    <Window.Resources>
        <Style TargetType="TextBlock" x:Key="SectionHeader">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#802F2D"/>
        </Style>
        <Style TargetType="TextBlock" x:Key="SubLabel">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Foreground" Value="#802F2D"/>
        </Style>
        <Style TargetType="Button" x:Key="TinyBtn">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#802F2D"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border BorderBrush="#802F2D" BorderThickness="1" CornerRadius="3" Padding="4,1">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBox" x:Key="BlackCombo">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" BorderBrush="#802F2D" BorderThickness="1" Background="#000000"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Name="Border" Background="#000000" BorderBrush="#802F2D" BorderThickness="1" CornerRadius="5">
                                            <Grid>
                                                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="30"/></Grid.ColumnDefinitions>
                                                <Path Grid.Column="1" Data="M 0 0 L 4 4 L 8 0 Z" Fill="White" HorizontalAlignment="Center" VerticalAlignment="Center" Stretch="Uniform" Width="8"/>
                                            </Grid>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,0,30,0" VerticalAlignment="Center" HorizontalAlignment="Left" />
                            <Popup x:Name="PART_Popup" AllowsTransparency="true" Placement="Bottom" IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}" Focusable="false">
                                <Border x:Name="DropDownBorder" Background="#000000" BorderBrush="#802F2D" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}">
                                    <ScrollViewer CanContentScroll="true">
                                        <ItemsPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border BorderBrush="#333333" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>  <!-- Custom Title Bar -->
                <RowDefinition Height="120"/> <!-- Header -->
                <RowDefinition Height="*"/>   <!-- Content -->
                <RowDefinition Height="80"/>  <!-- Footer -->
            </Grid.RowDefinitions>

            <!-- Title Bar -->
            <Border x:Name="HeaderBar" Grid.Row="0" Background="#1A1A1A">
                <Grid Margin="10,0">
                    <TextBlock Text="Update Coordinates Configuration" Foreground="#AAAAAA" VerticalAlignment="Center" FontSize="11"/>
                    <Button x:Name="Btn_Close" Content="×" HorizontalAlignment="Right" Background="Transparent" Foreground="White" BorderThickness="0" FontSize="18" VerticalAlignment="Center"/>
                </Grid>
            </Border>

            <!-- Header -->
            <Border Grid.Row="1" Padding="25,10">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" VerticalAlignment="Center">
                        <TextBlock Text="Update Coordinates" FontSize="24" FontWeight="Bold" Foreground="White"/>
                        <TextBlock Text="Select foundations and columns to calculate coordinates." FontSize="11" Foreground="#888888" Margin="0,5,0,0"/>
                    </StackPanel>
                    <Image x:Name="UI_Logo" Grid.Column="1" Width="85" Height="85" Stretch="Uniform" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </Grid>
            </Border>

            <Separator Grid.Row="1" VerticalAlignment="Bottom" Background="#802F2D" Height="1" Margin="10,0"/>

            <!-- Main Content -->
            <StackPanel Grid.Row="2" Margin="25,15">
                <!-- Foundations -->
                <Grid Margin="0,0,0,5">
                    <TextBlock Text="STRUCTURAL FOUNDATIONS" Style="{StaticResource SectionHeader}"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="Btn_SAll_F" Content="Select All" Style="{StaticResource TinyBtn}" Margin="0,0,5,0" Tag="F"/>
                        <Button x:Name="Btn_Clr_F" Content="Clear" Style="{StaticResource TinyBtn}" Tag="F"/>
                    </StackPanel>
                </Grid>
                <Border Height="180" BorderBrush="#333333" BorderThickness="1" Margin="0,0,0,20">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="Panel_Foundations" Margin="10"/>
                    </ScrollViewer>
                </Border>

                <!-- Columns -->
                <Grid Margin="0,0,0,5">
                    <TextBlock Text="STRUCTURAL COLUMNS" Style="{StaticResource SectionHeader}"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="Btn_SAll_C" Content="Select All" Style="{StaticResource TinyBtn}" Margin="0,0,5,0" Tag="C"/>
                        <Button x:Name="Btn_Clr_C" Content="Clear" Style="{StaticResource TinyBtn}" Tag="C"/>
                    </StackPanel>
                </Grid>
                <Border Height="180" BorderBrush="#333333" BorderThickness="1" Margin="0,0,0,20">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="Panel_Columns" Margin="10"/>
                    </ScrollViewer>
                </Border>

                <!-- Settings -->
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <StackPanel Grid.Column="0">
                        <TextBlock Text="Coordinate System" Style="{StaticResource SubLabel}" Margin="0,0,0,5"/>
                        <ComboBox x:Name="Combo_System" Style="{StaticResource BlackCombo}">
                            <ComboBoxItem Content="Survey Point" IsSelected="True"/>
                            <ComboBoxItem Content="Project Base Point"/>
                        </ComboBox>
                    </StackPanel>

                    <StackPanel Grid.Column="2">
                        <TextBlock Text="Output Units" Style="{StaticResource SubLabel}" Margin="0,0,0,5"/>
                        <ComboBox x:Name="Combo_Units" Style="{StaticResource BlackCombo}">
                            <ComboBoxItem Content="Millimeters (mm)"/>
                            <ComboBoxItem Content="Meters (m)" IsSelected="True"/>
                        </ComboBox>
                    </StackPanel>
                </Grid>
            </StackPanel>

            <!-- Footer -->
            <Border Grid.Row="3" Background="#0A0A0A" Padding="25,0">
                <Grid VerticalAlignment="Center">
                    <Button x:Name="Btn_Cancel" Content="Cancel" Width="100" Height="28" HorizontalAlignment="Left" Background="#2A2A2A" Foreground="White" BorderThickness="0">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border Background="{TemplateBinding Background}" CornerRadius="5">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <Button x:Name="Btn_Run" Content="Update Coordinates" Width="180" Height="28" HorizontalAlignment="Right" Background="#802F2D" Foreground="White" BorderThickness="0">
                         <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border Background="{TemplateBinding Background}" CornerRadius="5">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
"""
# Master Logo Path (Linking to existing toolkit asset)
MASTER_LOGO = os.path.abspath(os.path.join(SCRIPT_DIR, r"..\Copy from Link.pushbutton\logo.png"))
if not os.path.exists(MASTER_LOGO):
    MASTER_LOGO = os.path.join(SCRIPT_DIR, "logo.png")

def load_logo(image_control, path):
    if not image_control or not os.path.exists(path): return
    try:
        from System.IO import FileStream, FileMode, FileAccess
        from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
        bi = BitmapImage()
        with FileStream(path, FileMode.Open, FileAccess.Read) as fs:
            bi.BeginInit()
            bi.StreamSource = fs
            bi.CacheOption = BitmapCacheOption.OnLoad
            bi.EndInit()
        image_control.Source = bi
    except: pass

class CoordinateToolWindow(WPF.Window):
    def __init__(self, foundation_types, column_types):
        from System.Windows.Markup import XamlReader
        self.ui = XamlReader.Parse(XAML)
        
        # Elements
        self.header_bar = self.ui.FindName("HeaderBar")
        self.panel_f = self.ui.FindName("Panel_Foundations")
        self.panel_c = self.ui.FindName("Panel_Columns")
        self.combo_sys = self.ui.FindName("Combo_System")
        self.combo_unit = self.ui.FindName("Combo_Units")
        
        # Events
        self.header_bar.MouseLeftButtonDown += (lambda s, e: self.ui.DragMove())
        self.ui.FindName("Btn_Close").Click += self.OnCancel
        self.ui.FindName("Btn_SAll_F").Click += self.OnSelectAll
        self.ui.FindName("Btn_Clr_F").Click += self.OnClear
        self.ui.FindName("Btn_SAll_C").Click += self.OnSelectAll
        self.ui.FindName("Btn_Clr_C").Click += self.OnClear
        self.ui.FindName("Btn_Cancel").Click += self.OnCancel
        self.ui.FindName("Btn_Run").Click += self.OnRun
        
        # Load Official Logo
        load_logo(self.ui.FindName("UI_Logo"), MASTER_LOGO)

        # Populate
        self.cb_f = self.populate_list(self.panel_f, foundation_types)
        self.cb_c = self.populate_list(self.panel_c, column_types)
        self.result = None

    def populate_list(self, panel, types):
        checkboxes = []
        for t in types:
            try:
                fam = t.Family.Name if hasattr(t, "Family") else "Type"
                name = "{} - {}".format(fam, revit.query.get_name(t))
            except: name = revit.query.get_name(t) or "Unknown"
            cb = Controls.CheckBox(Content=name, Foreground=Media.Brushes.White, IsChecked=True, Margin=WPF.Thickness(0,2,0,2), Tag=t)
            panel.Children.Add(cb); checkboxes.append(cb)
        return checkboxes

    def OnSelectAll(self, sender, args):
        for cb in (self.cb_f if "Btn_SAll_F" in sender.Name else self.cb_c): cb.IsChecked = True
    def OnClear(self, sender, args):
        for cb in (self.cb_f if "Btn_Clr_F" in sender.Name else self.cb_c): cb.IsChecked = False
    def OnCancel(self, sender, args): self.ui.Close()
    def OnRun(self, sender, args):
        ids = [cb.Tag.Id for cb in self.cb_f + self.cb_c if cb.IsChecked]
        if not ids:
            forms.alert("Select at least one type.")
            return
        self.result = {"ids": ids, "mode": self.combo_sys.Text, "unit": self.combo_unit.Text}
        self.ui.Close()
    def show(self): self.ui.ShowDialog(); return self.result

class AlertWindow:
    def __init__(self, message):
        from System.Windows.Markup import XamlReader
        xaml_str = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Success" Width="450" Height="250" WindowStartupLocation="CenterScreen"
            AllowsTransparency="True" WindowStyle="None" Background="Transparent">
        <Border Background="#000000" CornerRadius="12" BorderBrush="#802F2D" BorderThickness="2" Padding="20">
            <Grid>
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                <TextBlock x:Name="Msg" Grid.Row="1" Text="MSG_TXT" Foreground="White" FontSize="18" FontWeight="Bold" TextWrapping="Wrap" TextAlignment="Center" VerticalAlignment="Center"/>
                <Button Name="Btn_Ok" Grid.Row="2" Content="OK" Width="100" Height="32" Margin="0,15,0,0" Background="#802F2D" Foreground="White" BorderThickness="0" Cursor="Hand">
                    <Button.Template><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="5"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Button.Template>
                </Button>
            </Grid>
        </Border>
    </Window>
    """.replace("MSG_TXT", str(message))
        self.ui = XamlReader.Parse(xaml_str)
        self.ui.FindName("Btn_Ok").Click += (lambda s, e: self.ui.Close())

    def show(self): self.ui.ShowDialog()

def show_custom_alert(message):
    try:
        dialog = AlertWindow(message)
        dialog.show()
    except Exception as e:
        import traceback
        print("Alert Error: {} | Trace: {}".format(e, traceback.format_exc()))
        try: forms.alert(message)
        except: pass

def ensure_shared_parameters(add_foundations, add_columns):
    try:
        app = doc.Application
        sp_file = app.SharedParametersFilename
        if not sp_file or not os.path.exists(sp_file):
            temp_sp = os.path.join(os.path.expanduser('~'), "Documents", "Revit_SharedParams.txt")
            if not os.path.exists(temp_sp):
                with open(temp_sp, 'w') as f:
                    f.write("# Shared Params File\n*META\tVERSION\tMINVERSION\nMETA\t2\t1\n*GROUP\tID\tNAME\n*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")
            app.SharedParametersFilename = temp_sp
            
        sp_file_def = app.OpenSharedParameterFile()
        if not sp_file_def: return False
            
        group = sp_file_def.Groups.get_Item("Coordinates Group") or sp_file_def.Groups.Create("Coordinates Group")
        param_names = ["Coord_X", "Coord_Y", "Coord_Z"]
        param_defs = []
        for p_name in param_names:
            p_def = group.Definitions.get_Item(p_name)
            if not p_def:
                try: opt = DB.ExternalDefinitionCreationOptions(p_name, DB.SpecTypeId.String.Text)
                except: opt = DB.ExternalDefinitionCreationOptions(p_name, DB.ParameterType.Text)
                p_def = group.Definitions.Create(opt)
            param_defs.append(p_def)
            
        cat_set = app.Create.NewCategorySet()
        if add_foundations: cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFoundation))
        if add_columns: cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralColumns))
        
        bm = doc.ParameterBindings
        try: target_grp = DB.GroupTypeId.IdentityData
        except: target_grp = DB.BuiltInParameterGroup.PG_IDENTITY_DATA
        
        with revit.Transaction("Setup Params"):
            for p_def in param_defs:
                if not bm.get_Item(p_def):
                    bm.Insert(p_def, app.Create.NewInstanceBinding(cat_set), target_grp)
        return True
    except Exception as e:
        print("Param Error: {}".format(e))
        return False

def run():

    # Helper to get only types that are used in the model
    def get_used_types(category):
        # 1. Get all instances of this category
        instances = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
        # 2. Extract unique type IDs
        used_type_ids = set([i.GetTypeId() for i in instances])
        # 3. Get the symbols
        symbols = [doc.GetElement(tid) for tid in used_type_ids if tid != DB.ElementId.InvalidElementId]
        return symbols

    f_types = get_used_types(BuiltInCategory.OST_StructuralFoundation)
    c_types = get_used_types(BuiltInCategory.OST_StructuralColumns)
    
    if not f_types and not c_types:
        forms.alert("No modeled foundations or columns found in the project.")
        return

    ui = CoordinateToolWindow(f_types, c_types)
    res = ui.show()
    if not res: return

    # Phase 1: Ensure Parameters are created and bound
    if not ensure_shared_parameters(True, True):
        show_custom_alert("Failed to setup Shared Parameters.")
        # Continue anyway, they might exist already
    proj_loc = doc.ActiveProjectLocation
    updated_ids = set()
    with revit.Transaction("Update Coords"):
        for tid in res["ids"]:
            cat = doc.GetElement(tid).Category.Id
            # Use ACTIVE VIEW collector to match observed modeled elements (Avoiding hidden/background objects)
            els = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategoryId(cat).WhereElementIsNotElementType().ToElements()
            for el in [e for e in els if e.GetTypeId() == tid]:
                if el.Id in updated_ids: continue # Avoid double-counting
                
                loc = el.Location
                if not loc: continue
                xyz = loc.Point if hasattr(loc, "Point") else loc.Curve.GetEndPoint(0) if hasattr(loc, "Curve") else None
                if not xyz: continue
                
                if res["mode"] == "Survey Point":
                    # Use shared coordinate system
                    pos = proj_loc.GetProjectPosition(xyz)
                    raw_x, raw_y, raw_z = pos.EastWest, pos.NorthSouth, pos.Elevation
                else:
                    # Use internal project base point
                    pbp = DB.BasePoint.GetProjectBasePoint(doc)
                    pbp_pt = pbp.get_BoundingBox(None).Max
                    raw_x, raw_y, raw_z = xyz.X - pbp_pt.X, xyz.Y - pbp_pt.Y, xyz.Z - pbp_pt.Z
                
                fmt = "{:.3f}"
                factor = 304.8 if "Millimeters" in res["unit"] else 0.3048
                
                # Apply values
                updated = False
                for p_name, raw_val in zip(PARAMETER_NAMES, [raw_x, raw_y, raw_z]):
                    p = el.LookupParameter(p_name)
                    if p:
                        if p.StorageType == DB.StorageType.String:
                            p.Set(fmt.format(raw_val * factor))
                        else:
                            p.Set(raw_val) # Double/Length handles conversion native
                        updated = True
                if updated: updated_ids.add(el.Id)
    
    count = len(updated_ids)
    msg = "Successfully updated {} modeled elements.".format(count)
    show_custom_alert(msg)


if __name__ == "__main__":
    run()
