# -*- coding: utf-8 -*-
"""
Smart Connect
Author: Antigravity AI
Select a main element and quickly place structural connections to intersecting elements.
"""
__title__ = "Smart Connect"
__author__ = "Antigravity AI"

import clr
import os
import traceback

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import FilteredElementCollector, ElementIntersectsElementFilter, Transaction, BuiltInCategory, ElementId
from Autodesk.Revit.DB.Structure import StructuralConnectionHandler, StructuralConnectionHandlerType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import System.Windows.Controls as Controls
import System.Windows.Media as Media
import System.Windows as WPF
from pyrevit import revit, DB, UI, forms
from System.Windows.Interop import WindowInteropHelper

doc = revit.doc
uidoc = revit.uidoc
SCRIPT_DIR = os.path.dirname(__file__)

XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Smart Connect Tool"
    Width="500" Height="800"
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
                        <Border BorderBrush="#802F2D" BorderThickness="1" CornerRadius="3" Padding="4,1" Background="{TemplateBinding Background}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2A0A0A"/>
                </Trigger>
            </Style.Triggers>
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
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="Border" Property="Background" Value="#1A1A1A" />
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,0,30,0" VerticalAlignment="Center" HorizontalAlignment="Left" />
                            <Popup x:Name="PART_Popup" AllowsTransparency="true" Placement="Bottom" IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}" Focusable="false">
                                <Border x:Name="DropDownBorder" Background="#000000" BorderBrush="#802F2D" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}">
                                    <ScrollViewer CanContentScroll="true" MaxHeight="250">
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
                <RowDefinition Height="30"/>  <!-- Title Bar -->
                <RowDefinition Height="120"/> <!-- Header -->
                <RowDefinition Height="*"/>   <!-- Content -->
                <RowDefinition Height="80"/>  <!-- Footer -->
            </Grid.RowDefinitions>

            <!-- Title Bar -->
            <Border x:Name="HeaderBar" Grid.Row="0" Background="#1A1A1A">
                <Grid Margin="10,0">
                    <TextBlock Text="Smart Connect Tool" Foreground="#AAAAAA" VerticalAlignment="Center" FontSize="11"/>
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
                        <TextBlock Text="Smart Connect" FontSize="26" FontWeight="Bold" Foreground="White"/>
                        <TextBlock x:Name="Txt_MainElement" Text="Main Element: None" FontSize="11" Foreground="#888888" Margin="0,5,0,0"/>
                    </StackPanel>
                    <Image x:Name="UI_Logo" Grid.Column="1" Width="85" Height="85" Stretch="Uniform" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </Grid>
            </Border>

            <Separator Grid.Row="1" VerticalAlignment="Bottom" Background="#802F2D" Height="1" Margin="10,0"/>

            <!-- Main Content -->
            <StackPanel Grid.Row="2" Margin="25,15">
                
                <StackPanel Margin="0,0,0,15">
                    <TextBlock Text="Connection Type" Style="{StaticResource SubLabel}" Margin="0,0,0,5"/>
                    <ComboBox x:Name="Combo_ConnectionType" Style="{StaticResource BlackCombo}"/>
                    <CheckBox x:Name="Chk_FlipOrder" Content="Flip Main/Target cutting order (Check if main beam gets cut)" Foreground="#AAAAAA" Margin="0,8,0,0" FontSize="11" Cursor="Hand"/>
                </StackPanel>

                <Grid Margin="0,5,0,5">
                    <TextBlock Text="INTERSECTING ELEMENTS" Style="{StaticResource SectionHeader}"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="Btn_SAll" Content="Select All" Style="{StaticResource TinyBtn}" Margin="0,0,5,0"/>
                        <Button x:Name="Btn_Clr" Content="Clear" Style="{StaticResource TinyBtn}"/>
                    </StackPanel>
                </Grid>
                
                <Border Height="360" BorderBrush="#333333" BorderThickness="1" Margin="0,0,0,10">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="Panel_Elements" Margin="10"/>
                    </ScrollViewer>
                </Border>
                <TextBlock x:Name="Txt_Status" Text="0 elements selected" Foreground="#AAAAAA" FontSize="11" HorizontalAlignment="Right"/>
            </StackPanel>

            <!-- Footer -->
            <Border Grid.Row="3" Background="#0A0A0A" Padding="25,0">
                <Grid VerticalAlignment="Center">
                    <Button x:Name="Btn_Cancel" Content="Cancel" Width="100" Height="28" HorizontalAlignment="Left" Background="#2A2A2A" Foreground="White" BorderThickness="0" Cursor="Hand">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border Background="{TemplateBinding Background}" CornerRadius="5">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <Button x:Name="Btn_Run" Content="Place Connections" Width="140" Height="28" HorizontalAlignment="Right" Background="#802F2D" Foreground="White" BorderThickness="0" Cursor="Hand">
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

# Try to find a logo in other tools
MASTER_LOGO = os.path.abspath(os.path.join(SCRIPT_DIR, r"..\..\Coordination.panel\Copy from Link.pushbutton\logo.png"))

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

class ConnectionWindow(WPF.Window):
    def __init__(self, main_element, intersecting_elements, connection_types):
        from System.Windows.Markup import XamlReader
        self.ui = XamlReader.Parse(XAML)
        
        self.main_element = main_element
        self.intersecting_elements = intersecting_elements
        
        # UI Elements
        self.header_bar = self.ui.FindName("HeaderBar")
        self.combo_type = self.ui.FindName("Combo_ConnectionType")
        self.panel_elements = self.ui.FindName("Panel_Elements")
        self.txt_status = self.ui.FindName("Txt_Status")
        self.btn_run = self.ui.FindName("Btn_Run")
        self.txt_main = self.ui.FindName("Txt_MainElement")
        
        # Setup Text
        try:
            main_cat = main_element.Category.Name if main_element.Category else "Unknown"
        except:
            main_cat = "Unknown"
            
        try:
            main_name = main_element.Name
        except:
            main_name = "Unnamed"
            
        self.txt_main.Text = "Main Element: {} - {} [{}]".format(main_cat, main_name, main_element.Id)
        
        # Events
        self.header_bar.MouseLeftButtonDown += (lambda s, e: self.ui.DragMove())
        self.ui.FindName("Btn_Close").Click += self.OnCancel
        self.ui.FindName("Btn_Cancel").Click += self.OnCancel
        self.ui.FindName("Btn_SAll").Click += self.OnSelectAll
        self.ui.FindName("Btn_Clr").Click += self.OnClear
        self.btn_run.Click += self.OnRun
        
        # Populate Connection Types
        for ct in connection_types:
            item = Controls.ComboBoxItem()
            
            # Get the connection name
            try:
                # Try standard property
                conn_name = ct.Name
                # Optionally add family name if it exists and is different
                try:
                    fam_name = ct.FamilyName
                    if fam_name and fam_name != conn_name:
                        conn_name = "{} - {}".format(fam_name, conn_name)
                except:
                    pass
            except:
                # Fallback to BuiltInParameter if property fails
                try:
                    param = ct.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    conn_name = param.AsString() if param else ("Connection Type " + str(ct.Id))
                except:
                    conn_name = "Connection Type " + str(ct.Id)
                    
            if not conn_name:
                conn_name = "Connection Type " + str(ct.Id)
                
            item.Content = conn_name
            item.Tag = ct
            self.combo_type.Items.Add(item)
            
        if connection_types:
            self.combo_type.SelectedIndex = 0
            
        # Populate List
        self.checkboxes = []
        for el in self.intersecting_elements:
            try:
                cat_name = el.Category.Name if el.Category else "Unknown"
            except:
                cat_name = "Unknown"
                
            try:
                el_name = el.Name
            except:
                el_name = "Unnamed"
                
            disp_name = "{} - {} [{}]".format(cat_name, el_name, el.Id)
            
            cb = Controls.CheckBox()
            cb.Content = disp_name
            cb.Foreground = Media.Brushes.White
            cb.IsChecked = True
            cb.Margin = WPF.Thickness(0,2,0,2)
            cb.Tag = el
            cb.Checked += self.UpdateStatus
            cb.Unchecked += self.UpdateStatus
            
            self.panel_elements.Children.Add(cb)
            self.checkboxes.append(cb)
            
        self.UpdateStatus(None, None)
        load_logo(self.ui.FindName("UI_Logo"), MASTER_LOGO)
        
        try:
            helper = WindowInteropHelper(self.ui)
            helper.Owner = revit.handle
        except: pass
        
        self.result = None

    def UpdateStatus(self, sender, args):
        count = sum(1 for cb in self.checkboxes if cb.IsChecked == True)
        self.txt_status.Text = "{} elements selected".format(count)

    def OnSelectAll(self, sender, args):
        for cb in self.checkboxes: cb.IsChecked = True

    def OnClear(self, sender, args):
        for cb in self.checkboxes: cb.IsChecked = False

    def OnCancel(self, sender, args):
        self.ui.Close()

    def OnRun(self, sender, args):
        selected_item = self.combo_type.SelectedItem
        if not selected_item:
            forms.alert("Please select a structural connection type.")
            return

        selected_elements = []
        for cb in self.checkboxes:
            if cb.IsChecked == True:
                selected_elements.append(cb.Tag)
                
        if not selected_elements:
            forms.alert("Please select at least one intersecting element.")
            return
            
        flip_checked = self.ui.FindName("Chk_FlipOrder").IsChecked
        self.result = {"elements": selected_elements, "conn_type": selected_item.Tag, "flip": flip_checked}
        self.ui.Close()

    def show(self):
        self.ui.ShowDialog()
        return self.result

def show_custom_alert(message):
    from System.Windows.Markup import XamlReader
    xaml_str = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Success" Width="450" Height="250" WindowStartupLocation="CenterScreen"
            AllowsTransparency="True" WindowStyle="None" Background="Transparent" Topmost="True">
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
    """.replace("MSG_TXT", str(message).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    try:
        ui = XamlReader.Parse(xaml_str)
        ui.FindName("Btn_Ok").Click += (lambda s, e: ui.Close())
        ui.ShowDialog()
    except:
        forms.alert(message)

def get_structural_connection_types():
    types = list(FilteredElementCollector(doc).OfClass(StructuralConnectionHandlerType).WhereElementIsElementType().ToElements())
    try:
        mods = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnectionModifiers).WhereElementIsElementType().ToElements()
        existing_ids = [t.Id for t in types]
        for m in mods:
            if m.Id not in existing_ids:
                types.append(m)
    except: pass
    return types

def run():
    # 1. User selects main element
    selection = revit.get_selection()
    main_el = None
    if len(selection) == 1:
        main_el = selection.first
    elif len(selection) == 0:
        try:
            with forms.WarningBar(title="Select the main element (e.g. Concrete Beam, Steel Column)"):
                ref = uidoc.Selection.PickObject(ObjectType.Element, "Select the main element to connect to")
                main_el = doc.GetElement(ref)
        except Exception:
            return # User cancelled selection
    else:
        forms.alert("Please select exactly ONE main element first, or run the tool with nothing selected to pick one.")
        return
        
    if not main_el: return

    # 2. Find intersecting elements
    # Using bounding box intersection filtering to get potential crossing elements
    try:
        bbox = main_el.get_BoundingBox(None)
        if not bbox:
            # Fallback if element has no bounding box
            filter = ElementIntersectsElementFilter(main_el)
        else:
            # BoundingBox is faster and more reliable for general crossing
            outline = DB.Outline(bbox.Min, bbox.Max)
            filter = DB.BoundingBoxIntersectsFilter(outline)
            
        intersecting_candidate = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().WherePasses(filter).ToElements()
    except Exception as e:
        forms.alert("Error detecting intersections: {}".format(e))
        return

    # Filter out valid connectable items (Structural Framing, Columns, Walls, Foundations)
    valid_categories = {
        BuiltInCategory.OST_StructuralFraming.value__,
        BuiltInCategory.OST_StructuralColumns.value__,
        BuiltInCategory.OST_StructuralFoundation.value__,
        BuiltInCategory.OST_Walls.value__,
        BuiltInCategory.OST_Floors.value__
    }
    
    crossing_elements = []
    for el in intersecting_candidate:
        if el.Id == main_el.Id: continue # Skip itself
        if not el.Category: continue
        if el.Category.Id.IntegerValue in valid_categories:
            crossing_elements.append(el)

    if not crossing_elements:
        forms.alert("No valid structural elements cross the selected element in this view.")
        return

    conn_types = get_structural_connection_types()
    if not conn_types:
        forms.alert("No Structural Connections are loaded in this project. Please load them via the 'Steel' tab in Revit.")
        return

    # 3. Show UI
    window = ConnectionWindow(main_el, crossing_elements, conn_types)
    res = window.show()
    
    if not res: return
    
    targets = res["elements"]
    conn_type = res["conn_type"]
    flip = res.get("flip", False)

    # 4. Create connections
    success_count = 0
    failed_count = 0
    errors = []

    with Transaction(doc, "Smart Connect: Place Connections") as t:
        t.Start()
        for target in targets:
            try:
                # StructuralConnectionHandler.Create requires a List[ElementId]
                # The order is critical. For Modifiers like Notch/Cope, the FIRST element
                # is the one that gets cut (Primary). The SECOND element is the cutting tool (Secondary).
                from System.Collections.Generic import List
                id_list = List[ElementId]()
                
                if flip:
                    id_list.Add(main_el.Id)
                    id_list.Add(target.Id)
                else:
                    id_list.Add(target.Id)
                    id_list.Add(main_el.Id)
                
                StructuralConnectionHandler.Create(doc, id_list, conn_type.Id)
                success_count += 1
            except Exception as e:
                failed_count += 1
                errors.append(str(e))
        t.Commit()
        
    msg = "Successfully placed {} connections of type '{}'.".format(success_count, getattr(conn_type, "Name", "Unknown"))
    if failed_count > 0:
        msg += "\n\nFailed to place {} connections.\nCheck if geometry is valid for this connection type.".format(failed_count)

    show_custom_alert(msg)

if __name__ == "__main__":
    run()
