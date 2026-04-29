# -*- coding: utf-8 -*-
"""
Copy Elements from Revit Link
Copies selected element types from one or more chosen Revit Links into the current project.
Author: Udarie / Riyan Private Limited
"""
import clr
import os
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ElementTransformUtils,
    CopyPasteOptions,
    Transaction,
    ElementId,
    CategoryType,
    ElementMulticategoryFilter,
    IFailuresPreprocessor,
    FailureSeverity,
    FailureProcessingResult,
)
from System.Collections.Generic import List
from System import Uri, UriKind
from System.Windows.Media.Imaging import BitmapImage
import System.Windows.Controls as Controls
import System.Windows.Media as Media
import System.Windows as WPF
from pyrevit import revit

doc   = revit.doc
uidoc = revit.uidoc

# Logo path (same folder as this script)
SCRIPT_DIR = os.path.dirname(__file__)
LOGO_PATH  = os.path.join(SCRIPT_DIR, "logo.png")

# Brand colors (Riyan maroon)
MAROON      = "#7B2C2C"
MAROON_DARK = "#621F1F"
MAROON_LITE = "#9B3C3C"

def get_categories_in_links(link_instances):
    """Return sorted list of (IntegerValue, name) tuples for all
    model categories that appear in any of the provided link instances.
    Excludes non-physical and system categories."""
    
    EXCLUDE_CATS = {
        "Materials", "Material Assets", "Internal Origin",
        "Legend Components", "Lines", "HVAC Zones", "Rooms",
        "Spaces", "Areas", "Project Base Point", "Survey Point",
        "Raster Images", "Cameras", "Section Boxes", "Scope Boxes",
        "Mass", "Detail Items", "Pads", "Location Data",
        "Site", "Curtain Systems", "Shaft Openings",
        "Project Information", "Primary Contours", "Pipe Segments",
        "Railing Rail Path Extension Lines", "Topography", "Center lines"
    }

    seen = {} # {IntegerValue: Name}
    for li in link_instances:
        link_doc = li.GetLinkDocument()
        if not link_doc:
            continue
            
        # Only get actual modeled instances (no types, no view-specific 2D detailing)
        collector = FilteredElementCollector(link_doc)\
                    .WhereElementIsNotElementType()\
                    .WhereElementIsViewIndependent()
                    
        for elem in collector:
            cat = elem.Category
            if cat and cat.CategoryType == CategoryType.Model:
                cat_name = cat.Name
                # Exclude junk/system categories and any hidden ones starting with <
                if cat_name not in EXCLUDE_CATS and not cat_name.startswith("<"):
                    cat_val = cat.Id.IntegerValue
                    if cat_val not in seen:
                        seen[cat_val] = cat_name
                        
    # Return as list of (IntegerValue, Name)
    return sorted(seen.items(), key=lambda x: x[1])

# ---------------------------------------------------------------------------
# WPF Window XAML  — white background, maroon accents
# ---------------------------------------------------------------------------
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Copy Elements from Revit Links"
    Width="480" Height="660"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    FontFamily="Segoe UI"
    Background="Black">

    <Window.Resources>

        <!-- Primary (maroon) Button -->
        <Style TargetType="Button" x:Key="PrimaryBtn">
            <Setter Property="Background"      Value="#7B2C2C"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="22,9"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#621F1F"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#4E1818"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary (outline) Button -->
        <Style TargetType="Button" x:Key="SecondaryBtn">
            <Setter Property="Background"      Value="Black"/>
            <Setter Property="Foreground"      Value="#9B3C3C"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush"     Value="#7B2C2C"/>
            <Setter Property="Padding"         Value="10,4"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1A1A1A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Cancel Button -->
        <Style TargetType="Button" x:Key="CancelBtn">
            <Setter Property="Background"      Value="Black"/>
            <Setter Property="Foreground"      Value="#A0A0A0"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush"     Value="#404040"/>
            <Setter Property="Padding"         Value="22,9"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1A1A1A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- CheckBox -->
        <Style TargetType="CheckBox" x:Key="StyledCheck">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="FontSize"   Value="13"/>
            <Setter Property="Margin"     Value="0,4,0,4"/>
            <Setter Property="Cursor"     Value="Hand"/>
        </Style>

        <!-- Section label -->
        <Style TargetType="TextBlock" x:Key="SectionLabel">
            <Setter Property="FontSize"   Value="10"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#9B3C3C"/>
            <Setter Property="Margin"     Value="0,0,0,6"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- header bar with logo -->
            <RowDefinition Height="*"/>      <!-- content -->
        </Grid.RowDefinitions>

        <!-- Header bar -->
        <Border Grid.Row="0" Background="#111111" BorderBrush="#7B2C2C" BorderThickness="0,0,0,2" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="Copy from Revit Links"
                               FontSize="18" FontWeight="Bold"
                               Foreground="White"/>
                    <TextBlock Text="Select links and element types to copy into this project."
                               FontSize="11" Foreground="#A0A0A0" Margin="0,3,0,0"/>
                </StackPanel>
                <!-- Logo (loaded in code-behind) -->
                <Image x:Name="LogoImage" Grid.Column="1"
                       Height="52" Width="90"
                       HorizontalAlignment="Right"
                       VerticalAlignment="Center"
                       Margin="16,0,0,0"
                       RenderOptions.BitmapScalingMode="HighQuality"/>
            </Grid>
        </Border>

        <!-- Content area -->
        <Grid Grid.Row="1" Margin="24,20,24,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>   <!-- links label + buttons -->
                <RowDefinition Height="120"/>    <!-- links list -->
                <RowDefinition Height="Auto"/>   <!-- divider -->
                <RowDefinition Height="Auto"/>   <!-- categories label + buttons -->
                <RowDefinition Height="*"/>      <!-- categories list -->
                <RowDefinition Height="Auto"/>   <!-- action buttons -->
            </Grid.RowDefinitions>

            <!-- Links label + Select/Clear -->
            <Grid Grid.Row="0" Margin="0,0,0,6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="REVIT LINKS"
                           Style="{StaticResource SectionLabel}"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="LinkSelectAllBtn" Content="Select All"
                            Style="{StaticResource SecondaryBtn}" Margin="0,0,6,0"/>
                    <Button x:Name="LinkClearAllBtn"  Content="Clear"
                            Style="{StaticResource SecondaryBtn}"/>
                </StackPanel>
            </Grid>

            <!-- Links checkbox list box -->
            <Border Grid.Row="1"
                    BorderBrush="#333333" BorderThickness="1"
                    CornerRadius="5" Background="#111111">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10,6">
                    <StackPanel x:Name="LinkPanel"/>
                </ScrollViewer>
            </Border>

            <!-- Divider -->
            <Border Grid.Row="2" Height="1" Background="#333333" Margin="0,16,0,16"/>

            <!-- Categories label + Select/Clear -->
            <Grid Grid.Row="3" Margin="0,0,0,6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="ELEMENT TYPES"
                           Style="{StaticResource SectionLabel}"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="CatSelectAllBtn" Content="Select All"
                            Style="{StaticResource SecondaryBtn}" Margin="0,0,6,0"/>
                    <Button x:Name="CatClearAllBtn"  Content="Clear"
                            Style="{StaticResource SecondaryBtn}"/>
                </StackPanel>
            </Grid>

            <!-- Category checkboxes -->
            <Border Grid.Row="4"
                    BorderBrush="#333333" BorderThickness="1"
                    CornerRadius="5" Background="#111111">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10,6">
                    <StackPanel x:Name="CategoryPanel"/>
                </ScrollViewer>
            </Border>

            <!-- Action buttons -->
            <StackPanel Grid.Row="5" Orientation="Horizontal"
                        HorizontalAlignment="Right" Margin="0,18,0,0">
                <Button x:Name="CancelBtn2" Content="Cancel"
                        Style="{StaticResource CancelBtn}" Margin="0,0,10,0"/>
                <Button x:Name="CopyBtn"   Content="Copy Elements"
                        Style="{StaticResource PrimaryBtn}"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"""

# ---------------------------------------------------------------------------
# Window Logic
# ---------------------------------------------------------------------------
class CopyFromLinkWindow(object):
    def __init__(self, link_instances):
        self.link_instances   = link_instances
        self.selected_links   = []
        self.selected_cats    = []
        self.result           = False
        self._link_checkboxes = []
        self._cat_checkboxes  = []

        from System.Windows.Markup import XamlReader
        self.window = XamlReader.Parse(XAML)

        # Load logo image
        logo_ctrl = self.window.FindName("LogoImage")
        if logo_ctrl is not None and os.path.exists(LOGO_PATH):
            try:
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(LOGO_PATH, UriKind.Absolute)
                bmp.EndInit()
                logo_ctrl.Source = bmp
            except Exception:
                pass

        # Named controls
        self.link_panel   = self.window.FindName("LinkPanel")
        self.cat_panel    = self.window.FindName("CategoryPanel")
        self.copy_btn     = self.window.FindName("CopyBtn")
        self.cancel_btn   = self.window.FindName("CancelBtn2")
        self.lnk_all_btn  = self.window.FindName("LinkSelectAllBtn")
        self.lnk_none_btn = self.window.FindName("LinkClearAllBtn")
        self.cat_all_btn  = self.window.FindName("CatSelectAllBtn")
        self.cat_none_btn = self.window.FindName("CatClearAllBtn")

        check_style = self.window.FindResource("StyledCheck")

        # Create link checkboxes
        for li in self.link_instances:
            cb = Controls.CheckBox()
            cb.Content   = li.Name
            cb.Style     = check_style
            cb.IsChecked = False
            self.link_panel.Children.Add(cb)
            self._link_checkboxes.append(cb)
            
            # Register events
            cb.Checked   += self._on_link_toggle
            cb.Unchecked += self._on_link_toggle

        # Default select first link and refresh
        if self._link_checkboxes:
            self._link_checkboxes[0].IsChecked = True
        self._refresh_categories()

        # Button events
        self.lnk_all_btn.Click  += lambda s, e: self._set_links_state(True)
        self.lnk_none_btn.Click += lambda s, e: self._set_links_state(False)
        self.cat_all_btn.Click  += lambda s, e: self._set_cats_state(True)
        self.cat_none_btn.Click += lambda s, e: self._set_cats_state(False)
        self.copy_btn.Click     += self._on_copy
        self.cancel_btn.Click   += self._on_cancel

    def _refresh_categories(self):
        self.cat_panel.Children.Clear()
        self._cat_checkboxes = []
        
        selected_link_objs = []
        for i, cb in enumerate(self._link_checkboxes):
            if cb.IsChecked:
                selected_link_objs.append(self.link_instances[i])
        
        if not selected_link_objs:
            lbl = Controls.TextBlock()
            lbl.Text = "Select at least one link to see element types."
            lbl.Foreground = Media.SolidColorBrush(Media.Colors.Gray)
            lbl.FontSize = 11
            self.cat_panel.Children.Add(lbl)
            return

        cats = get_categories_in_links(selected_link_objs)
        check_style = self.window.FindResource("StyledCheck")
        
        for cat_val, cat_name in cats:
            cb = Controls.CheckBox()
            cb.Content   = cat_name
            cb.Style     = check_style
            cb.IsChecked = False
            self.cat_panel.Children.Add(cb)
            # Store the integer value instead of the ElementId object
            self._cat_checkboxes.append((cat_val, cb))

    def _on_link_toggle(self, sender, args):
        self._refresh_categories()

    def _set_links_state(self, state):
        for cb in self._link_checkboxes:
            cb.IsChecked = state
        self._refresh_categories()

    def _set_cats_state(self, state):
        for _, cb in self._cat_checkboxes:
            cb.IsChecked = state

    def _on_copy(self, sender, args):
        self.selected_links = []
        for i, cb in enumerate(self._link_checkboxes):
            if cb.IsChecked:
                self.selected_links.append(self.link_instances[i])
                
        if not self.selected_links:
            CustomMessageBox.show("Please select at least one Revit Link.", "No Link Selected")
            return

        self.selected_cats = [cat_val for cat_val, cb in self._cat_checkboxes if cb.IsChecked]
        if not self.selected_cats:
            CustomMessageBox.show("Please select at least one element type.", "No Element Types Selected")
            return

        # Hosted Elements Check
        walls_bic_val = -2000011 # OST_Walls
        doors_bic_val = -2000023 # OST_Doors
        wins_bic_val  = -2000014 # OST_Windows
        
        has_doors = doors_bic_val in self.selected_cats
        has_wins  = wins_bic_val  in self.selected_cats
        has_walls = walls_bic_val in self.selected_cats
        
        if (has_doors or has_wins) and not has_walls:
            warn_msg = "You have selected Doors or Windows WITHOUT selecting Walls.\n\n" \
                       "Revit cannot copy these elements into the project if their host Walls are not also selected.\n\n" \
                       "Do you want to continue anyway? (Recommended: go back and select Walls)"
            if not CustomMessageBox.show(warn_msg, "Missing Hosts Warning", yes_no=True):
                return

        self.result = True
        self.window.Close()

    def _on_cancel(self, sender, args):
        self.result = False
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self.result


# ---------------------------------------------------------------------------
# Custom Dialogue / Message Box
# ---------------------------------------------------------------------------
class CustomMessageBox:
    @staticmethod
    def show(message, title="Message", yes_no=False):
        """Show a custom message box. If yes_no is True, returns True for Yes, False for No."""
        xaml_str = """
        <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Width="460" SizeToContent="Height"
            WindowStartupLocation="CenterScreen"
            ResizeMode="NoResize"
            FontFamily="Segoe UI"
            Background="Black"
            WindowStyle="None"
            AllowsTransparency="True">
            
            <Border BorderBrush="#7B2C2C" BorderThickness="1" CornerRadius="8" Background="Black">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#111111" CornerRadius="8,8,0,0" Padding="15,10">
                        <TextBlock x:Name="TitleBlock" Foreground="White" FontWeight="Bold" FontSize="14"/>
                    </Border>

                    <StackPanel Grid.Row="1" Margin="20,24,20,20">
                        <TextBlock x:Name="MsgBlock" Foreground="#E0E0E0" FontSize="13" TextWrapping="Wrap" MaxWidth="410"/>
                    </StackPanel>

                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,20,20">
                        <Button x:Name="NoBtn" Content="Cancel" Width="80" Height="28" Cursor="Hand" Foreground="#A0A0A0" Margin="0,0,10,0" Visibility="Collapsed">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border x:Name="border" Background="#222222" CornerRadius="4">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                        <Button x:Name="OkBtn" Content="OK" Width="80" Height="28" Cursor="Hand" Foreground="White" FontWeight="SemiBold">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border x:Name="border" Background="#7B2C2C" CornerRadius="4">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="border" Property="Background" Value="#621F1F"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                    </StackPanel>
                </Grid>
            </Border>
        </Window>
        """
        from System.Windows.Markup import XamlReader
        window = XamlReader.Parse(xaml_str)
        
        window.Title = str(title)
        window.FindName("TitleBlock").Text = str(title)
        window.FindName("MsgBlock").Text = str(message)
        
        ok_btn = window.FindName("OkBtn")
        no_btn = window.FindName("NoBtn")
        
        if yes_no:
            ok_btn.Content = "Continue"
            no_btn.Visibility = WPF.Visibility.Visible
            
        result = [False]
        
        def on_ok(sender, args):
            result[0] = True
            window.DialogResult = True
            window.Close()
            
        def on_no(sender, args):
            result[0] = False
            window.DialogResult = False
            window.Close()
            
        ok_btn.Click += on_ok
        no_btn.Click += on_no
        
        window.ShowDialog()
        return result[0]


# ---------------------------------------------------------------------------
# Failure Handling (Suppress warnings like "inserts outside of hosts")
# ---------------------------------------------------------------------------
class CopyWarningsSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        fail_list = failuresAccessor.GetFailureMessages()
        for f in fail_list:
            if f.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(f)
        return FailureProcessingResult.Continue


# ---------------------------------------------------------------------------
# Revit helper functions
# ---------------------------------------------------------------------------
def get_link_instances(document):
    collector = FilteredElementCollector(document).OfClass(RevitLinkInstance)
    return [li for li in collector if li.GetLinkDocument() is not None]


def collect_elements_by_categories(link_doc, cat_vals):
    """Collect all element IDs in link_doc for the given category integer values.
    Returns a simple list of ElementIds."""
    found_ids = List[ElementId]()
    
    # Iterate manually for 100% reliability, enforcing view independence
    all_elements = FilteredElementCollector(link_doc)\
                   .WhereElementIsNotElementType()\
                   .WhereElementIsViewIndependent()
                   
    for elem in all_elements:
        cat = elem.Category
        if cat and cat.Id.IntegerValue in cat_vals:
            found_ids.Add(elem.Id)
            
    return found_ids


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def run():
    links = get_link_instances(doc)
    if not links:
        CustomMessageBox.show(
            "No loaded Revit Links found in this project.\n\n"
            "Please load at least one Revit Link and try again.",
            "No Revit Links"
        )
        return

    ui = CopyFromLinkWindow(links)
    if not ui.show():
        return

    total_copied = 0
    errors       = []
    copy_options = CopyPasteOptions()

    summary_parts = []
    
    for link_instance in ui.selected_links:
        link_doc  = link_instance.GetLinkDocument()
        if not link_doc:
            continue
            
        transform = link_instance.GetTotalTransform()
        
        # 1. Collect
        ids_to_copy = collect_elements_by_categories(link_doc, ui.selected_cats)
        
        if ids_to_copy.Count == 0:
            continue

        # 2. Attempt Copy
        t = Transaction(doc, "Copy from Link: {}".format(link_instance.Name))
        try:
            # Suppress warnings (especially for hosted elements)
            opts = t.GetFailureHandlingOptions()
            opts.SetFailuresPreprocessor(CopyWarningsSwallower())
            t.SetFailureHandlingOptions(opts)
            
            t.Start()
            copied_ids = ElementTransformUtils.CopyElements(
                link_doc, ids_to_copy, doc, transform, copy_options
            )
            t.Commit()
            
            total_copied += len(list(copied_ids))
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            errors.append("'{}': {}".format(link_instance.Name, str(ex)))

    if total_copied > 0 or not errors:
        msg = "Elements copied successfully!\n\n"
        msg += "Total elements copied: {}".format(total_copied)
        if total_copied == 0:
            msg = "No elements were copied.\n\n"
            msg += "Note: If you were copying Doors or Windows, please make sure to select their host Walls as well."
    else:
        msg = "Copy process finished with errors."

    if errors:
        msg += "\n\nWarnings/Errors:\n" + "\n".join(errors)

    CustomMessageBox.show(msg, "Copy Elements Result")


run()
