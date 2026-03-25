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
    BuiltInCategory,
    ElementTransformUtils,
    CopyPasteOptions,
    Transaction,
    ElementId,
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

# ---------------------------------------------------------------------------
# Available element categories to copy
# ---------------------------------------------------------------------------
CATEGORIES = [
    ("Walls",               BuiltInCategory.OST_Walls),
    ("Structural Columns",  BuiltInCategory.OST_StructuralColumns),
    ("Structural Framing",  BuiltInCategory.OST_StructuralFraming),
    ("Floors",              BuiltInCategory.OST_Floors),
    ("Doors",               BuiltInCategory.OST_Doors),
    ("Windows",             BuiltInCategory.OST_Windows),
    ("Roofs",               BuiltInCategory.OST_Roofs),
    ("Stairs",              BuiltInCategory.OST_Stairs),
    ("Railings",            BuiltInCategory.OST_StairsRailing),
]

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

        # Load logo image programmatically (avoids XAML URI issues)
        logo_ctrl = self.window.FindName("LogoImage")
        if logo_ctrl is not None and os.path.exists(LOGO_PATH):
            try:
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(LOGO_PATH, UriKind.Absolute)
                bmp.EndInit()
                logo_ctrl.Source = bmp
            except Exception:
                pass  # silently skip if logo can't load

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

        # Link checkboxes
        for li in self.link_instances:
            cb = Controls.CheckBox()
            cb.Content   = li.Name
            cb.Style     = check_style
            cb.IsChecked = False
            self.link_panel.Children.Add(cb)
            self._link_checkboxes.append(cb)
        if self._link_checkboxes:
            self._link_checkboxes[0].IsChecked = True

        # Category checkboxes
        for cat_name, bic in CATEGORIES:
            cb = Controls.CheckBox()
            cb.Content   = cat_name
            cb.Style     = check_style
            cb.IsChecked = False
            self.cat_panel.Children.Add(cb)
            self._cat_checkboxes.append((bic, cb))

        # Events
        self.lnk_all_btn.Click  += lambda s, e: self._set_links(True)
        self.lnk_none_btn.Click += lambda s, e: self._set_links(False)
        self.cat_all_btn.Click  += lambda s, e: self._set_cats(True)
        self.cat_none_btn.Click += lambda s, e: self._set_cats(False)
        self.copy_btn.Click     += self._on_copy
        self.cancel_btn.Click   += self._on_cancel

    def _set_links(self, state):
        for cb in self._link_checkboxes:
            cb.IsChecked = state

    def _set_cats(self, state):
        for _, cb in self._cat_checkboxes:
            cb.IsChecked = state

    def _on_copy(self, sender, args):
        self.selected_links = [
            self.link_instances[i]
            for i, cb in enumerate(self._link_checkboxes)
            if cb.IsChecked
        ]
        if not self.selected_links:
            CustomMessageBox.show(
                "Please select at least one Revit Link.",
                "No Link Selected"
            )
            return

        self.selected_cats = [bic for bic, cb in self._cat_checkboxes if cb.IsChecked]
        if not self.selected_cats:
            CustomMessageBox.show(
                "Please select at least one element type.",
                "No Element Types Selected"
            )
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
    def show(message, title="Message"):
        xaml_str = """
        <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="{title}"
            Width="400" SizeToContent="Height"
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

                    <!-- Header bar -->
                    <Border Grid.Row="0" Background="#111111" CornerRadius="8,8,0,0" Padding="15,10">
                        <TextBlock Text="{title}" Foreground="White" FontWeight="Bold" FontSize="14"/>
                    </Border>

                    <!-- Content area -->
                    <StackPanel Grid.Row="1" Margin="20,24,20,20">
                        <TextBlock Text="{message}" Foreground="#E0E0E0" FontSize="13" TextWrapping="Wrap" MaxWidth="350"/>
                    </StackPanel>

                    <!-- Action buttons -->
                    <StackPanel Grid.Row="2" HorizontalAlignment="Right" Margin="0,0,20,20">
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
        """.replace("{title}", str(title)).replace("{message}", str(message))

        from System.Windows.Markup import XamlReader
        window = XamlReader.Parse(xaml_str)
        
        ok_btn = window.FindName("OkBtn")
        def on_ok(sender, args):
            window.DialogResult = True
            window.Close()
        ok_btn.Click += on_ok
        
        window.ShowDialog()


# ---------------------------------------------------------------------------
# Revit helper functions
# ---------------------------------------------------------------------------
def get_link_instances(document):
    collector = FilteredElementCollector(document).OfClass(RevitLinkInstance)
    return [li for li in collector if li.GetLinkDocument() is not None]


def collect_elements_by_bic(link_doc, bic_list):
    ids = List[ElementId]()
    for bic in bic_list:
        col = (FilteredElementCollector(link_doc)
               .OfCategory(bic)
               .WhereElementIsNotElementType()
               .ToElementIds())
        for eid in col:
            ids.Add(eid)
    return ids


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

    for link_instance in ui.selected_links:
        link_doc  = link_instance.GetLinkDocument()
        transform = link_instance.GetTotalTransform()
        ids = collect_elements_by_bic(link_doc, ui.selected_cats)

        if ids.Count == 0:
            errors.append("'{}': no matching elements found.".format(link_instance.Name))
            continue

        t = Transaction(doc, "Copy from Link: {}".format(link_instance.Name))
        try:
            t.Start()
            copied = ElementTransformUtils.CopyElements(
                link_doc, ids, doc, transform, copy_options
            )
            t.Commit()
            total_copied += len(list(copied))
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            errors.append("'{}': {}".format(link_instance.Name, str(ex)))

    msg = "Copied {} element(s) from {} link(s).".format(
        total_copied, len(ui.selected_links))
    if errors:
        msg += "\n\nWarnings:\n" + "\n".join(errors)

    CustomMessageBox.show(
        msg,
        "Copy Complete" if not errors else "Copy Completed with Warnings"
    )


run()
