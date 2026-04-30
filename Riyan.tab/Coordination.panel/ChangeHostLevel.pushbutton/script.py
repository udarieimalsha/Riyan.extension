"""ChangeHost / Move Level
Provides two features: 
1. Re-hosts selected elements to a new level without changing absolute elevation.
2. Changes a level's actual elevation without moving the elements hosted on it.
"""

__title__ = "Change\nHost/Level"
__author__ = "Chalana"

import os
from pyrevit import revit, DB, UI, forms
from pyrevit.forms import WPFWindow
import traceback
import clr
import System
from System.Collections.Generic import List
clr.AddReference("System.Windows.Presentation")
clr.AddReference("System.Drawing")
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Media import Brushes, Color, SolidColorBrush
from System import Uri

doc = revit.doc
app = doc.Application
uidoc = revit.uidoc

# Known level and offset parameters
BASE_LEVEL_PARAM_NAMES = [
    "FAMILY_BASE_LEVEL_PARAM",
    "WALL_BASE_CONSTRAINT",
    "INSTANCE_REFERENCE_LEVEL_PARAM",
    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
    "SCHEDULE_LEVEL_PARAM",
    "LEVEL_PARAM",
    "ROOF_BASE_LEVEL_PARAM",
    "RBS_START_LEVEL_PARAM",
    "STAIRS_BASE_LEVEL_PARAM"
]

BASE_OFFSET_PARAM_NAMES = [
    "FAMILY_BASE_LEVEL_OFFSET_PARAM",
    "WALL_BASE_OFFSET",
    "INSTANCE_FREE_HOST_OFFSET_PARAM",
    "ROOF_LEVEL_OFFSET_PARAM",
    "RBS_START_LEVEL_OFFSET_PARAM",
    "FLOOR_HEIGHTABOVELEVEL_PARAM",
    "CEILING_HEIGHTABOVELEVEL_PARAM"
]

TOP_LEVEL_PARAM_NAMES = [
    "FAMILY_TOP_LEVEL_PARAM",
    "WALL_HEIGHT_TYPE",
    "STAIRS_TOP_LEVEL_PARAM"
]

TOP_OFFSET_PARAM_NAMES = [
    "FAMILY_TOP_LEVEL_OFFSET_PARAM",
    "WALL_TOP_OFFSET"
]

def get_bips(name_list):
    bips = []
    for name in name_list:
        bip = getattr(DB.BuiltInParameter, name, None)
        if bip is not None:
            bips.append(bip)
    return bips

BASE_LEVEL_PARAMS = get_bips(BASE_LEVEL_PARAM_NAMES)
BASE_OFFSET_PARAMS = get_bips(BASE_OFFSET_PARAM_NAMES)
TOP_LEVEL_PARAMS = get_bips(TOP_LEVEL_PARAM_NAMES)
TOP_OFFSET_PARAMS = get_bips(TOP_OFFSET_PARAM_NAMES)

def get_param_from_list(element, bip_list):
    """For writing - skips read-only params."""
    for bip in bip_list:
        p = element.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            return p
    return None

def get_level_id_from_list(element, bip_list):
    """For reading level IDs only - includes read-only params (e.g. beams)."""
    for bip in bip_list:
        p = element.get_Parameter(bip)
        if p is not None and p.StorageType == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                return eid
    
    # Fallback: Search all parameters for any Level
    for p in element.Parameters:
        if p.StorageType == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                test_el = doc.GetElement(eid)
                if test_el and isinstance(test_el, DB.Level):
                    return eid
                    
    return DB.ElementId.InvalidElementId

def get_any_level_param(element, bip_list):
    """Returns any level param (including read-only) for force-setting beam levels."""
    for bip in bip_list:
        p = element.get_Parameter(bip)
        if p is not None and p.StorageType == DB.StorageType.ElementId:
            return p
    return None

class AlertWindow(WPFWindow):
    def __init__(self, xaml_file_name, message):
        WPFWindow.__init__(self, xaml_file_name)
        self.MessageText.Text = message

    def TitleBar_MouseDown(self, sender, args):
        if args.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
            self.DragMove()

    def OKButton_Click(self, sender, args):
        self.DialogResult = True
        self.Close()

def show_custom_alert(message):
    try:
        xaml_file = os.path.join(os.path.dirname(__file__), "alert.xaml")
        if os.path.exists(xaml_file):
            alert = AlertWindow(xaml_file, message)
            alert.ShowDialog()
        else:
            forms.alert(message)
    except:
        forms.alert(message)

# ----------- Unit Conversion logic -------------
def get_length_unit(d):
    units = d.GetUnits()
    try:
        # Revit 2022+ API
        return units.GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()
    except:
        pass
    try:
        # Revit 2021 and older API
        return units.GetFormatOptions(DB.UnitType.UT_Length).DisplayUnits
    except:
        return None

def to_display_units(internal_val, d):
    unit = get_length_unit(d)
    if unit is not None:
        try:
            return DB.UnitUtils.ConvertFromInternalUnits(internal_val, unit)
        except:
            pass
    return internal_val

def to_internal_units(display_val, d):
    unit = get_length_unit(d)
    if unit is not None:
        try:
            return DB.UnitUtils.ConvertToInternalUnits(display_val, unit)
        except:
            pass
    return display_val

# ----------- Backend Processing logic -------------

def process_elements(selected_elements, target_level, source_level=None):
    if not selected_elements:
        show_custom_alert("No elements selected to re-host.")
        return
        
    if not target_level:
        show_custom_alert("No target level selected.")
        return
        
    t = DB.Transaction(doc, "Rehost Elements")
    t.Start()
    
    count = 0
    try:
        for el in selected_elements:
            is_beam = el.Category and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_StructuralFraming)
            changed = False

            # Base logic
            base_lvl_param = get_param_from_list(el, BASE_LEVEL_PARAMS)
            base_off_param = get_param_from_list(el, BASE_OFFSET_PARAMS)
            
            if base_lvl_param:
                old_lvl_id = base_lvl_param.AsElementId()
                if not source_level or old_lvl_id == source_level.Id:
                    old_lvl = doc.GetElement(old_lvl_id)
                    if old_lvl and base_off_param:
                        base_absolute_z = old_lvl.Elevation + base_off_param.AsDouble()
                        new_off = base_absolute_z - target_level.Elevation
                        base_lvl_param.Set(target_level.Id)
                        base_off_param.Set(new_off)
                        changed = True
            
            # Top logic
            top_lvl_param = get_param_from_list(el, TOP_LEVEL_PARAMS)
            top_off_param = get_param_from_list(el, TOP_OFFSET_PARAMS)
            if top_lvl_param and top_off_param:
                old_lvl_id = top_lvl_param.AsElementId()
                if not source_level or old_lvl_id == source_level.Id:
                    old_lvl = doc.GetElement(old_lvl_id)
                    if old_lvl:
                        abs_top_z = old_lvl.Elevation + top_off_param.AsDouble()
                        new_top_off = abs_top_z - target_level.Elevation
                        top_lvl_param.Set(target_level.Id)
                        top_off_param.Set(new_top_off)
                        changed = True

            if is_beam and not changed:
                # Basic beam fallback
                trial_bips = [
                    DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                    DB.BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,
                    DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                    DB.BuiltInParameter.HOST_ID_PARAM
                ]
                for bip in trial_bips:
                    try:
                        p = el.get_Parameter(bip)
                        if p and p.Set(target_level.Id):
                            changed = True
                            break
                    except: continue

            if is_beam and not changed:
                # Advanced beam fallback for Work-Plane-Based beams
                p_sp = el.get_Parameter(DB.BuiltInParameter.SKETCH_PLANE_PARAM)
                if p_sp:
                    if not p_sp.IsReadOnly:
                        old_sp_id = p_sp.AsElementId()
                        if old_sp_id != DB.ElementId.InvalidElementId:
                            # We need to maintain the physical Z position.
                            start_off_p = el.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION)
                            end_off_p = el.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION)
                            
                            old_lvl_id = get_level_id_from_list(el, BASE_LEVEL_PARAMS)
                            old_lvl = doc.GetElement(old_lvl_id)
                            
                            if old_lvl and start_off_p and end_off_p:
                                abs_start_z = old_lvl.Elevation + start_off_p.AsDouble()
                                abs_end_z = old_lvl.Elevation + end_off_p.AsDouble()
                                
                                try:
                                    # Find or create the SketchPlane for the target level
                                    target_sp = None
                                    sp_collector = DB.FilteredElementCollector(doc).OfClass(DB.SketchPlane).ToElements()
                                    for sp in sp_collector:
                                        if sp.Name == target_level.Name:
                                            target_sp = sp
                                            break
                                    
                                    if not target_sp:
                                        target_sp = DB.SketchPlane.Create(doc, target_level.Id)
                                        
                                    if target_sp:
                                        # Change the Work Plane
                                        p_sp.Set(target_sp.Id)
                                        doc.Regenerate() # Force Revit to process the plane change
                                        
                                        # Immediately adjust offsets so it doesn't move physically
                                        start_off_p.Set(abs_start_z - target_level.Elevation)
                                        end_off_p.Set(abs_end_z - target_level.Elevation)
                                        changed = True
                                except: pass
                    else:
                        # LOCKED WORK PLANE FALLBACK: Clone and Swap
                        # LOCKED WORK PLANE FALLBACK: True Clone via NewFamilyInstance
                        try:
                            loc = el.Location
                            if isinstance(loc, DB.LocationCurve):
                                curve = loc.Curve
                                symbol = el.Symbol
                                if not symbol.IsActive:
                                    symbol.Activate()
                                
                                # Create a brand new beam directly on the target level (Unlocked!)
                                new_beam = doc.Create.NewFamilyInstance(curve, symbol, target_level, DB.Structure.StructuralType.Beam)
                                
                                # List of parameters to NOT copy (so we don't mess up the new hosting)
                                exclude_bips = [
                                    DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                                    DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                                    DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                                    DB.BuiltInParameter.LEVEL_PARAM,
                                    DB.BuiltInParameter.SKETCH_PLANE_PARAM,
                                    DB.BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION,
                                    DB.BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION,
                                    DB.BuiltInParameter.HOST_ID_PARAM,
                                    DB.BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,
                                    DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM
                                ]
                                exclude_ids = [DB.ElementId(bip) for bip in exclude_bips]
                                
                                # Copy properties
                                for p in el.Parameters:
                                    if p.Id not in exclude_ids and not p.IsReadOnly:
                                        try:
                                            new_p = new_beam.get_Parameter(p.Definition)
                                            if new_p and not new_p.IsReadOnly:
                                                if p.StorageType == DB.StorageType.Double:
                                                    new_p.Set(p.AsDouble())
                                                elif p.StorageType == DB.StorageType.Integer:
                                                    new_p.Set(p.AsInteger())
                                                elif p.StorageType == DB.StorageType.String:
                                                    new_p.Set(p.AsString())
                                                elif p.StorageType == DB.StorageType.ElementId:
                                                    new_p.Set(p.AsElementId())
                                        except: pass
                                
                                # EXPLICITLY enforce the target level and adjust offsets to maintain physical position
                                doc.Regenerate()
                                ref_param = new_beam.get_Parameter(DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                                if ref_param and not ref_param.IsReadOnly:
                                    current_lvl_id = ref_param.AsElementId()
                                    if current_lvl_id != target_level.Id:
                                        start_off = new_beam.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION)
                                        end_off = new_beam.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION)
                                        current_lvl = doc.GetElement(current_lvl_id)
                                        
                                        if current_lvl and start_off and end_off:
                                            abs_start = current_lvl.Elevation + start_off.AsDouble()
                                            abs_end = current_lvl.Elevation + end_off.AsDouble()
                                            
                                            ref_param.Set(target_level.Id)
                                            start_off.Set(abs_start - target_level.Elevation)
                                            end_off.Set(abs_end - target_level.Elevation)
                                
                                doc.Delete(el.Id)
                                changed = True
                        except: pass

            if changed:
                count += 1
                
        t.Commit()
        
        total = len(selected_elements)
        if count == total:
            show_custom_alert("Successfully re-hosted {} element(s).".format(count))
        elif count > 0:
            show_custom_alert("Re-hosted {} of {} element(s).".format(count, total))
        else:
            show_custom_alert("Could not re-host any elements. Beam reference levels may be read-only. Try selecting elements manually and re-host via Revit Properties panel.")
            
    except Exception as e:
        if t.HasStarted() and not t.HasEnded(): 
            t.RollBack()
        traceback.print_exc()
        show_custom_alert("Error occurred:\n{}".format(str(e)))


def get_hosted_elements_for_levels(levels):
    cats = [
        DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Columns, DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_Floors, DB.BuiltInCategory.OST_Roofs, DB.BuiltInCategory.OST_Stairs,
        DB.BuiltInCategory.OST_StructuralFoundation, DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_Windows, DB.BuiltInCategory.OST_Doors, DB.BuiltInCategory.OST_GenericModel,
        DB.BuiltInCategory.OST_PlumbingFixtures, DB.BuiltInCategory.OST_MechanicalEquipment,
        DB.BuiltInCategory.OST_ElectricalFixtures, DB.BuiltInCategory.OST_LightingFixtures
    ]
    multi_filter = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](cats))
    elements = DB.FilteredElementCollector(doc).WherePasses(multi_filter).WhereElementIsNotElementType().ToElements()
    
    level_dict = {lvl.Id: [] for lvl in levels}
    
    for el in elements:
        b_id = get_level_id_from_list(el, BASE_LEVEL_PARAMS)
        t_id = get_level_id_from_list(el, TOP_LEVEL_PARAMS)
        
        if b_id in level_dict:
            level_dict[b_id].append(el)
        if t_id in level_dict and t_id != b_id:
            level_dict[t_id].append(el)
            
    return level_dict

def process_level_modification(level_to_mod, new_elevation_internal, ticked_ids):
    old_elevation = level_to_mod.Elevation
    elev_diff = new_elevation_internal - old_elevation
    if abs(elev_diff) < 0.0001:
        show_custom_alert("Level is already at this elevation.")
        return

    t = DB.Transaction(doc, "Move Level and Maintain Elements")
    t.Start()
    try:
        # Move the level's elevation
        level_to_mod.Elevation = new_elevation_internal
        
        # Adjust all ticked elements
        count = 0
        for el_id in ticked_ids:
            el = doc.GetElement(el_id)
            if not el: continue
            
            is_beam = el.Category and el.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_StructuralFraming)
            
            # Base logic
            base_lvl_param = get_param_from_list(el, BASE_LEVEL_PARAMS)
            base_off_param = get_param_from_list(el, BASE_OFFSET_PARAMS)
            
            updated = False
            if base_lvl_param and base_lvl_param.AsElementId() == level_to_mod.Id:
                if is_beam:
                    end0_param = el.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION)
                    end1_param = el.get_Parameter(DB.BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION)
                    if end0_param and end1_param:
                        end0_param.Set(end0_param.AsDouble() - elev_diff)
                        end1_param.Set(end1_param.AsDouble() - elev_diff)
                        updated = True
                elif base_off_param:
                    old_off = base_off_param.AsDouble()
                    base_off_param.Set(old_off - elev_diff)
                    updated = True
                
            # Top logic
            top_lvl_param = get_param_from_list(el, TOP_LEVEL_PARAMS)
            top_off_param = get_param_from_list(el, TOP_OFFSET_PARAMS)
            if top_lvl_param and top_lvl_param.AsElementId() == level_to_mod.Id and top_off_param:
                old_off = top_off_param.AsDouble()
                top_off_param.Set(old_off - elev_diff)
                updated = True
                
            if updated:
                count += 1
                
        t.Commit()
        show_custom_alert("Successfully moved Level and adjusted offsets for {} element(s).".format(count))
    except Exception as e:
        if t.HasStarted() and not t.HasEnded(): 
            t.RollBack()
        traceback.print_exc()
        show_custom_alert("Error occurred:\n{}".format(str(e)))

# ----------- UI Form logic -------------

class ChangeHostOptionsForm(WPFWindow):
    def __init__(self, xaml_file_name, levels, selected_elements):
        WPFWindow.__init__(self, xaml_file_name)
        
        self.levels = levels
        self.selected_elements = selected_elements
        self.elements_by_level = get_hosted_elements_for_levels(self.levels)
        
        level_names_keep = ["[Keep Current]"] + [lvl.Name for lvl in self.levels]
        level_names_mod = [lvl.Name for lvl in self.levels]
        
        pure_black = Brushes.Black
        white_brush = Brushes.White
        self.Background = pure_black
        
        # Find controls FIRST
        self.source_level_combo = self.FindName("SourceLevelCombo")
        self.rehost_elements_list_panel = self.FindName("RehostElementsListPanel")
        
        self.level_to_replace_combo = self.FindName("LevelToReplaceCombo")
        self.target_level_combo = self.FindName("TargetLevelCombo")
        self.mod_level_combo = self.FindName("ModLevelCombo")
        self.current_elev_text = self.FindName("CurrentElevText")
        self.new_elev_text = self.FindName("NewElevText")
        self.main_tabs = self.FindName("MainTabControl")
        self.elements_list_panel = self.FindName("ElementsListPanel")
        self.logo_img = self.FindName("LogoImage")
        
        # Tab 1 config
        if self.source_level_combo:
            source_level_names = ["[Selected in Revit]"] + [lvl.Name for lvl in self.levels]
            self.source_level_combo.ItemsSource = source_level_names
            self.source_level_combo.Background = pure_black
            self.source_level_combo.Foreground = white_brush
            self.source_level_combo.SelectedIndex = 0
            
        # Find default base and top levels from selected elements
        default_base_id = None
        default_top_id = None
        if self.selected_elements:
            first_el = self.selected_elements[0]
            base_lvl_param = get_param_from_list(first_el, BASE_LEVEL_PARAMS)
            top_lvl_param = get_param_from_list(first_el, TOP_LEVEL_PARAMS)
            if base_lvl_param: default_base_id = base_lvl_param.AsElementId()
            if top_lvl_param: default_top_id = top_lvl_param.AsElementId()

        if self.level_to_replace_combo:
            self.level_to_replace_combo.ItemsSource = ["[Select Level to Replace]"] + [lvl.Name for lvl in self.levels]
            self.level_to_replace_combo.Background = pure_black
            self.level_to_replace_combo.Foreground = white_brush
            self.level_to_replace_combo.SelectedIndex = 0

        if self.target_level_combo:
            self.target_level_combo.ItemsSource = ["[Select Target Level]"] + [lvl.Name for lvl in self.levels]
            self.target_level_combo.Background = pure_black
            self.target_level_combo.Foreground = white_brush
            self.target_level_combo.SelectedIndex = 0
            
        # For Tab 2
        default_level_id = default_base_id

        # Tab 2 config
        if self.mod_level_combo:
            self.mod_level_combo.ItemsSource = level_names_mod
            self.mod_level_combo.Background = pure_black
            self.mod_level_combo.Foreground = white_brush
            
            selected_idx = 0
            if default_level_id and default_level_id != DB.ElementId.InvalidElementId:
                for i, lvl in enumerate(self.levels):
                    if lvl.Id == default_level_id:
                        selected_idx = i
                        break
                        
            if self.levels:
                self.mod_level_combo.SelectedIndex = selected_idx
            
        self.load_logo()
        
        # Auto-detect level from pre-selected elements
        self._auto_set_replace_level(self.selected_elements)
        
    def _auto_set_replace_level(self, elements):
        """Auto-detect and set the Level To Replace combo from the given elements."""
        if not self.level_to_replace_combo or not elements:
            return
        detected_level_id = get_level_id_from_list(elements[0], BASE_LEVEL_PARAMS)
        if detected_level_id and detected_level_id != DB.ElementId.InvalidElementId:
            for i, lvl in enumerate(self.levels):
                if lvl.Id == detected_level_id:
                    self.level_to_replace_combo.SelectedIndex = i + 1  # +1 for placeholder
                    break
        
    def SourceLevelCombo_SelectionChanged(self, sender, args):
        if not self.source_level_combo or not self.rehost_elements_list_panel: return
        idx = self.source_level_combo.SelectedIndex
        self.rehost_elements_list_panel.Children.Clear()
        
        hosted_els = []
        if idx == 0:
            hosted_els = self.selected_elements
            # Auto-detect level from selected elements
            self._auto_set_replace_level(hosted_els)
        elif idx > 0 and idx <= len(self.levels):
            lvl = self.levels[idx - 1]
            hosted_els = self.elements_by_level.get(lvl.Id, [])
            # Set level_to_replace to the source level
            if self.level_to_replace_combo:
                self.level_to_replace_combo.SelectedIndex = idx  # +1 offset matches since source also has placeholder at 0
            
        for el in hosted_els:
            cb = System.Windows.Controls.CheckBox()
            cat_name = el.Category.Name if el.Category else "Unknown"
            cb.Content = "{} [{}]".format(el.Name, cat_name)
            cb.IsChecked = True
            cb.Foreground = Brushes.White
            cb.Margin = System.Windows.Thickness(0, 0, 0, 5)
            cb.Tag = el.Id
            self.rehost_elements_list_panel.Children.Add(cb)

    def ModLevelCombo_SelectionChanged(self, sender, args):
        if not self.mod_level_combo: return
        idx = self.mod_level_combo.SelectedIndex
        if idx >= 0 and idx < len(self.levels):
            lvl = self.levels[idx]
            display_elev = to_display_units(lvl.Elevation, doc)
            disp_str = "{:.3f}".format(display_elev).rstrip('0').rstrip('.') if '.' in "{:.3f}".format(display_elev) else "{:.3f}".format(display_elev)
            if self.current_elev_text:
                self.current_elev_text.Text = disp_str
                
            if self.elements_list_panel:
                self.elements_list_panel.Children.Clear()
                hosted_els = self.elements_by_level.get(lvl.Id, [])
                for el in hosted_els:
                    cb = System.Windows.Controls.CheckBox()
                    cat_name = el.Category.Name if el.Category else "Unknown"
                    cb.Content = "{} [{}]".format(el.Name, cat_name)
                    cb.IsChecked = True
                    cb.Foreground = Brushes.White
                    cb.Margin = System.Windows.Thickness(0, 0, 0, 5)
                    cb.Tag = el.Id
                    self.elements_list_panel.Children.Add(cb)
                
    def PickElementsBtn_Click(self, sender, args):
        self.do_pick = True
        self.Close()

    def RunButton_Click(self, sender, args):
        active_tab = self.main_tabs.SelectedIndex if self.main_tabs else 0
        
        if active_tab == 0:
            # Rehost Elements
            ticked_ids = []
            if self.rehost_elements_list_panel:
                for child in self.rehost_elements_list_panel.Children:
                    if child.IsChecked:
                        ticked_ids.append(child.Tag)
            
            if not ticked_ids:
                show_custom_alert("Please select elements to re-host from the list.")
                return
            
            r_idx = self.level_to_replace_combo.SelectedIndex
            if r_idx <= 0:
                show_custom_alert("Please select the Level To Replace.")
                return

            t_idx = self.target_level_combo.SelectedIndex
            if t_idx <= 0:
                show_custom_alert("Please select a Target Level.")
                return
                
            elements_to_process = [doc.GetElement(eid) for eid in ticked_ids if doc.GetElement(eid)]
            target_level = self.levels[t_idx - 1]
            source_level = self.levels[r_idx - 1]
            
            self.Close()
            process_elements(elements_to_process, target_level, source_level)
            
        elif active_tab == 1:
            # Move Level
            idx = self.mod_level_combo.SelectedIndex
            if idx < 0:
                show_custom_alert("Please select a level to move.")
                return
                
            new_val_str = self.new_elev_text.Text
            if not new_val_str:
                show_custom_alert("Please enter a new elevation.")
                return
                
            try:
                new_display_val = float(new_val_str)
            except ValueError:
                show_custom_alert("Please enter a valid number.")
                return
                
            target_level = self.levels[idx]
            target_elev_int = to_internal_units(new_display_val, doc)
            
            ticked_ids = []
            if self.elements_list_panel:
                for child in self.elements_list_panel.Children:
                    if child.IsChecked:
                        ticked_ids.append(child.Tag)
            
            self.Close()
            process_level_modification(target_level, target_elev_int, ticked_ids)

    def load_logo(self):
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                uri = Uri(logo_path)
                bitmap = BitmapImage(uri)
                if self.logo_img:
                    self.logo_img.Source = bitmap
        except Exception as e:
            pass

    def TitleBar_MouseDown(self, sender, args):
        if args.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
            self.DragMove()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


def main():
    if doc.IsFamilyDocument:
        show_custom_alert("This tool must be run in a Project document.")
        return
        
    levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
    levels = sorted(list(levels_collector), key=lambda x: x.Elevation)
    
    if not levels:
        show_custom_alert("No levels found in this project.")
        return
        
    selected_ids = uidoc.Selection.GetElementIds()
    selected_elements = [doc.GetElement(id) for id in selected_ids]
    
    script_dir = os.path.dirname(__file__)
    xaml_file = os.path.join(script_dir, "ui.xaml")
    
    while True:
        form = ChangeHostOptionsForm(xaml_file, levels, selected_elements)
        form.ShowDialog()
        
        if getattr(form, "do_pick", False):
            try:
                import System.Diagnostics
                from System.Runtime.InteropServices import DllImport
                class User32:
                    @staticmethod
                    @DllImport("user32.dll")
                    def SetForegroundWindow(hWnd):
                        pass
                proc = System.Diagnostics.Process.GetCurrentProcess()
                User32.SetForegroundWindow(proc.MainWindowHandle)
            except:
                pass
                
            try:
                sel = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, "Select elements to re-host. Press Enter or click Finish when done.")
                if sel:
                    selected_elements = [doc.GetElement(r) for r in sel]
            except Exception:
                pass
            continue
            
        break

if __name__ == '__main__':
    main()
