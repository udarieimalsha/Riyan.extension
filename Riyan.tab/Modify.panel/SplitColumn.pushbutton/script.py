# -*- coding: utf-8 -*-
__title__  = "Split Column"
__author__ = "Chalana Perera"
__doc__    = """Bundle: SplitColumn (pushbutton)
Splits selected structural columns at chosen levels."""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import os
import math
from Autodesk.Revit.DB import (
    FilteredElementCollector, Level, FamilyInstance, BuiltInCategory,
    BuiltInParameter, ElementTransformUtils, XYZ, Line
)
from Autodesk.Revit.DB import LocationPoint
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

clr.AddReference('System')
clr.AddReference('PresentationCore')
import System
from System.Windows.Media.Imaging import BitmapImage

from pyrevit import revit, forms
import sys

doc  = revit.doc
uidoc = revit.uidoc

# ─────────────────────────────────────────────
# Filter
# ─────────────────────────────────────────────
class ColumnSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return (isinstance(elem, FamilyInstance)
                and elem.Category is not None
                and elem.Category.Id.IntegerValue
                    == int(BuiltInCategory.OST_StructuralColumns))
    def AllowReference(self, ref, pos):
        return False

# ─────────────────────────────────────────────
# Collect columns
# ─────────────────────────────────────────────
sel_ids = uidoc.Selection.GetElementIds()
columns = []

if sel_ids:
    for eid in sel_ids:
        e = doc.GetElement(eid)
        if (isinstance(e, FamilyInstance)
                and e.Category is not None
                and e.Category.Id.IntegerValue
                    == int(BuiltInCategory.OST_StructuralColumns)):
            columns.append(e)
else:
    from Autodesk.Revit.Exceptions import OperationCanceledException
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, ColumnSelectionFilter(),
            "Select Structural Columns to Split")
        for r in refs:
            columns.append(doc.GetElement(r.ElementId))
    except OperationCanceledException:
        sys.exit()

if not columns:
    forms.alert("No structural columns selected.", exitscript=True)

# ─────────────────────────────────────────────
# Find candidate split levels
# ─────────────────────────────────────────────
all_levels = sorted(
    FilteredElementCollector(doc).OfClass(Level).ToElements(),
    key=lambda l: l.Elevation)

candidates = set()

for col in columns:
    bp = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
    tp = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
    if not bp or not tp:
        continue
        
    base_lvl = doc.GetElement(bp.AsElementId())
    top_lvl  = doc.GetElement(tp.AsElementId())
    if not base_lvl or not top_lvl:
        continue

    bop = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
    top_op = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
    orig_base_off = bop.AsDouble() if bop else 0.0
    orig_top_off  = top_op.AsDouble() if top_op else 0.0

    base_elev = base_lvl.Elevation + orig_base_off
    top_elev  = top_lvl.Elevation  + orig_top_off

    # Include all levels physically touched/crossed by this column
    for l in all_levels:
        if base_elev - 0.001 <= l.Elevation <= top_elev + 0.001:
            candidates.add(l)

if not candidates:
    forms.alert("No levels found intersecting the selected columns.", exitscript=True)

sorted_candidates = sorted(list(candidates), key=lambda l: l.Elevation)

# ─────────────────────────────────────────────
# Custom WPF UI
# ─────────────────────────────────────────────
class LevelItem(object):
    def __init__(self, level):
        self.level = level
        self._name = level.Name
        self._is_checked = True
        self._offset = "0"
        
    @property
    def Name(self):
        return self._name
        
    @property
    def IsChecked(self):
        return self._is_checked
        
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value

    @property
    def Offset(self):
        return self._offset
        
    @Offset.setter
    def Offset(self, value):
        self._offset = value

class SplitUI(forms.WPFWindow):
    def __init__(self, xaml_file_name, cand_levels):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        import System.Windows.Media.Imaging as wmi
        import System.IO as sysio
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.png')
        if os.path.exists(logo_path):
            try:
                stream = sysio.FileStream(logo_path, sysio.FileMode.Open, sysio.FileAccess.Read)
                bitmap = wmi.BitmapImage()
                bitmap.BeginInit()
                bitmap.StreamSource = stream
                bitmap.CacheOption = wmi.BitmapCacheOption.OnLoad
                bitmap.EndInit()
                stream.Close()
                self.logo_img.Source = bitmap
            except:
                pass
            
        self.items = [LevelItem(l) for l in cand_levels]
        self.level_lb.ItemsSource = self.items
        self.selected_points = []

    def split_clicked(self, sender, e):
        for item in self.items:
            if item.IsChecked:
                try:
                    off_ft = float(item.Offset) / 304.8
                except ValueError:
                    off_ft = 0.0
                self.selected_points.append((item.level, off_ft))
        self.Close()

    def cancel_clicked(self, sender, e):
        self.selected_points = []
        self.Close()

xaml_path = os.path.join(os.path.dirname(__file__), 'ui.xaml')
ui = SplitUI(xaml_path, sorted_candidates)
ui.ShowDialog()

split_points_input = ui.selected_points
if not split_points_input:
    sys.exit()

# ─────────────────────────────────────────────
# Safe Parameter Setter
# ─────────────────────────────────────────────
def safe_set_levels(new_col, sb_id, st_id, sbo, sto):
    p_base = new_col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
    p_top  = new_col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
    p_base_off = new_col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
    p_top_off  = new_col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)

    # 1. Throw Top Offset way up to safely bypass Top <= Base constraints
    if p_top_off and not p_top_off.IsReadOnly:
        p_top_off.Set(50000.0) # 50 meters

    # 2. Set Levels
    if p_top  and not p_top.IsReadOnly: p_top.Set(st_id)
    if p_base and not p_base.IsReadOnly: p_base.Set(sb_id)

    # 3. Regenerate (CRITICAL for Steel Columns with Analytical Models)
    doc.Regenerate()

    # 4. Apply correct physical offsets
    if p_base_off and not p_base_off.IsReadOnly: p_base_off.Set(sbo)
    if p_top_off  and not p_top_off.IsReadOnly:  p_top_off.Set(sto)

# ─────────────────────────────────────────────
# Main split loop
# ─────────────────────────────────────────────
with revit.Transaction("Split Columns by Levels"):
    for col in columns:
        bp  = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
        tp  = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
        if not bp or not tp:
            continue

        base_lvl = doc.GetElement(bp.AsElementId())
        top_lvl  = doc.GetElement(tp.AsElementId())
        if not base_lvl or not top_lvl:
            continue

        bop    = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        top_op = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
        orig_base_off = bop.AsDouble()    if bop    else 0.0
        orig_top_off  = top_op.AsDouble() if top_op else 0.0

        base_elev = base_lvl.Elevation + orig_base_off
        top_elev  = top_lvl.Elevation  + orig_top_off

        # Valid levels are those strictly inside or equal to the column's physical bounds
        split_points = []
        for l, off in split_points_input:
            split_z = l.Elevation + off
            if base_elev - 0.001 <= split_z <= top_elev + 0.001:
                split_points.append((l, off, split_z))
                
        # Sort split points by exact absolute elevation
        split_points.sort(key=lambda x: x[2])
        
        if not split_points:
            continue

        # Build segments
        segs = []
        segs.append((base_lvl,  split_points[0][0],  orig_base_off, split_points[0][1]))
        for i in range(len(split_points) - 1):
            lvl1, off1, z1 = split_points[i]
            lvl2, off2, z2 = split_points[i + 1]
            segs.append((lvl1, lvl2, off1, off2))
        segs.append((split_points[-1][0], top_lvl, split_points[-1][1], orig_top_off))

        # Filter out zero-length segments
        final_segs = []
        for sb, st, sbo, sto in segs:
            z_base = sb.Elevation + sbo
            z_top = st.Elevation + sto
            if z_top - z_base > 0.001:
                final_segs.append((sb, st, sbo, sto))

        if not final_segs:
            continue

        # Get original physical info
        col_loc = col.Location
        if not isinstance(col_loc, LocationPoint):
            continue

        origin = col_loc.Point
        symbol = col.Symbol
        if not symbol.IsActive:
            symbol.Activate()

        try:
            xf = col.GetTransform()
            col_rotation = math.atan2(xf.BasisX.Y, xf.BasisX.X)
        except Exception:
            col_rotation = 0.0

        usage_param = col.get_Parameter(BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM)
        orig_usage = usage_param.AsInteger() if usage_param else None

        # Create segments
        for sb, st, sbo, sto in final_segs:
            pt = XYZ(origin.X, origin.Y, sb.Elevation)
            new_col = doc.Create.NewFamilyInstance(pt, symbol, sb, StructuralType.Column)

            if abs(col_rotation) > 0.001:
                axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
                ElementTransformUtils.RotateElement(doc, new_col.Id, axis, col_rotation)

            if orig_usage is not None:
                new_usage = new_col.get_Parameter(BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM)
                if new_usage and not new_usage.IsReadOnly:
                    new_usage.Set(orig_usage)

            safe_set_levels(new_col, sb.Id, st.Id, sbo, sto)

        doc.Delete(col.Id)
