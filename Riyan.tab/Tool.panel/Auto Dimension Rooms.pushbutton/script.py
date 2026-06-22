# -*- coding: utf-8 -*-
import clr
import math
import os

from pyrevit import revit, DB, UI, forms
from pyrevit.framework import List

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView

# ------------------------------------------------------------------------------
# UI Forms
# ------------------------------------------------------------------------------

class AutoDimensionForm(forms.WPFWindow):
    def __init__(self, xaml_file_name, **kwargs):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        self.dimension_types_dict = kwargs.get('dimension_types_dict', {})
        
        # Populate Comboboxes
        self.DimStyleCombo.ItemsSource = self.dimension_types_dict.keys()
        if self.dimension_types_dict:
            self.DimStyleCombo.SelectedIndex = 0
            
    def TitleBar_MouseDown(self, sender, e):
        if e.ChangedButton == UI.MouseButtons.Left:
            self.DragMove()
            
    def CloseBtn_Click(self, sender, e):
        self.DialogResult = False
        self.Close()
        
    def RunBtn_Click(self, sender, e):
        # Target
        idx = self.TargetCombo.SelectedIndex
        if idx == 0: self.target_choice = "Wall Faces"
        elif idx == 1: self.target_choice = "Wall Cores"
        else: self.target_choice = "Grids"
        
        # Dim Style
        style_name = self.DimStyleCombo.SelectedItem
        self.dim_type = self.dimension_types_dict.get(style_name)
            
        self.DialogResult = True
        self.Close()

# ------------------------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------------------------

def get_dimension_types():
    return DB.FilteredElementCollector(doc)\
             .OfClass(DB.DimensionType)\
             .ToElements()

def get_rooms_in_view(v):
    return DB.FilteredElementCollector(doc, v.Id)\
             .OfCategory(DB.BuiltInCategory.OST_Rooms)\
             .WhereElementIsNotElementType()\
             .ToElements()

def get_grids_in_view(v):
    return DB.FilteredElementCollector(doc, v.Id)\
             .OfClass(DB.Grid)\
             .ToElements()

def get_view_content_extents(v):
    """Get the bounding box of ALL visible elements in the view,
    including annotations, grids, section marks, etc."""
    vmin_x = vmin_y = float('inf')
    vmax_x = vmax_y = float('-inf')
    
    all_elements = DB.FilteredElementCollector(doc, v.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElements()
    
    for elem in all_elements:
        try:
            bb = elem.get_BoundingBox(v)
            if bb:
                if bb.Min.X < vmin_x: vmin_x = bb.Min.X
                if bb.Min.Y < vmin_y: vmin_y = bb.Min.Y
                if bb.Max.X > vmax_x: vmax_x = bb.Max.X
                if bb.Max.Y > vmax_y: vmax_y = bb.Max.Y
        except Exception:
            pass
    
    return vmin_x, vmin_y, vmax_x, vmax_y

def get_wall_references(wall, target_type, room_center=None):
    """Get references for a wall. For Wall Faces, returns only the face
    closest to the room center (interior face) to avoid wall-thickness dims."""
    refs = []
    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    opt.View = view
    
    geom_elem = wall.get_Geometry(opt)
    if not geom_elem:
        return refs
        
    if target_type == "Wall Faces":
        try:
            ext_refs = DB.HostObjectUtils.GetSideFaces(wall, DB.ShellLayerType.Exterior)
            int_refs = DB.HostObjectUtils.GetSideFaces(wall, DB.ShellLayerType.Interior)
            
            best_ref = None
            min_dist = float('inf')
            
            for r in list(ext_refs) + list(int_refs):
                try:
                    face = wall.GetGeometryObjectFromReference(r)
                    if face and room_center:
                        bbox = face.GetBoundingBox()
                        uv_center = (bbox.Min + bbox.Max) / 2.0
                        pt = face.Evaluate(uv_center)
                        dist = pt.DistanceTo(room_center)
                        if dist < min_dist:
                            min_dist = dist
                            best_ref = r
                except Exception:
                    pass
            if best_ref:
                refs.append(best_ref)
            elif int_refs:
                refs.append(int_refs[0])
        except Exception:
            pass
    elif target_type == "Wall Cores":
        loc = wall.Location
        if isinstance(loc, DB.LocationCurve) and loc.Curve.Reference:
            refs.append(loc.Curve.Reference)
    return refs

# ------------------------------------------------------------------------------
# Main Logic
# ------------------------------------------------------------------------------

def main():
    if view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan]:
        forms.alert("Please run this tool in a Plan view.")
        return
        
    dim_types = get_dimension_types()
    if not dim_types:
        forms.alert("No Dimension Types found.")
        return
        
    dim_type_dict = {}
    for dt in dim_types:
        try:
            dt_name = dt.Name
        except Exception:
            try:
                dt_name = dt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            except Exception:
                dt_name = "DimType " + str(dt.Id)
        if dt_name:
            dim_type_dict[dt_name] = dt
            
    # Show Custom UI
    xaml_path = os.path.join(os.path.dirname(__file__), "AutoDimension.xaml")
    form = AutoDimensionForm(xaml_path, dimension_types_dict=dim_type_dict)
    
    if not form.ShowDialog():
        return
        
    target_choice = form.target_choice
    dim_type = form.dim_type

    rooms = get_rooms_in_view(view)
    if not rooms:
        forms.alert("No Rooms found in the active view.")
        return

    grids = get_grids_in_view(view)
    
    # Start Transaction
    with revit.Transaction("Auto Dimension Rooms"):
        # First, find the overall bounding box of all rooms
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')
        
        valid_rooms = []
        for room in rooms:
            if room.Area > 0:
                bbox = room.get_BoundingBox(view)
                if bbox:
                    valid_rooms.append((room, bbox))
                    if bbox.Min.X < min_x: min_x = bbox.Min.X
                    if bbox.Min.Y < min_y: min_y = bbox.Min.Y
                    if bbox.Max.X > max_x: max_x = bbox.Max.X
                    if bbox.Max.Y > max_y: max_y = bbox.Max.Y

        if not valid_rooms:
            forms.alert("No valid room geometry found.")
            return

        # Calculate building center
        bldg_cx = (min_x + max_x) / 2.0
        bldg_cy = (min_y + max_y) / 2.0
        
        global_refs_top = {}
        global_refs_bottom = {}
        global_refs_left = {}
        global_refs_right = {}
        
        for room, bbox in valid_rooms:
            center = (bbox.Min + bbox.Max) / 2.0
            try:
                room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString() or "Room"
            except Exception:
                room_name = "Room"
            
            if target_choice == "Grids":
                for g in grids:
                    gc = g.Curve
                    if gc:
                        direction = gc.Direction
                        try:
                            mid_pt = gc.Evaluate(0.5, True)
                        except Exception:
                            continue
                        # Only include grids that pass through or near this room
                        if abs(direction.X) > 0.9:
                            if bbox.Min.Y - 1 < mid_pt.Y < bbox.Max.Y + 1:
                                key = round(mid_pt.Y / TOL)
                                if center.X >= bldg_cx:
                                    if key not in global_refs_right: global_refs_right[key] = (DB.Reference(g), room_name)
                                else:
                                    if key not in global_refs_left: global_refs_left[key] = (DB.Reference(g), room_name)
                        elif abs(direction.Y) > 0.9:
                            if bbox.Min.X - 1 < mid_pt.X < bbox.Max.X + 1:
                                key = round(mid_pt.X / TOL)
                                if center.Y >= bldg_cy:
                                    if key not in global_refs_top: global_refs_top[key] = (DB.Reference(g), room_name)
                                else:
                                    if key not in global_refs_bottom: global_refs_bottom[key] = (DB.Reference(g), room_name)
            else:
                options = DB.SpatialElementBoundaryOptions()
                segments = room.GetBoundarySegments(options)
                if segments:
                    for boundary_list in segments:
                        for segment in boundary_list:
                            element = doc.GetElement(segment.ElementId)
                            if isinstance(element, DB.Wall):
                                w_refs = get_wall_references(element, target_choice, center)
                                wall_dir = segment.GetCurve().Direction
                                seg_curve = segment.GetCurve()
                                try:
                                    mid_pt = seg_curve.Evaluate(0.5, True)
                                except Exception:
                                    continue
                                for r in w_refs:
                                    if abs(wall_dir.X) > 0.9:
                                        key = round(mid_pt.Y / TOL)
                                        if center.X >= bldg_cx:
                                            if key not in global_refs_right: global_refs_right[key] = (r, room_name)
                                        else:
                                            if key not in global_refs_left: global_refs_left[key] = (r, room_name)
                                    elif abs(wall_dir.Y) > 0.9:
                                        key = round(mid_pt.X / TOL)
                                        if center.Y >= bldg_cy:
                                            if key not in global_refs_top: global_refs_top[key] = (r, room_name)
                                        else:
                                            if key not in global_refs_bottom: global_refs_bottom[key] = (r, room_name)

        # Sort all references
        sorted_top = sorted(global_refs_top.items())
        sorted_bottom = sorted(global_refs_bottom.items())
        sorted_left = sorted(global_refs_left.items())
        sorted_right = sorted(global_refs_right.items())
        
        def create_dim_from_sorted(sorted_refs, is_horizontal, is_positive_side):
            if len(sorted_refs) < 2: return None
            ref_array = DB.ReferenceArray()
            for _, (r, _) in sorted_refs:
                ref_array.Append(r)
            
            # Position the line
            offset_dist = 4.0 # feet
            if is_horizontal:
                y_pos = (max_y + offset_dist) if is_positive_side else (min_y - offset_dist)
                pt1 = DB.XYZ(min_x, y_pos, 0)
                pt2 = DB.XYZ(max_x, y_pos, 0)
            else:
                x_pos = (max_x + offset_dist) if is_positive_side else (min_x - offset_dist)
                pt1 = DB.XYZ(x_pos, min_y, 0)
                pt2 = DB.XYZ(x_pos, max_y, 0)
                
            line = DB.Line.CreateBound(pt1, pt2)
            try:
                dim = doc.Create.NewDimension(view, line, ref_array, dim_type)
                return (dim, sorted_refs)
            except Exception:
                return None
        
        created_dims = []
        
        dim_top = create_dim_from_sorted(sorted_top, True, True)
        if dim_top: created_dims.append(dim_top)
        
        dim_bottom = create_dim_from_sorted(sorted_bottom, True, False)
        if dim_bottom: created_dims.append(dim_bottom)
        
        dim_right = create_dim_from_sorted(sorted_right, False, True)
        if dim_right: created_dims.append(dim_right)
        
        dim_left = create_dim_from_sorted(sorted_left, False, False)
        if dim_left: created_dims.append(dim_left)
        
        # Set Below Text to Room Name
        for dim, sorted_refs in created_dims:
            if dim.Segments.Size > 0:
                seg_list = list(dim.Segments)
                for i, segment in enumerate(seg_list):
                    try:
                        val = segment.Value
                        if val is not None and val > MIN_DIST:
                            if i < len(sorted_refs) - 1:
                                segment.Below = sorted_refs[i][1][1]
                    except Exception:
                        pass
            else:
                try:
                    if dim.Value and dim.Value > MIN_DIST and sorted_refs:
                        dim.Below = sorted_refs[0][1][1]
                except Exception:
                    pass

if __name__ == '__main__':
    main()
