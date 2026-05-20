# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

from Autodesk.Revit import DB, UI
from pyrevit import revit, forms, script
import os.path as op
import math

# Initialize
doc = revit.doc
uidoc = revit.uidoc

class BrickMasonryWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.params = None
        # Set initial values for first selection
        self.update_dimensions(0)
        self.ShowDialog()

    def update_dimensions(self, index):
        # 0: Standard Brick (215x102.5x65)
        # 1: Solid Concrete Block (440x100x215)
        # 2: Hollow Concrete Block (440x215x215)
        # 3: Paving Block (200x100x60)
        if index == 0:
            self.txt_length.Text = "215"
            self.txt_width.Text = "102.5"
            self.txt_height.Text = "65"
        elif index == 1:
            self.txt_length.Text = "440"
            self.txt_width.Text = "100"
            self.txt_height.Text = "215"
        elif index == 2:
            self.txt_length.Text = "440"
            self.txt_width.Text = "215"
            self.txt_height.Text = "215"
        elif index == 3:
            self.txt_length.Text = "200"
            self.txt_width.Text = "100"
            self.txt_height.Text = "60"

    def cb_block_type_SelectionChanged(self, sender, e):
        if hasattr(self, 'txt_length'): # Ensure UI is loaded
            self.update_dimensions(sender.SelectedIndex)

    def get_inputs(self):
        try:
            return {
                'b_l': float(self.txt_length.Text) / 304.8,
                'b_w': float(self.txt_width.Text) / 304.8,
                'b_h': float(self.txt_height.Text) / 304.8,
                'm_gap': float(self.txt_mortar.Text) / 304.8,
                'pattern_idx': self.cb_pattern.SelectedIndex,
                'block_idx': self.cb_block_type.SelectedIndex,
                'hide_wall': self.chk_hide_wall.IsChecked
            }
        except:
            forms.alert("Please enter valid numeric values.")
            return None

    def btn_convert_walls_Click(self, sender, e):
        self.params = self.get_inputs()
        if not self.params: return
        self.Close()

        # Pick Walls
        try:
            with forms.WarningBar(title="Select walls to convert into blocks"):
                references = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, "Select walls")
            
            if references:
                walls = [doc.GetElement(ref) for ref in references if isinstance(doc.GetElement(ref), DB.Wall)]
                if walls:
                    with revit.Transaction("Convert Walls to Bricks"):
                        for wall in walls:
                            convert_wall_to_bricks(wall, self.params)
                else:
                    forms.alert("No valid walls selected.")
        except:
            pass

def convert_wall_to_bricks(wall, p):
    curve = wall.Location.Curve
    height_param = wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    wall_height = height_param.AsDouble() if height_param else 10.0
    
    create_bricks(curve, wall_height, p)
    
    if p['hide_wall']:
        from System.Collections.Generic import List
        ids = List[DB.ElementId]()
        ids.Add(wall.Id)
        doc.ActiveView.HideElements(ids)

def create_bricks(curve, wall_height, p):
    b_l, b_w, b_h = p['b_l'], p['b_w'], p['b_h']
    m_gap = p['m_gap']
    pattern_idx = p['pattern_idx'] # 0: Running, 1: Stack, 2: English, 3: Flemish, 4: Common

    total_len = curve.Length
    num_courses = int(math.ceil(wall_height / (b_h + m_gap)))
    
    for c in range(num_courses):
        # Determine Course Type
        is_header_course = False
        if pattern_idx == 2: # English: alt stretcher and header courses
            is_header_course = (c % 2 != 0)
        elif pattern_idx == 4: # Common: header every 6th course
            is_header_course = ((c + 1) % 6 == 0)

        # Running bond offset for stretchers
        offset = 0
        if pattern_idx == 0 and c % 2 != 0:
            offset = (b_l + m_gap) / 2.0
        
        dist = -offset
        while dist < total_len:
            # Determine current block length and width for this specific spot
            curr_l = b_l
            if is_header_course:
                curr_l = b_w # Headers use width as length in single skin
            elif pattern_idx == 3: # Flemish: Alt stretcher and header in each row
                # We need a more complex tracking for Flemish in a while loop
                # I'll handle Flemish with a separate loop or internal toggle
                pass

            # Handle Flemish toggle inside the loop
            if pattern_idx == 3:
                # Calculate which block we are on based on dist
                # This is tricky with variable lengths.
                # Let's use a specialized logic for Flemish
                break 

            # Standard placement for other patterns
            place_block(curve, dist, curr_l, b_w, b_h, m_gap, c, total_len, p.get('block_idx', 0))
            dist += curr_l + m_gap
        
        # Specialized loop for Flemish Bond
        if pattern_idx == 3:
            dist = 0
            # Offset Flemish rows? Yes, usually headers align over stretcher centers
            if c % 2 != 0:
                dist = -(b_l + m_gap) / 2.0 # Start with a half offset or similar

            is_stretcher = True
            while dist < total_len:
                curr_l = b_l if is_stretcher else b_w
                place_block(curve, dist, curr_l, b_w, b_h, m_gap, c, total_len, p.get('block_idx', 0))
                dist += curr_l + m_gap
                is_stretcher = not is_stretcher

def place_block(curve, dist, b_l, b_w, b_h, m_gap, course_idx, total_len, block_idx):
    # Clip block at ends
    start_dist = max(0, dist)
    end_dist = min(total_len, dist + b_l)
    
    if end_dist - start_dist > 0.005:
        actual_l = end_dist - start_dist
        mid_dist = start_dist + (actual_l / 2.0)
        
        normalized_dist = max(0, min(1, mid_dist / total_len))
        transform_at_point = curve.ComputeDerivatives(normalized_dist, True)
        
        origin = transform_at_point.Origin
        tangent = transform_at_point.BasisX.Normalize()
        
        pos = origin + DB.XYZ(0, 0, course_idx * (b_h + m_gap))
        width_dir = DB.XYZ.BasisZ.CrossProduct(tangent).Normalize()
        
        t = DB.Transform.Identity
        t.Origin = pos
        t.BasisX = tangent
        t.BasisY = width_dir
        t.BasisZ = DB.XYZ.BasisZ
        
        # Outer Loop
        p0 = DB.XYZ(-actual_l/2.0, -b_w/2.0, 0)
        p1 = DB.XYZ(actual_l/2.0, -b_w/2.0, 0)
        p2 = DB.XYZ(actual_l/2.0, b_w/2.0, 0)
        p3 = DB.XYZ(-actual_l/2.0, b_w/2.0, 0)
        
        loops = []
        outer_pts = [t.OfPoint(pt) for pt in [p0, p1, p2, p3]]
        outer_loop = DB.CurveLoop()
        try:
            for i in range(4):
                if outer_pts[i].DistanceTo(outer_pts[(i+1)%4]) > 0.0026:
                    outer_loop.Append(DB.Line.CreateBound(outer_pts[i], outer_pts[(i+1)%4]))
            loops.append(outer_loop)
        except:
            return # Skip this block if it's too small

        # Hollow Logic (2 Holes)
        # Only if it's the hollow block type and not heavily clipped
        if block_idx == 2 and actual_l > b_l * 0.8:
            shell = 25 / 304.8 # 25mm shell
            web = 30 / 304.8   # 30mm middle web
            hole_l = (b_l - (2 * shell) - web) / 2.0
            hole_w = b_w - (2 * shell)

            if hole_l > 0 and hole_w > 0:
                # Hole 1 (Left)
                h1_min_x = -b_l/2.0 + shell
                h1_max_x = h1_min_x + hole_l
                h1_min_y = -b_w/2.0 + shell
                h1_max_y = h1_min_y + hole_w

                hp0 = DB.XYZ(h1_min_x, h1_min_y, 0)
                hp1 = DB.XYZ(h1_max_x, h1_min_y, 0)
                hp2 = DB.XYZ(h1_max_x, h1_max_y, 0)
                hp3 = DB.XYZ(h1_min_x, h1_max_y, 0)
                
                h1_pts = [t.OfPoint(pt) for pt in [hp0, hp3, hp2, hp1]] # Reverse for void
                h1_loop = DB.CurveLoop()
                valid_h1 = True
                for i in range(4):
                    if h1_pts[i].DistanceTo(h1_pts[(i+1)%4]) > 0.0026:
                        h1_loop.Append(DB.Line.CreateBound(h1_pts[i], h1_pts[(i+1)%4]))
                    else:
                        valid_h1 = False
                if valid_h1: loops.append(h1_loop)

                # Hole 2 (Right)
                h2_max_x = b_l/2.0 - shell
                h2_min_x = h2_max_x - hole_l
                
                hp4 = DB.XYZ(h2_min_x, h1_min_y, 0)
                hp5 = DB.XYZ(h2_max_x, h1_min_y, 0)
                hp6 = DB.XYZ(h2_max_x, h1_max_y, 0)
                hp7 = DB.XYZ(h2_min_x, h1_max_y, 0)

                h2_pts = [t.OfPoint(pt) for pt in [hp4, hp7, hp6, hp5]] # Reverse for void
                h2_loop = DB.CurveLoop()
                valid_h2 = True
                for i in range(4):
                    if h2_pts[i].DistanceTo(h2_pts[(i+1)%4]) > 0.0026:
                        h2_loop.Append(DB.Line.CreateBound(h2_pts[i], h2_pts[(i+1)%4]))
                    else:
                        valid_h2 = False
                if valid_h2: loops.append(h2_loop)

        try:
            from System.Collections.Generic import List
            loop_list = List[DB.CurveLoop]()
            for l in loops: loop_list.Add(l)
            
            solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(loop_list, DB.XYZ.BasisZ, b_h)
            ds = DB.DirectShape.CreateElement(doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
            ds.SetShape([solid])
            ds.Name = "Hollow Block" if block_idx == 2 else "Brick Block"
        except:
            pass

# Run
xaml_file = op.join(op.dirname(__file__), 'ui.xaml')
BrickMasonryWindow(xaml_file)
