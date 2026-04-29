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
        self.ShowDialog()

    def get_inputs(self):
        try:
            return {
                'b_l': float(self.txt_length.Text) / 304.8,
                'b_w': float(self.txt_width.Text) / 304.8,
                'b_h': float(self.txt_height.Text) / 304.8,
                'm_gap': float(self.txt_mortar.Text) / 304.8,
                'pattern_idx': self.cb_pattern.SelectedIndex,
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
    # Extract Wall Data
    curve = wall.Location.Curve
    # Get Unconnected Height
    height_param = wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    wall_height = height_param.AsDouble() if height_param else 10.0 # Default fallback
    
    # Generate Bricks
    create_bricks(curve, wall_height, p)
    
    # Hide Original Wall
    if p['hide_wall']:
        from System.Collections.Generic import List
        ids = List[DB.ElementId]()
        ids.Add(wall.Id)
        doc.ActiveView.HideElements(ids)

def create_bricks(curve, wall_height, p):
    b_l, b_w, b_h = p['b_l'], p['b_w'], p['b_h']
    m_gap = p['m_gap']
    pattern_idx = p['pattern_idx']

    total_len = curve.Length
    num_courses = int(math.ceil(wall_height / (b_h + m_gap)))
    
    for c in range(num_courses):
        # Pattern offset (Stretcher bond)
        offset = 0
        if pattern_idx == 0 and c % 2 != 0:
            offset = (b_l + m_gap) / 2.0
        
        dist = -offset
        while dist < total_len:
            # Clip block at ends
            start_dist = max(0, dist)
            end_dist = min(total_len, dist + b_l)
            
            if end_dist - start_dist > 0.001:
                actual_l = end_dist - start_dist
                mid_dist = start_dist + (actual_l / 2.0)
                
                # Evaluate point and tangent at mid_dist
                normalized_dist = mid_dist / total_len
                # Ensure it's within [0, 1]
                normalized_dist = max(0, min(1, normalized_dist))
                
                # Get point and derivatives
                transform_at_point = curve.ComputeDerivatives(normalized_dist, True)
                origin = transform_at_point.Origin
                tangent = transform_at_point.BasisX.Normalize()
                
                # Z Position
                pos = origin + DB.XYZ(0, 0, c * (b_h + m_gap))
                
                # Width direction
                width_dir = DB.XYZ.BasisZ.CrossProduct(tangent).Normalize()
                
                # Create Local Transform
                t = DB.Transform.Identity
                t.Origin = pos
                t.BasisX = tangent
                t.BasisY = width_dir
                t.BasisZ = DB.XYZ.BasisZ
                
                # Geometry
                p0 = DB.XYZ(-actual_l/2.0, -b_w/2.0, 0)
                p1 = DB.XYZ(actual_l/2.0, -b_w/2.0, 0)
                p2 = DB.XYZ(actual_l/2.0, b_w/2.0, 0)
                p3 = DB.XYZ(-actual_l/2.0, b_w/2.0, 0)
                
                pts = [t.OfPoint(pt) for pt in [p0, p1, p2, p3]]
                curve_loop = DB.CurveLoop()
                for i in range(4):
                    curve_loop.Append(DB.Line.CreateBound(pts[i], pts[(i+1)%4]))
                
                try:
                    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry([curve_loop], DB.XYZ.BasisZ, b_h)
                    ds = DB.DirectShape.CreateElement(doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
                    ds.SetShape([solid])
                    ds.Name = "Brick Block"
                except:
                    pass
            
            dist += b_l + m_gap

# Run
xaml_file = op.join(op.dirname(__file__), 'ui.xaml')
BrickMasonryWindow(xaml_file)
