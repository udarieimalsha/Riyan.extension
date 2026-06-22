# -*- coding: utf-8 -*-
"""Searches and picks elements in current view based on input tag defined in instance parameter. (Riyan)"""
__doc__ = "Searches and picks elements in current view based on input tag defined in instance parameter."
__title__ = "Instance Tag\nFinder"

import os
import clr
clr.AddReference("System.Windows.Presentation")
import System

from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit.forms import WPFWindow
from pyrevit.framework import List
from collections import defaultdict


class TagSearchWindow(WPFWindow):
    """Riyan-themed WPF window for instance tag input."""
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.result = None

    def TitleBar_MouseDown(self, sender, args):
        if args.LeftButton == System.Windows.Input.MouseButtonState.Pressed:
            self.DragMove()

    def SearchButton_Click(self, sender, args):
        val = self.TagInput.Text.strip()
        if val:
            self.result = val
            self.DialogResult = True
            self.Close()
        else:
            forms.alert("Please enter a tag value to search for.", title="Invalid Input")

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


if __name__ == '__main__':
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    curview = doc.ActiveView

    if isinstance(curview, DB.ViewSheet):
        forms.alert("You're on a Sheet. Activate a model view please.", exitscript=True)

    # Show Riyan-themed input window
    xaml_file = os.path.join(os.path.dirname(__file__), "search_input.xaml")
    target_tag = None

    if os.path.exists(xaml_file):
        try:
            window = TagSearchWindow(xaml_file)
            window.ShowDialog()
            if window.DialogResult and window.result:
                target_tag = window.result.lower()
        except Exception as e:
            forms.alert("Could not load input window:\n{}".format(e))
            target_tag = str(forms.ask_for_string("Enter tag name")).lower()
    else:
        target_tag = str(forms.ask_for_string("Enter tag name")).lower()

    if not target_tag:
        import sys; sys.exit()

    try:
        param_dic = defaultdict(list)
        family_types_elements = defaultdict(list)
        wall_id = []
        wall_id_list = None

        options_category = {
            'Structural Columns': DB.BuiltInCategory.OST_StructuralColumns,
            'Walls': DB.BuiltInCategory.OST_Walls,
            'Structural Framing': DB.BuiltInCategory.OST_StructuralFraming,
            'Floors': DB.BuiltInCategory.OST_Floors,
            'Foundation': DB.BuiltInCategory.OST_StructuralFoundation
        }

        selected_switch_category = forms.CommandSwitchWindow.show(
            sorted(options_category.keys()),
            message='Search for tag "{0}" in category:'.format(target_tag)
        )

        if not selected_switch_category:
            import sys; sys.exit()

        target_category = options_category[selected_switch_category]

        elements = DB.FilteredElementCollector(doc, curview.Id)\
                        .OfCategory(target_category)\
                        .WhereElementIsNotElementType()\
                        .ToElements()

        if not elements:
            forms.alert("No elements of {0} found in active view".format(selected_switch_category), exitscript=True)

        for ele in elements:
            family_types_elements[ele.Name].append(ele)

        for k, v in family_types_elements.items():
            param_dic[k].append(v[0].GetOrderedParameters())

        col_para = []
        for k, v in param_dic.items():
            parameters = [j.Definition.Name for i in v for j in i]
            for para in parameters:
                if para not in col_para:
                    col_para.append(para)

        options_parameter = {k: v for k, v in zip(col_para, col_para)}

        selected_switch_parameter = forms.CommandSwitchWindow.show(
            sorted(options_parameter.keys()),
            message='Search for instance parameter'
        )

        if not selected_switch_parameter:
            import sys; sys.exit()

        target_parameter = options_parameter[selected_switch_parameter]

        for wall in elements:
            para_list = wall.GetParameters(target_parameter)
            if len(para_list) > 1:
                forms.alert("More than one parameter with name {0} found".format(target_parameter), exitscript=True)
            try:
                para_value = para_list[0].AsString()
            except:
                forms.alert("This tool is only for searching text based tags", exitscript=True)

            if para_value:
                para_value = para_value.lower()
            if para_value == target_tag:
                wall_id.append(wall.Id)
                wall_id_list = List[DB.ElementId](wall_id)

        if wall_id:
            revit.get_selection().set_to(wall_id)
        else:
            forms.alert('No {0} with tag "{1}" found!!!'.format(selected_switch_category, target_tag))

    except Exception as e:
        print(e)
