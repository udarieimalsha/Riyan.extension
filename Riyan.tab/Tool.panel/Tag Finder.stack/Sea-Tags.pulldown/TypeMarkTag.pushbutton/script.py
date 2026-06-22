# -*- coding: utf-8 -*-
"""Searches and picks elements in current view based on input tag defined in Type Mark parameter. (Riyan)"""
__doc__ = "Searches and picks elements in current view based on input tag defined in Type Mark parameter."
__title__ = "Type Mark Tag\nFinder"

import os
import clr
clr.AddReference("System.Windows.Presentation")
import System

from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit.forms import WPFWindow


class TypeMarkSearchWindow(WPFWindow):
    """Riyan-themed WPF window for type mark tag input."""
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
            forms.alert("Please enter a Type Mark value to search for.", title="Invalid Input")

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
    xaml_file = os.path.join(os.path.dirname(__file__), "typemark_input.xaml")
    target_tag = None

    if os.path.exists(xaml_file):
        try:
            window = TypeMarkSearchWindow(xaml_file)
            window.ShowDialog()
            if window.DialogResult and window.result:
                target_tag = window.result
        except Exception as e:
            forms.alert("Could not load input window:\n{}".format(e))
            target_tag = str(forms.ask_for_string("Enter tag name"))
    else:
        target_tag = str(forms.ask_for_string("Enter tag name"))

    if not target_tag or target_tag == "None":
        import sys; sys.exit()

    try:
        flag = 0
        wall_id = []

        categorys = [
            DB.BuiltInCategory.OST_StructuralColumns,
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_StructuralFoundation
        ]

        for cat in categorys:
            target_category = cat
            target_parameter = DB.BuiltInParameter.ALL_MODEL_TYPE_MARK

            param_id = DB.ElementId(target_parameter)
            param_prov = DB.ParameterValueProvider(param_id)
            param_equality = DB.FilterStringEquals()

            value_rule = DB.FilterStringRule(param_prov, param_equality, target_tag, True)
            param_filter = DB.ElementParameterFilter(value_rule)

            elements = DB.FilteredElementCollector(doc, curview.Id)\
                    .OfCategory(target_category)\
                    .WhereElementIsNotElementType()\
                    .WherePasses(param_filter)\
                    .ToElementIds()

            if elements:
                uidoc.Selection.SetElementIds(elements)
                flag = 1
                break

        if flag == 0:
            forms.alert('Tag "{0}" not found in this view!!!'.format(target_tag))

    except Exception as e:
        print(e)
