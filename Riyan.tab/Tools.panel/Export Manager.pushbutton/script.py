# -*- coding: utf-8 -*-
import clr
import os
import re

from pyrevit import revit, DB, UI, forms

doc = revit.doc
uidoc = revit.uidoc

# ------------------------------------------------------------------------------
# View Models
# ------------------------------------------------------------------------------
class SheetViewModel(object):
    def __init__(self, sheet):
        self.Sheet = sheet
        self.SheetNumber = sheet.SheetNumber
        self.SheetName = sheet.Name
        self._is_selected = False
        
    @property
    def IsSelected(self):
        return self._is_selected
        
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value

# ------------------------------------------------------------------------------
# Helper for Custom Naming
# ------------------------------------------------------------------------------
def get_param_value(element, param_name_or_builtin):
    if isinstance(param_name_or_builtin, DB.BuiltInParameter):
        p = element.get_Parameter(param_name_or_builtin)
    else:
        p = element.LookupParameter(param_name_or_builtin)
        
    if p:
        return p.AsValueString() or p.AsString() or ""
    return ""

def generate_filename(sheet, rule):
    name = rule
    # Sheet Number
    name = name.replace("<Sheet Number>", sheet.SheetNumber)
    # Sheet Name
    name = name.replace("<Sheet Name>", sheet.Name)
    # Project Number
    proj_num = get_param_value(doc.ProjectInformation, DB.BuiltInParameter.PROJECT_NUMBER)
    name = name.replace("<Project Number>", proj_num)
    # Project Name
    proj_name = get_param_value(doc.ProjectInformation, DB.BuiltInParameter.PROJECT_NAME)
    name = name.replace("<Project Name>", proj_name)
    # Current Revision
    rev = get_param_value(sheet, DB.BuiltInParameter.SHEET_CURRENT_REVISION)
    name = name.replace("<Current Revision>", rev)
    # Drawn By
    drawn = get_param_value(sheet, DB.BuiltInParameter.SHEET_DRAWN_BY)
    name = name.replace("<Drawn By>", drawn)
    # Checked By
    checked = get_param_value(sheet, DB.BuiltInParameter.SHEET_CHECKED_BY)
    name = name.replace("<Checked By>", checked)

    # Clean invalid characters
    invalid_chars = '<>:"/\\|?*'
    for c in invalid_chars:
        name = name.replace(c, "_")
        
    return name.strip()

# ------------------------------------------------------------------------------
# UI Forms
# ------------------------------------------------------------------------------
class ExportManagerForm(forms.WPFWindow):
    def __init__(self, xaml_file_name, sheets):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        self.export_path = ""
        self.sheets = [SheetViewModel(s) for s in sheets]
        self.sheets.sort(key=lambda x: x.SheetNumber)
        self.SheetListBox.ItemsSource = self.sheets
        
        # Populate Setups
        self.print_settings = list(DB.FilteredElementCollector(doc).OfClass(DB.PrintSetting).ToElements())
        self.dwg_settings = list(DB.FilteredElementCollector(doc).OfClass(DB.ExportDWGSettings).ToElements())
        
        self.pdf_setting_names = [ps.Name for ps in self.print_settings]
        self.dwg_setting_names = [ds.Name for ds in self.dwg_settings]
        
        # Add 'Default' as first option
        self.pdf_setting_names.insert(0, "<In-Session / Default>")
        self.dwg_setting_names.insert(0, "<In-Session / Default>")
        
        self.CmbPdfSetup.ItemsSource = self.pdf_setting_names
        self.CmbDwgSetup.ItemsSource = self.dwg_setting_names
        
        if self.pdf_setting_names: self.CmbPdfSetup.SelectedIndex = 0
        if self.dwg_setting_names: self.CmbDwgSetup.SelectedIndex = 0
        self.CmbInsertTag.SelectedIndex = 0
        
    def TitleBar_MouseDown(self, sender, e):
        if e.ChangedButton == UI.MouseButtons.Left:
            self.DragMove()
            
    def CloseBtn_Click(self, sender, e):
        self.DialogResult = False
        self.Close()
        
    def CbSelectAll_Click(self, sender, e):
        is_checked = self.CbSelectAll.IsChecked
        for sv in self.sheets:
            sv.IsSelected = is_checked
        self.SheetListBox.Items.Refresh()

    def CmbInsertTag_SelectionChanged(self, sender, e):
        if not hasattr(self, 'TxtNamingRule'):
            return
        if sender.SelectedIndex > 0:
            tag = sender.SelectedItem.Content
            current_text = self.TxtNamingRule.Text
            self.TxtNamingRule.Text = current_text + "<{}>".format(tag)
            # Reset selection
            sender.SelectedIndex = 0

    def BtnBrowse_Click(self, sender, e):
        selected_folder = forms.pick_folder(title="Select Export Destination")
        if selected_folder:
            self.export_path = selected_folder
            self.TxtExportPath.Text = self.export_path

    def RunBtn_Click(self, sender, e):
        self.selected_sheets = [sv.Sheet for sv in self.sheets if sv.IsSelected]
        if not self.selected_sheets:
            forms.alert("Please select at least one sheet.")
            return
            
        self.export_path = self.TxtExportPath.Text
        if not self.export_path or not os.path.isdir(self.export_path):
            forms.alert("Please select a valid export directory.")
            return
            
        self.export_pdf = self.CbPDF.IsChecked
        self.export_dwg = self.CbDWG.IsChecked
        self.naming_rule = self.TxtNamingRule.Text
        
        # Get selected setups
        pdf_idx = self.CmbPdfSetup.SelectedIndex
        self.selected_pdf_setting = self.print_settings[pdf_idx - 1] if pdf_idx > 0 else None
        
        dwg_idx = self.CmbDwgSetup.SelectedIndex
        self.selected_dwg_setting = self.dwg_settings[dwg_idx - 1] if dwg_idx > 0 else None
        
        if not self.export_pdf and not self.export_dwg:
            forms.alert("Please select at least one export format.")
            return
            
        self.DialogResult = True
        self.Close()

# ------------------------------------------------------------------------------
# Export Logic
# ------------------------------------------------------------------------------
def export_dwg(folder, sheet, filename, dwg_setting):
    opt = DB.DWGExportOptions()
    if dwg_setting:
        opt = dwg_setting.GetDWGExportOptions()
        
    from System.Collections.Generic import List
    views = List[DB.ElementId]()
    views.Add(sheet.Id)
        
    doc.Export(folder, filename, views, opt)

def export_pdf_2022(folder, sheet, filename, print_setting):
    try:
        opt = DB.PDFExportOptions()
        opt.FileName = filename
        
        # Try to map print_setting to PDFExportOptions if possible
        if print_setting:
            try:
                ps = print_setting.PrintParameters
                opt.ZoomType = ps.ZoomType
                opt.ZoomPercentage = ps.Zoom
            except:
                pass
        
        from System.Collections.Generic import List
        views = List[DB.ElementId]()
        views.Add(sheet.Id)
        
        doc.Export(folder, views, opt)
        return True
    except Exception as e:
        import traceback
        forms.alert("Failed to export PDF for sheet {}:\n{}".format(sheet.SheetNumber, traceback.format_exc()))
        return False

# ------------------------------------------------------------------------------
# Main Execution
# ------------------------------------------------------------------------------
def main():
    sheets = DB.FilteredElementCollector(doc)\
               .OfCategory(DB.BuiltInCategory.OST_Sheets)\
               .WhereElementIsNotElementType()\
               .ToElements()
               
    if not sheets:
        forms.alert("No Sheets found in the current project.")
        return
        
    xaml_path = os.path.join(os.path.dirname(__file__), "ExportUI.xaml")
    form = ExportManagerForm(xaml_path, sheets)
    
    if not form.ShowDialog():
        return
        
    selected = form.selected_sheets
    folder = form.export_path
    rule = form.naming_rule
    
    revit_version = int(__revit__.Application.VersionNumber)
    
    with forms.ProgressBar(title="Exporting Sheets...") as pb:
        total = len(selected)
        for idx, sheet in enumerate(selected):
            pb.update_progress(idx, total)
            
            filename = generate_filename(sheet, rule)
            
            if form.export_dwg:
                export_dwg(folder, sheet, filename, form.selected_dwg_setting)
                
            if form.export_pdf:
                if revit_version >= 2022:
                    export_pdf_2022(folder, sheet, filename, form.selected_pdf_setting)
                else:
                    if idx == 0: # Only alert once
                        forms.alert("Native PDF export requires Revit 2022+. PDFs will be skipped.", title="Version Error")
                    
    forms.alert("Export completed successfully!", title="Export Manager")

if __name__ == '__main__':
    main()
