import omni.ext
import omni.ui as ui
import omni.usd
import carb
import usd
from pxr import Tf
from pxr import Usd, Sdf, UsdShade

# A bunch of messiness needed to work with com
# Thank you Mati, for figuring this out!
import omni.kit.pipapi

omni.kit.pipapi.install("pywin32")

import os
import sys
from pathlib import Path
import pythonwin.pywin

dlls_path = Path(pythonwin.pywin.__file__).parent.parent.parent / "pywin32_system32"
com = Path(pythonwin.pywin.__file__).parent.parent.parent / "win32"
lib = Path(pythonwin.pywin.__file__).parent.parent.parent / "win32" / "lib"
pywin = Path(pythonwin.pywin.__file__).parent.parent.parent / "pythonwin"
carb.log_info(dlls_path)
sys.path.insert(0, str(com))
sys.path.insert(0, str(lib))
sys.path.insert(0, str(pywin))
carb.log_info(sys.path)
os.environ["PATH"] = f"{dlls_path};{os.environ['PATH']}"
carb.log_info(os.environ["PATH"])

import win32com.client

# End of messiness needed to work with outlook.
class WorksheetEvents:

    def OnChange(self, *args):
        #1. Find out if G4 is what changed
        try:
            address = str(args[0].Address)
            if address != '$G$4':
                return
        except Exception as e:
            carb.log_error('Could not detect cell changes' + e)

        #2. If so, check if the excel value is different from the scene value
        stage = omni.usd.get_context().get_stage()

        basePlate = stage.GetPrimAtPath('/World/Vehicle/Chassis/Chassis/BasePlate__refset_MODEL/Mesh_14')

        if not basePlate.IsValid():
            carb.log_error("Can't find Base Plate")
            return
        
        mat = UsdShade.MaterialBindingAPI(basePlate).GetDirectBindingRel().GetTargets()

        OmniverseMatName = mat[0].name
        ExcelMatName = args[0].Value

        if OmniverseMatName == ExcelMatName:
            return
        
        #3. If so, edit the scene to match excel.
        newMat = UsdShade.Material.Get(stage, Sdf.Path('/World/Materials/' + ExcelMatName))
        UsdShade.MaterialBindingAPI(basePlate).Bind(newMat)     


# Any class derived from `omni.ext.IExt` in top level module (defined in `python.modules` of `extension.toml`) will be
# instantiated when extension gets enabled and `on_startup(ext_id)` will be called. Later when extension gets disabled
# on_shutdown() is called.
class StrainflowExcelverseExtension(omni.ext.IExt):
    # ext_id is current extension id. It can be used with extension manager to query additional information, like where
    # this extension is located on filesystem.
    def on_startup(self, ext_id):
        print("[strainflow.excelverse] strainflow excelverse startup")
        
        self._window = ui.Window("ExcelVerse", width=300, height=300)

        with self._window.frame:
            with ui.VStack():
                
                self._sheet_path = ui.SimpleStringModel("C:\\Users\\ebowman\\source\\repos\\RC-Car-CAD\\BOM.xlsx")
                ui.StringField(self._sheet_path, height=30)
                              
                with ui.HStack(style={"margin": 10}):
                    ui.Spacer()
                    ui.Button("Link", clicked_fn=self.on_Link_Click, width=300, height=300)
                    ui.Spacer()

        # while True:
        #     pythoncom.PumpWaitingMessages()

    def _mat_changed(self, *args):
        # 1. Check if the base plate material in excel is different
        mat = UsdShade.MaterialBindingAPI(self._basePlate).GetDirectBindingRel().GetTargets()

        OmniverseMatName = mat[0].name
        ExcelMatName = self._excel_worksheet.Range('G4').Value
        
        if OmniverseMatName == ExcelMatName:
            return
        
        # 2. If so change it.
        self._excel_worksheet.Range('G4').Value = OmniverseMatName



    def on_Link_Click(self):

        # Link to Scene
        self._stage = omni.usd.get_context().get_stage()

        self._basePlate = self._stage.GetPrimAtPath('/World/Vehicle/Chassis/Chassis/BasePlate__refset_MODEL/Mesh_14')
                
        if not self._basePlate.IsValid():
            carb.log_error("RC Car Scene Not Open")
            return

        mat_attr = self._basePlate.GetAttribute("material:binding")
        self._mat_subs = omni.usd.get_watcher().subscribe_to_change_info_path(mat_attr.GetPath(), self._mat_changed)

        # Link to Excel
        self._excel_app = win32com.client.DispatchEx("excel.application")
        self._excel_app.Visible = True

        #Open workbook instead of grab
        self._excel_workbook = self._excel_app.Workbooks.Open(self._sheet_path.as_string)

        try:
            if hasattr(self._excel_workbook, 'Worksheets'):
                self._excel_worksheet = self._excel_workbook.Worksheets(1)
            else:
                self._excel_worksheet = self._excel_workbook._dispobj_.Worksheets(1)
        except:
            carb.log_info("Could not find Worksheets attribute")
            return

        self._excel_events = win32com.client.WithEvents(self._excel_worksheet, WorksheetEvents)

    def on_shutdown(self):
        
        # Share in livestream
        self._excel_events = None
        self._excel_worksheet = None
        
        if hasattr(self, '_excel_workbook'):
            if self._excel_workbook is not None:
                self._excel_workbook.Close(False)
                self._excel_workbook = None

        if hasattr(self, '_excel_app'):
            if self._excel_app is not None:
                self._excel_app.Application.Quit()
                self._excel_app = None

        print("[strainflow.excelverse] strainflow excelverse shutdown")
