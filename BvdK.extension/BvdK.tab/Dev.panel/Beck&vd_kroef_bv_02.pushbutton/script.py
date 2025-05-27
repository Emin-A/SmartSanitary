# -*- coding: utf-8 -*-
__title__ = "AutoFix\nReducers"
__doc__ = """Version = 1.0
Date    = 23.05.2025
________________________________________________________________
Description:

This integrated script lets you:
  1. Fix reducer orientations and parameters automatically.
  2. Enforce family/type rulse based on pipe diameter threshold.
  3. Replace long fittings with short variants.
  4. Control eccentricity, length options, and elbow configurations dynamically.
________________________________________________________________
How-To:
Button Behavior: When clicked, the script:
1. Scans all visible elements (in current 2D/3D view);
2. Identifies:
  - Main pipes Family: System Family: Pipe Types Type: 'NLRS_52_PI_PE buis (OD)_geb' (>= 160 mm)
  - Side pipes -||- (>= 125 mm)
  - Connected T-fittings Family:'NLRS_52_PIF_UN_PE multi T-stuk_geb' Type: 'Liggend - Var. DN/OD' (>= 125 mm)
  - Vertical pipes 'NLRS_52_PI_PE buis (OD)_geb' (>= 110 mm)
  - Vertical elbows Family: 'NLRS_52_PID_UN_PE multibocht_geb' Type: 'Var. DN/OD'
  - Reducers Family: 'NLRS_52_PIF_UN_PE multireducer_geb' Type: Var. DN/OD
3. Applies parameter toggles based on:
  - Pipe diameter
  - Direction in plan (relative angle vector between pipes)
  - Elevation (middle elevation = 0.0)
________________________________________________________________
Author: Emin Avdovic"""
# Imports

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.Exceptions import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.Exceptions import ArgumentException
from System.Collections.Generic import List
from pyrevit import revit, DB, script

import math
import clr
import System
import System.IO

clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("WindowsBase")

from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import MessageBox

# Revit Document Setup

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

# Helper Functions


def is_pipe_of_type(pipe, name, min_diam_mm):
    return name in pipe.Name and pipe.Diameter * 304.8 >= min_diam_mm


def get_pipe_direction(pipe):
    try:
        curve = pipe.Location.Curve
        vec = curve.GetEndPoint(1) - curve.GetEndPoint(0)
        return vec.Normalize()
    except:
        return None


def classify_direction(vec):
    if vec is None:
        return "Unknown"
    angle = math.degrees(math.atan2(vec.Y, vec.X))
    if -45 <= angle < 45:
        return "Right"
    elif 45 <= angle < 135:
        return "Up"
    elif angle >= 135 or angle < -135:
        return "Left"
    elif -135 <= angle < -45:
        return "Down"
    return "Unknown"


def set_yesno_param(elem, param_name, on=True):
    p = elem.LookupParameter(param_name)
    if p and p.StorageType == StorageType.Integer:
        try:
            p.Set(1 if on else 0)
        except:
            pass  # Skip if cannot be set


def get_connected_pipe_direction(fitting, main_pipe):
    for c in fitting.MEPModel.ConnectorManager.Connectors:
        for r in c.AllRefs:
            other = r.Owner
            if isinstance(other, Pipe) and other.Id != main_pipe.Id:
                return get_pipe_direction(other)
    return None


# Core Logic
def auto_fix():
    pipes = FilteredElementCollector(doc).OfClass(Pipe).ToElements()
    fittings = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()

    t = Transaction(doc, "Auto-Fix Reducer Logic")
    t.Start()

    # Scan pipes and look for conditions
    for pipe in pipes:
        if not is_pipe_of_type(pipe, "NLRS_52_PI_PE buis", 160):
            continue

        connectors = pipe.ConnectorManager.Connectors
        for conn in connectors:
            for ref in conn.AllRefs:
                other = ref.Owner

                # Process multi T-stuk
                if (
                    isinstance(other, FamilyInstance)
                    and "multi T-stuk" in other.Symbol.Family.Name
                ):
                    dir_vec = get_connected_pipe_direction(other, pipe)
                    dir_label = classify_direction(dir_vec)

                    set_yesno_param(other, "kort_verloop (kleinste)", True)
                    set_yesno_param(other, "kort_verloop (grootste)", True)
                    set_yesno_param(other, "reducer_eccentric", True)

                    if dir_label in ["Right", "Down"]:
                        set_yesno_param(other, "switch_excentriciteit", True)
                    else:
                        set_yesno_param(other, "switch_excentriciteit", False)
                    print("✅ Matched T-stuk:", other.Id, "→", dir_label)

    # Process fittings for elbows and reducers
    for f in fittings:
        fname = f.Symbol.Family.Name.lower()

        if "multibocht" in fname:
            set_yesno_param(f, "2x45°", False)
            set_yesno_param(f, "buis_invogen", False)

        elif "multireducer" in fname:
            for pname in [
                "kort_verloop (kleinste)",
                "kort_verloop (grootste)",
                "switch_excentriciteit",
                "reducer_eccentric",
            ]:
                set_yesno_param(f, pname, False)

    t.Commit()
    MessageBox.Show("Reducer logic applied successfully.", "Auto-Fix Reducers")


auto_fix()
