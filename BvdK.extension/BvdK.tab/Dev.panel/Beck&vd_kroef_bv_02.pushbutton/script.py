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
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import *
from Autodesk.Revit.Exceptions import *
from Autodesk.Revit.DB.Structure import *
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


# Helpers
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
    # Do not allow disabling 'reducer_eccentric'
    if param_name == "reducer_eccentric" and not on:
        return
    p = elem.LookupParameter(param_name)
    if p and p.StorageType == StorageType.Integer:
        try:
            p.Set(1 if on else 0)
        except:
            pass


def get_connected_pipe_direction(fitting, main_pipe):
    try:
        for c in fitting.MEPModel.ConnectorManager.Connectors:
            for r in c.AllRefs:
                other = r.Owner
                if isinstance(other, Pipe) and other.Id != main_pipe.Id:
                    return get_pipe_direction(other)
    except:
        pass
    return None


def is_reducer_fully_connected(fitting):
    try:
        connectors = fitting.MEPModel.ConnectorManager.Connectors
        connected = [
            r for c in connectors for r in c.AllRefs if r.Owner.Id != fitting.Id
        ]
        return len(connected) >= 2
    except:
        return False


def try_update_fitting(elem_id, param_map):
    fitting = doc.GetElement(elem_id)
    if fitting is None:
        return False
    t = Transaction(doc, "Update Fitting")
    t.Start()
    try:
        for name in param_map:
            set_yesno_param(fitting, name, param_map[name])
        t.Commit()
        return True
    except:
        t.RollBack()
        return False


# === Core Logic ===


def auto_fix():
    pipes = FilteredElementCollector(doc).OfClass(Pipe).ToElements()
    fittings = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
    visited = set()
    updated = 0
    skipped = 0

    tg = TransactionGroup(doc, "Safe Reducer Update")
    tg.Start()

    # --- Process T-stuks ---
    for pipe in pipes:
        if not is_pipe_of_type(pipe, "NLRS_52_PI_PE buis", 160):
            continue

        for conn in pipe.ConnectorManager.Connectors:
            for ref in conn.AllRefs:
                other_id = ref.Owner.Id
                if other_id.IntegerValue in visited:
                    continue
                visited.add(other_id.IntegerValue)

                other = doc.GetElement(other_id)
                if other is None or not isinstance(other, FamilyInstance):
                    continue
                if not other.Symbol.Family.Name.startswith(
                    "NLRS_52_PIF_UN_PE multi T-stuk"
                ):
                    continue

                dir_vec = get_connected_pipe_direction(other, pipe)
                dir_label = classify_direction(dir_vec)

                param_map = {
                    "kort_verloop (kleinste)": True,
                    "kort_verloop (grootste)": True,
                    "reducer_eccentric": True,
                    "switch_excentriciteit": dir_label in ["Right", "Down"],
                }

                if try_update_fitting(other.Id, param_map):
                    print("\u2705 Matched T-stuk:", other.Id, "\u2192", dir_label)
                    updated += 1
                else:
                    print("⚠️ Skipped T-stuk (not editable):", other.Id)
                    skipped += 1

    # --- Process elbows and reducers last ---
    for f in fittings:
        try:
            f_id = f.Id
            f_valid = doc.GetElement(f_id)
            if f_valid is None:
                continue
            name = f_valid.Symbol.Family.Name.lower()

            if "multibocht" in name:
                param_map = {"2x45°": False, "buis_invogen": False}
                if try_update_fitting(f_id, param_map):
                    updated += 1
                else:
                    skipped += 1

            elif "multireducer" in name:
                if is_reducer_fully_connected(f_valid):
                    skipped += 1
                    continue  # ✅ Skip fully connected reducers
                param_map = {
                    "kort_verloop (kleinste)": False,
                    "kort_verloop (grootste)": False,
                    "switch_excentriciteit": False,
                    "reducer_eccentric": False,  # Safe override (won’t be turned off)
                }
                if try_update_fitting(f_id, param_map):
                    updated += 1
                else:
                    skipped += 1
        except Exception as e:
            print("⚠️ Skipped invalid fitting:", str(e))
            skipped += 1
            continue
    tg.Assimilate()
    MessageBox.Show(
        "✅ Finished.\nUpdated: " + str(updated) + "\nSkipped: " + str(skipped),
        "AutoFix Reducers",
    )


auto_fix()
