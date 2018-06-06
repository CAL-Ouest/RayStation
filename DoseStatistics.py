import wpf, os, sys, System, clr, random
from System import Windows
from System.Windows import *
from System.Windows.Input import Key
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Windows.Controls import *
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlWBATemplate, XlChartType
import math

from connect import *

def create_array(m, n):
    dims = System.Array.CreateInstance(System.Int32, 2)
    dims[0] = m
    dims[1] = n
    return System.Array.CreateInstance(System.Object, dims)

def WritePatientInfo(worksheet, current_cell, patient_name, beam_set_name):
    print("WritePatientInfo")
    array = create_array(2, 2)
    array[0,0] = "Patient"
    array[0,1] = patient_name
    array[1,0] = "Beam set"
    array[1,1] = beam_set_name
    array_range = worksheet.Range(current_cell, current_cell.Cells(2, 2))
    array_range.Value = array
    return current_cell.Cells(4, 1)

def WriteClinicalGoals(worksheet, current_cell, plan):
    print("WriteClinicalGoals")
    # if goal.ForRegionOfInterest.Name in rois]
    clinical_goals = [goal for goal in plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions]

def WriteDoseStatistics(worksheet, current_cell, plan, patient):
    print("WriteDoseStatistics")
    structure_set = plan.GetStructureSet()
    plan_dose = plan.TreatmentCourse.TotalDose
    roi_names = [r.OfRoi.Name for r in structure_set.RoiGeometries if r.PrimaryShape != None]
    ncols = 9
    nrows = len(roi_names)

    header_row = create_array(1, ncols)
    data_array = create_array(nrows, ncols)

    header_row[0, 0] = "ROI"
    header_row[0, 1] = "Volume [cc]"
    header_row[0, 2] = "D99 [cGy]"
    header_row[0, 3] = "D98 [cGy]"
    header_row[0, 4] = "D95 [cGy]"
    header_row[0, 5] = "Average [cGy]"
    header_row[0, 6] = "D50 [cGy]"
    header_row[0, 7] = "D2 [cGy]"
    header_row[0, 8] = "D1 [cGy]"

    for roi_idx, roi_name in enumerate(roi_names):
        volume = patient.PatientModel.StructureSets[0].RoiGeometries[roi_name].GetRoiVolume()
        d99, d98, d95, d50, d2, d1 = plan_dose.GetDoseAtRelativeVolumes(RoiName = roi_name, RelativeVolumes = [0.99, 0.98, 0.95, 0.5, 0.02, 0.01])
        average = plan_dose.GetDoseStatistics(RoiName=roi_name, DoseType='Average')

        data_array[roi_idx, 0] = roi_key
        data_array[roi_idx, 1] = volume
        data_array[roi_idx, 2] = d99
        data_array[roi_idx, 3] = d98
        data_array[roi_idx, 4] = d95
        data_array[roi_idx, 5] = average
        data_array[roi_idx, 6] = d50
        data_array[roi_idx, 7] = d2
        data_array[roi_idx, 8] = d1

    print("Create Header")
    header_range = worksheet.Range(current_cell, current_cell.Cells(1, ncols))
    header_range.Value = header_row

    # write data
    print("Create DataArray; rows={:d}, cols={:d}".format(data_array.GetLength(0), data_array.GetLength(1)))
    data_range = worksheet.Range(current_cell.Cells(2, 1),
                                 current_cell.Cells(1 + data_array.GetLength(0), data_array.GetLength(1)))
    data_range.Value = data_array

    return current_cell.Cells(nrows + 3, 1)

# Execution starts here
try:
    case = get_current("Case")
    patient = get_current("Patient")
    plan = get_current("Plan")
    beam_set = get_current("BeamSet")

    excel = ApplicationClass(Visible=True)
    workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    worksheet = workbook.Worksheets[1]
    current_cell = worksheet.Cells(1, 1)
    current_cell = current_cell.Cells(1, 10)
    current_cell = WritePatientInfo(worksheet, current_cell, patient.PatientName, beam_set.DicomPlanLabel)
    current_cell = WriteDoseStatistics(worksheet, current_cell, plan, patient)

except Exception as e:\
    wait_user_input("Exception thrown:" + str(e))

finally:
    if worksheet:
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet)
    if workbook:
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook)
    if excel:
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel)
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()




