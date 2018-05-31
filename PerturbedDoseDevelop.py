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

class Perturbation:
    def __init__(self, dDensity, dX, dY, dZ):
        self.DensityPerturbation = dDensity
        self.IsocenterShift = {'x': dX, 'y': dY, 'z': dZ}

    def __str__(self):
        return "{0:.2f}".format(float(self.DensityPerturbation)) + "_" + \
               "{0:.2f}".format(float(self.IsocenterShift.get('x'))) + "_" + \
               "{0:.2f}".format(float(self.IsocenterShift.get('y'))) + "_" + \
               "{0:.2f}".format(float(self.IsocenterShift.get('z')))

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.is_close(self.dDensity, other.dDensity) and self.is_close(self.dX, other.dX) and self.is_close(self.dY, other.dY) and self.is_close(self.dZ, other.dZ)

    def is_close(self, a, b, rel_tol=1e-9, abs_tol=0.0):
        return abs(a-b) <= max(rel_tol * max(abs(a), abs(b)), abs_tol)

def create_array(m, n):
    dims = System.Array.CreateInstance(System.Int32, 2)
    dims[0] = m
    dims[1] = n
    return System.Array.CreateInstance(System.Object, dims)

def WritePerturbationSettings(worksheet, current_cell, perturbation_settings):
    print("WritePerturbationSettings")
    nrows = len(perturbation_settings)
    array = create_array(nrows, 2)
    for index, key in enumerate(perturbation_settings):
        array[index, 0] = key
        array[index, 1] = perturbation_settings[key]
    array_range = worksheet.Range(current_cell, current_cell.Cells(nrows, 2))
    array_range.Value = array
    return current_cell.Cells(2 + nrows, 1)

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

def WriteDoseStatistics(worksheet, current_cell, stat_pert, fractions):
    print("WriteDoseStatistics")
    npert = len(stat_pert)  # npert = len(next (iter (dvhs.values()))) # python >= 3
    nrois = len(stat_pert.itervalues().next())
    nrows = nrois * npert
    ncols = 8

    header_row = create_array(1, ncols)
    data_array = create_array(nrows, ncols)

    header_row[0, 0] = "ROI"
    header_row[0, 1] = "Pert"
    header_row[0, 2] = "D99 [cGy]"
    header_row[0, 3] = "D98 [cGy]"
    header_row[0, 4] = "D95 [cGy]"
    header_row[0, 5] = "D50 [cGy]"
    header_row[0, 6] = "D2 [cGy]"
    header_row[0, 7] = "D1 [cGy]"

    rois = stat_pert.itervalues().next().keys()
    for roi_idx, roi_key in enumerate(rois):
        for p_idx, p_key in enumerate(stat_pert):
            irow = roi_idx * npert + p_idx
            data_array[irow, 0] = roi_key
            data_array[irow, 1] = p_key
            data_array[irow, 2] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[0])))
            data_array[irow, 3] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[1])))
            data_array[irow, 4] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[2])))
            data_array[irow, 5] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[3])))
            data_array[irow, 6] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[4])))
            data_array[irow, 7] = int(round(fractions*float(stat_pert.get(p_key).get(roi_key)[5])))

    # create header
    print("Create Header")
    header_range = worksheet.Range(current_cell, current_cell.Cells(1, ncols))
    header_range.Value = header_row

    # write data
    print("Create DataArray; rows={:d}, cols={:d}".format(data_array.GetLength(0), data_array.GetLength(1)))
    data_range = worksheet.Range(current_cell.Cells(2, 1),
                                 current_cell.Cells(1 + data_array.GetLength(0), data_array.GetLength(1)))
    data_range.Value = data_array

    return current_cell.Cells(nrows + 3, 1)

def CreateCharts(worksheet, current_cell, N):
    print("CreateCharts")
    charts = [worksheet.ChartObjects().Add(current_cell.Left + 10.0, current_cell.Top + 10 + 300 * idx, 500.0,
                                           300.0).Chart for idx in range(N)]
    return charts

def GetMinMaxCurves(dvhs, bins):
    print("GetMinMaxCurves")
    rois = dvhs.itervalues().next().keys()

    # init min/max perturbated dvh to the first (random) one
    first_pert = dvhs.keys()[0]
    dvhs_minmax = {}
    dvhs_minmax["min"] = {roi : list(dvhs.get(first_pert).get(roi)) for roi in rois}
    dvhs_minmax["max"] = {roi : list(dvhs.get(first_pert).get(roi)) for roi in rois}

    # fill with min max values
    for roi_idx, roi_key in enumerate(rois):
        for p_idx, p_key in enumerate(dvhs):
            for bin_idx, bin in enumerate(bins):
                value = float(dvhs.get(p_key).get(roi_key)[bin_idx])
                dvhs_minmax.get("min").get(roi_key)[bin_idx] = min(value, float(dvhs_minmax.get("min").get(roi_key)[bin_idx]))
                dvhs_minmax.get("max").get(roi_key)[bin_idx] = max(value, float(dvhs_minmax.get("max").get(roi_key)[bin_idx]))

    return dvhs_minmax

def WriteIndividualCurves(worksheet, charts, current_cell, dvhs_pert, bins, colors):
    print("WriteIndividualCurves")
    nrows = len(bins)
    pert_names = dvhs_pert.keys()
    npert = len(dvhs_pert)  # npert = len(next (iter (dvhs.values()))) # python >= 3
    nrois = len(dvhs_pert.itervalues().next())
    ncols = 1 + nrois * npert

    header_row = create_array(1, ncols)
    data_array = create_array(nrows, ncols)

    header_row[0, 0] = "Bin"
    for bin_idx, bin in enumerate(bins):
        data_array[bin_idx, 0] = float(bin)

    rois = dvhs_pert.itervalues().next().keys()
    for roi_idx, roi_key in enumerate(rois):
        for p_idx, p_key in enumerate(dvhs_pert):
            col_idx = 1 + npert * roi_idx + p_idx
            header_row[0, col_idx] = roi_key + "_" + p_key
            for bin_idx, bin in enumerate(bins):
                data_array[bin_idx, col_idx] = float(dvhs_pert.get(p_key).get(roi_key)[bin_idx])

    # create header
    print("Create Header")
    header_range = worksheet.Range(current_cell, current_cell.Cells(1, ncols))
    header_range.Value = header_row

    # write data
    print("Create DataArray; rows={:d}, cols={:d}".format(data_array.GetLength(0), data_array.GetLength(1)))
    data_range = worksheet.Range(current_cell.Cells(2, 1), current_cell.Cells(1 + data_array.GetLength(0), data_array.GetLength(1)))
    data_range.Value = data_array

    print("Create series collections")
    series_collections = [c.seriesCollection() for c in charts]
    print("Filling data")
    x_values = worksheet.Range(current_cell.Cells(2, 1), current_cell.Cells(1 + nrows, 1))
    for roi_idx, roi in enumerate(rois):
        charts[roi_idx].Legend.Clear()
        for pert_idx in range(0, npert):
            col_idx = 2 + npert*roi_idx + pert_idx
            y_values = worksheet.Range(current_cell.Cells(2, col_idx), current_cell.Cells(1 + nrows, col_idx))
            s = series_collections[roi_idx].NewSeries()
            s.Name = pert_names[pert_idx]
            s.Border.Color = int(colors.get(roi).B.ToString("X2") + colors.get(roi).G.ToString("X2") + colors.get(roi).R.ToString("X2"), 16)
            s.XValues = x_values
            s.Values = y_values
        charts[roi_idx].ChartType = XlChartType.xlXYScatterLinesNoMarkers
        title = 'DVH ' + roi
        value_title = 'Volume [%]'
        category_title = 'Dose [cGy]'
        charts[roi_idx].ChartWizard(Title=title, ValueTitle=value_title, CategoryTitle=category_title)

    for s in series_collections:
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(s)

    return current_cell.Cells(nrows + 2, 1)

def WriteMinMaxCurves(worksheet, chart, current_cell, dvhs_minmax, bins, colors):
    print("WriteMinMaxCurves")

    nrows = len(bins)
    nrois = len(dvhs_minmax.itervalues().next())
    ncols = 1 + 2 * nrois

    header_row = create_array(1, ncols)
    data_array = create_array(nrows, ncols) # std + min + max = 3

    header_row[0, 0] = "Bin"
    for bin_idx, bin in enumerate(bins):
        data_array[bin_idx, 0] = float(bin)

    rois = dvhs_minmax.itervalues().next().keys()

    for roi_idx, roi_key in enumerate(rois):
        col_idx_min = 1 + 2 * roi_idx
        col_idx_max = 2 + 2 * roi_idx
        header_row[0, col_idx_min] = roi_key + "_min"
        header_row[0, col_idx_max] = roi_key + "_max"
        for bin_idx, bin in enumerate(bins):
            data_array[bin_idx, col_idx_max] = dvhs_minmax.get("max").get(roi_key)[bin_idx]
            data_array[bin_idx, col_idx_min] = dvhs_minmax.get("min").get(roi_key)[bin_idx]

    # create header
    print("Create Header")
    header_range = worksheet.Range(current_cell, current_cell.Cells(1, ncols))
    header_range.Value = header_row

    print("Create DataArray; rows={:d}, cols={:d}".format(data_array.GetLength(0), data_array.GetLength(1)))
    data_range = worksheet.Range(current_cell.Cells(2, 1), current_cell.Cells(1 + data_array.GetLength(0), data_array.GetLength(1)))
    data_range.Value = data_array

    chart.ChartType = XlChartType.xlXYScatterLinesNoMarkers
    title = 'DVH Min-Max Curves'
    value_title = 'Volume [%]'
    category_title = 'Dose [cGy]'
    chart.ChartWizard(Title=title, ValueTitle=value_title, CategoryTitle=category_title)
    series_collection = chart.seriesCollection()

    x_values = worksheet.Range(current_cell.Cells(2, 1), current_cell.Cells(1 + nrows, 1))
    for roi_idx, roi in enumerate(rois):
        print("Making graph for roi " + roi)
        y_values_min = worksheet.Range(current_cell.Cells(2, 2 + roi_idx * 2), current_cell.Cells(1 + nrows, 2 + roi_idx * 2))
        y_values_max = worksheet.Range(current_cell.Cells(2, 3 + roi_idx * 2), current_cell.Cells(1 + nrows, 3 + roi_idx * 2))

        color_idx = int(colors.get(roi).B.ToString("X2") + colors.get(roi).G.ToString("X2") + colors.get(roi).R.ToString("X2"), 16)
        s_min = series_collection.NewSeries()
        s_max = series_collection.NewSeries()
        s_min.Name = roi + "_min"
        s_max.Name = roi + "_max"
        s_min.Border.Color = color_idx
        s_max.Border.Color = color_idx
        s_min.XValues = x_values
        s_max.XValues = x_values
        s_min.Values = y_values_min
        s_max.Values = y_values_max

    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(series_collection)

    return current_cell.Cells(3 + nrows, 1)

def WriteExcel(patient_name, beam_set_name, perturbation_settings, dvhs_pert, stat_pert, bins, fractions, colors):
    print("WriteExcel")
    print(fractions)
    try:
        excel = ApplicationClass(Visible=True)
        workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        worksheet = workbook.Worksheets[1]
        current_cell = worksheet.Cells(1, 1)
        rois = dvhs_pert.itervalues().next().keys()
        charts = CreateCharts(worksheet, current_cell, len(rois) + 1)
        current_cell = current_cell.Cells(1, 10)
        current_cell = WritePatientInfo(worksheet, current_cell, patient_name, beam_set_name)
        current_cell = WritePerturbationSettings(worksheet, current_cell, perturbation_settings)
        current_cell = WriteDoseStatistics(worksheet, current_cell, stat_pert, fractions)
        current_cell = WriteIndividualCurves(worksheet, charts[1:], current_cell, dvhs_pert, bins, colors)
        dvhs_minmax = GetMinMaxCurves(dvhs_pert, bins)
        current_cell = WriteMinMaxCurves(worksheet, charts[0], current_cell, dvhs_minmax, bins, colors)

    finally:
        if charts:
            for c in charts:
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(c)
        if worksheet:
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet)
        if workbook:
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook)
        if excel:
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel)
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()



# "dose_data" can be from either TreatmentPlan.TreatmentCourse.TotalDose.DoseValues.DoseData
# or Case.TreatmentDelivery.FractionEvaluations[i].DoseOnExaminations[j].DoseEvaluations
#
# returns dictionary {'roi_name' : dvh, ...}, where dvh is a list of floats (dose) with size = len(bins)
#def CalculateDVH(dose, roi_names, bins):
#    dvhs = {roi : dose.GetRelativeVolumeAtDoseValues(RoiName=roi, DoseValues=bins) for roi in roi_names}
#    return dvhs

def ComputeGaussianPerturbedDose(patient, case, plan, beam_set, dialog):
    print("ComputeGaussianPerturbedDose")
    structure_set = plan.GetStructureSet()
    examination_name = structure_set.OnExamination.Name
    fraction_number = 0

    XSigma = dialog.XSigma
    YSigma = dialog.YSigma
    ZSigma = dialog.ZSigma
    DPSigma = float(dialog.DPSigma) / 100.0

    perturbation_settings = { "N pert" : dialog.NoOfCalcs, "X" : XSigma, "Y" : YSigma, "Z" : ZSigma, "Density" : DPSigma, \
                              "Distribution" : "Gaussian" if dialog.Gaussian.IsChecked else "Square"}

    roi_names = dialog.selected_rois

    perturbations = []
    for k in range(dialog.NoOfCalcs):
        if dialog.Gaussian.IsChecked:
            x = random.normalvariate(0.0, XSigma)
            y = random.normalvariate(0.0, YSigma)
            z = random.normalvariate(0.0, ZSigma)
            d = random.normalvariate(0.0, DPSigma)
            d = min(max(d, -0.5), 0.5)
            perturbations.append(Perturbation(d, x, y, z))
        else:
            x = random.uniform(-XSigma, XSigma)
            y = random.uniform(-YSigma, YSigma)
            z = random.uniform(-ZSigma, ZSigma)
            d = random.uniform(-DPSigma, DPSigma)
            d = min(max(d, -0.5), 0.5)
            perturbations.append(Perturbation(d, x, y, z))

    bins = range(0, max(list(plan.TreatmentCourse.TotalDose.DoseValues.DoseData)), int(30))
    bins = [int(b) for b in bins] # Dont want long format, since some raystation functions wont accept it

    print("Calculating perturbed dose distributions")
    dvhs_pert = {}
    stat_pert = {}
    for p in perturbations:
        beam_set.ComputePerturbedDose(DensityPerturbation=p.DensityPerturbation, \
                                      IsocenterShift=p.IsocenterShift, \
                                      ExaminationNames=[examination_name], FractionNumbers=[fraction_number])

        print("Calculating perturbed dvh curves")
        frac_eval = next(f for f in case.TreatmentDelivery.FractionEvaluations if f.FractionNumber == fraction_number)
        dose_exam = next(d for d in frac_eval.DoseOnExaminations if d.OnExamination.Name == examination_name)
        eval_dose = dose_exam.DoseEvaluations[dose_exam.DoseEvaluations.Count-1]
        fractions = beam_set.FractionationPattern.NumberOfFractions
        bins_per_fraction = [float(b)/float(fractions) for b in bins]
        dvhs_pert[str(p)] = {roi : eval_dose.GetRelativeVolumeAtDoseValues(RoiName=roi, DoseValues=bins_per_fraction) for roi in roi_names}
        stat_pert[str(p)] = {roi : eval_dose.GetDoseAtRelativeVolumes(RoiName=roi, RelativeVolumes=[0.99, 0.98, 0.95, 0.50, 0.02, 0.01]) for roi in roi_names}

        if dialog.DeleteOnExit.IsChecked:
                eval_dose.DeleteEvaluationDose()

    print("Create excel sheet")
    colors = {r.OfRoi.Name : r.OfRoi.Color for r in structure_set.RoiGeometries if r.OfRoi.Name in roi_names}
    WriteExcel(patient.PatientName, beam_set.DicomPlanLabel, perturbation_settings, dvhs_pert, stat_pert, bins, fractions, colors)

def GetRobustRois(plan, beam_set):
    optim = plan.PlanOptimizations[beam_set.Number - 1]
    rois_objective_robust = {c.ForRegionOfInterest.Name for c in optim.Objective.ConstituentFunctions if c.UseRobustness}
    rois_constraint_robust = {c.ForRegionOfInterest.Name for c in optim.Constraints if c.UseRobustness}
    return rois_constraint_robust.union(rois_objective_robust)

class GaussianPerturbedDoseDialog(Windows.Window):
    def __init__(self, title, case, patient, plan, beam_set):

        wpf.LoadComponent(self, "N:/Phy/RayStation_Scripts/PerturbedDoseDevelop.xaml")
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Title = title
        self.ControlBox = False
        self.case = case
        self.patient = patient
        self.plan = plan
        self.beam_set = beam_set
        self.rois = {}
        self.selected_rois = []
        robust_rois = GetRobustRois(plan, beam_set)
        structure_set = plan.GetStructureSet()
        for r in structure_set.RoiGeometries:
            if r.OfRoi.Name in robust_rois:
                self.AddRow(r.OfRoi.Name)
        text = "Case: {0} Plan: {1} BeamSet: {2}"
        self.CaseInfo.Text = text.format(self.case.CaseName, self.plan.Name , self.beam_set.DicomPlanLabel)

    def ComputeClicked(self, sender, event):
        if not self.GetData():
            return
        print "Starting perturbed dose calculation"
        for r in self.RoiSelection.Children:
            if r.GetValue(Grid.ColumnProperty) == 0 and r.IsChecked:
                row = r.GetValue(Grid.RowProperty)
                self.selected_rois.append(self.rois[row])
        self.DialogResult = True
        if len(self.selected_rois) > 0:
            ComputeGaussianPerturbedDose(self.patient, self.case, self.plan, self.beam_set, self)
        else:
            print "No rois selected"

    def CloseClicked(self, sender, event):
        self.DialogResult = False

    def AddRow(self, roi_name):
        row = self.RoiSelection.RowDefinitions.Count
        self.RoiSelection.RowDefinitions.Add(RowDefinition())
        cb = CheckBox()
        cb.Margin = Thickness(10,5,5,5)
        cb.SetValue(Grid.RowProperty, row)
        cb.SetValue(Grid.ColumnProperty,0)
        cb.Content = roi_name
        self.RoiSelection.Children.Add(cb)
        self.rois[row] = roi_name

    def NumericKeyDown(self, sender, event):
        if (event.Key >= Key.D0 and event.Key <= Key.D9) or (
                event.Key >= Key.NumPad0 and event.Key <= Key.NumPad9) or event.Key == Key.Tab or event.Key == Key.Enter:
            event.Handled = False
            return
        if sender.Name != 'NoOfCalcs' and (
                event.Key == Key.Decimal or event.Key == Key.OemPeriod) and '.' not in sender.Text:
            event.Handled = False
            return
        event.Handled = True

    def GetData(self):
        no_input = "No input for {0}."
        if self.XSigma.Text == '':
            MessageBox.Show(no_input.format('X isocenter shift'))
            return False
        self.XSigma = float(self.XSigma.Text)

        if self.YSigma.Text == '':
            MessageBox.Show(no_input.format('Y isocenter shift'))
            return False
        self.YSigma = float(self.YSigma.Text)

        if self.ZSigma.Text == '':
            MessageBox.Show(no_input.format('Z isocenter shift'))
            return False
        self.ZSigma = float(self.ZSigma.Text)

        if self.DPSigma.Text == '':
            MessageBox.Show(no_input.format('sigma of density perturbation'))
            return False
        self.DPSigma = float(self.DPSigma.Text)

        if self.NoOfCalcs.Text == '':
            MessageBox.Show(no_input.format('number of calculations'))
            return False
        self.NoOfCalcs = int(self.NoOfCalcs.Text)

        return True


# Execution starts here
try:
    case = get_current("Case")
    patient = get_current("Patient")
    plan = get_current("Plan")
    beam_set = get_current("BeamSet")

    dialog = GaussianPerturbedDoseDialog("Perturbed doses", case, patient, plan, beam_set)
    dialog.ShowDialog()
except Exception as e:
    await_user_input("Exception thrown when fetching patient data:" + str(e))




