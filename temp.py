# IronPython imports to enable Excel interop
import clr
import os

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# Define working directory
workingDir = AbsUserPathName("F:\EXCEL")


def updateHandler():
    # Update can take long, so disable the Excel warning -
    # "Excel is waiting for another application to complete an OLE action"
    ex.Application.DisplayAlerts = False


    # Define key ranges in the Workbook
    modulusCell = worksheet.Range["A3"]
    diffusionCell = worksheet.Range["B3"]
    growthCell = work.sheet.Range["C3"]
    defCell = worksheet.Range["D3"]


    # Get the Workbench Parameters
    Force = Parameters.GetParameter(Name="P1")
    Hpressure = Parameters.GetParameter(Name="P2")
    Pressure = Parameters.GetParameter(Name="P3")
    Total_Volume = Parameters.GetParameter(Name="P4")

    # Assign values to the input parameters
    parameter1Param.Expression = ForceCell.Value2.ToString()
    parameter2Param.Expression = HpressureCell.Value2.ToString()
    parameter3Param.Expression = PressurehCell.Value2.ToString()

    # Mark the deformation parameter as updating in the workbook
    defCell.Value2 = "Updating..."

    # Run the project update
    Update()

    # Update the workbook value from the WB parameter
    defCell.Value2 = defParam.Value

    # restore alert setting
    ex.Application.DisplayAlerts = True

# Open the Workbench Project
Open(FilePath=os.path.join(workingDir, "try.wbpj"))

# Open Excel and the workbook
ex = Excel.ApplicationClass()
ex.Visible = True
workbook = ex.Workbooks.Open(os.path.join(workingDir, "ParameterExample1.xlsx"))
worksheet = workbook.ActiveSheet

# Apply the update handler to the workbook button
OLEbutton = worksheet.OLEObjects("CommandButton1")
commandButton = OLEbutton.Object
commandButton.CLICK += updateHandler
