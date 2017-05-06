using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace CalculateAirflow
{
    [Transaction(TransactionMode.Automatic)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            try
            {
                
                double totalSupplyCFM = 0;
                double totalReturnCFM = 0;
                double totalExhaustCFM = 0;
                double netAirflow = 0;
                string totalFlowMessage = "";
                string netAirflowMessage = "";
                string displayMessage = "";
                bool hasInvalidSelection = false;
                
                // Select some elements in Revit before invoking this command

                // Get the handle of current document.
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Get the element selection of current document.
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (0 == selectedIds.Count)
                {
                    // If no elements selected.
                    TaskDialog.Show("Revit", "You haven't selected any elements.");
                }
                else
                {
                    
                    foreach (ElementId id in selectedIds)
                    {
                        Element elem = doc.GetElement(id);

                        string systemClassification = string.Empty;
                        systemClassification = elem.LookupParameter("System Classification").AsString();

                        double airFlow = 0;
                        airFlow = elem.LookupParameter("Flow").AsDouble();

                        if (elem.Category.Name == "Air Terminals")
                        {
                            if (systemClassification == "Supply Air")
                            {
                                totalSupplyCFM += airFlow;
                            }

                            if (systemClassification == "Return Air")
                            {
                                totalReturnCFM += airFlow;
                            }

                            if (systemClassification == "Exhaust Air")
                            {
                                totalExhaustCFM += airFlow;
                            }

                        }
                        else
                        {
                            hasInvalidSelection = true;
                        }

                    }
                    
                    if (!hasInvalidSelection)
                    {

                        totalSupplyCFM = totalSupplyCFM * 60; // CONVERTS FLOW TO CUBIC FEET PER MINUTE
                        totalReturnCFM = totalReturnCFM * 60; // CONVERTS FLOW TO CUBIC FEET PER MINUTE
                        totalExhaustCFM = totalExhaustCFM * 60; // CONVERTS FLOW TO CUBIC FEET PER MINUTE

                        netAirflow = (totalSupplyCFM - totalReturnCFM - totalExhaustCFM); // CALCULATES NET AIRFLOW
                        netAirflow = Math.Round(netAirflow, 0);

                        if (netAirflow > 0)
                        {
                            netAirflowMessage = "Net Airflow: +" + netAirflow + " CFM";
                        }
                        else
                        {
                            netAirflowMessage = "Net Airflow: " + netAirflow + " CFM";
                        }

                        totalFlowMessage = "Total Airflows: " + Environment.NewLine + Environment.NewLine + "Supply: " + totalSupplyCFM + " CFM" + Environment.NewLine + "Return: " + totalReturnCFM + " CFM" + Environment.NewLine + "Exhaust: " + totalExhaustCFM + " CFM" + Environment.NewLine + Environment.NewLine;

                        displayMessage = totalFlowMessage + netAirflowMessage;

                        //DISPLAYS TOTAL AIRFLOW MESSAGE
                        TaskDialog.Show("Calculate Airflow", displayMessage);
                        
                    }
                    else
                    {
                        TaskDialog.Show("Invalid Selection", "Make sure you only have air terminals selected.");
                    }

                }
            }
            catch (Exception e)
            {
                TaskDialog errorTaskDialog = new TaskDialog("Error");
                errorTaskDialog.MainInstruction = "An error occurred while calculating total airflows. Please screenshot and report this message.";
                errorTaskDialog.MainContent = "Error Message: " + e.Message;

                errorTaskDialog.Show();
                return Autodesk.Revit.UI.Result.Failed;
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }
}
