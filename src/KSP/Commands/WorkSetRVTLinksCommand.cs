namespace KSP
{
    using System;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class WorkSetRVTLinksCommand : IExternalCommand
    {
        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            if (doc.IsWorkshared == false)
            {
                TaskDialog.Show("Worksets", "Документ не для совместной работы");
            }
            if (doc.IsWorkshared == true)
            {
                CreateWorkset(doc, "Связь АР");
                CreateWorkset(doc, "Связь КР");
                CreateWorkset(doc, "Связь ИНЖ");
                int arWsId = 1;
                int krWsId = 1;
                int inzhWsId = 1;
                IList<Workset> worksetList = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets();
                foreach (Workset workset in worksetList)
                {
                    if (workset.Name.Contains("Связь АР"))
                    {
                        arWsId = workset.Id.IntegerValue;
                    }
                    if (workset.Name.Contains("Связь КР"))
                    {
                        krWsId = workset.Id.IntegerValue;
                    }
                    if (workset.Name.Contains("Связь ИНЖ"))
                    {
                        inzhWsId = workset.Id.IntegerValue;
                    }
                }

                FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType();
                foreach (Element item in linkCollector)
                {
                    RevitLinkInstance rvtLink = item as RevitLinkInstance;
                    using (Transaction tx = new Transaction(doc, "Change Workset"))
                    {
                        tx.Start();
                        try
                        {
                            rvtLink.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(inzhWsId);
                            if (rvtLink.Name.Contains("AR") || rvtLink.Name.Contains("АР") || rvtLink.Name.Contains("AI") || rvtLink.Name.Contains("АИ"))
                                rvtLink.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(arWsId);
                            if (rvtLink.Name.Contains("KR") || rvtLink.Name.Contains("КР"))
                                rvtLink.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(krWsId);
                        }
                        catch
                        { }

                        tx.Commit();
                    }
                }
            }
            return Result.Succeeded;
        }
		private Workset CreateWorkset(Document document, String wsname)
		{
			Workset newWorkset = null;
			// Worksets can only be created in a document with worksharing enabled
			if (document.IsWorkshared)
			{
				string worksetName = wsname;
				// Workset name must not be in use by another workset
				if (WorksetTable.IsWorksetNameUnique(document, worksetName))
				{
					using (Transaction worksetTransaction = new Transaction(document, "Set preview view id"))
					{
						worksetTransaction.Start();
						newWorkset = Workset.Create(document, worksetName);
						worksetTransaction.Commit();
					}
				}
			}

			return newWorkset;
		}

	}

}
