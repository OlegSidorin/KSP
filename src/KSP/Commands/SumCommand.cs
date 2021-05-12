namespace KSP
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using System.Threading.Tasks;
	using Autodesk.Revit.UI;
	using Autodesk.Revit.DB;

	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
	class SumCommand : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIDocument uidoc = commandData.Application.ActiveUIDocument;
			Document doc = uidoc.Document;
			Element myElement = null;

			//Element myElement = doc.GetElement(uidoc.Selection.GetElementIds().First());
			//ElementType myElementType = doc.GetElement(myElement.GetTypeId()) as ElementType;


			IList<ElementId> selected = uidoc.Selection.GetElementIds().ToList();


			Double summa = 0;
			int vsego = 0;

			foreach (ElementId elmentid in selected)
			{
				myElement = doc.GetElement(elmentid);
				try
				{
					using (Transaction tr = new Transaction(doc, "trans"))
					{
						tr.Start();
						if (myElement.LookupParameter("Объем") != null)
						{
							summa = summa + myElement.GetParameters("Объем")[0].AsDouble() * 0.3048 * 0.3048 * 0.3048;
							vsego = vsego + 1;
						};
						tr.Commit();
					}
				}

				#region catch and finally
				catch (Exception ex)
				{
					TaskDialog.Show("Catch", "Фигня, потому что:" + Environment.NewLine + ex.Message);
				}
				finally
				{

				}
				#endregion
			}

			TaskDialog.Show("Summa", String.Format("Общий объем: {0}м3, элементов: {1}", Math.Round(summa, 2).ToString(), vsego.ToString()));

			return Result.Succeeded;
		}
	}
}