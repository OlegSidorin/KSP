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
    class CheckCodCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Element myElement = doc.GetElement(uiDoc.Selection.GetElementIds().First());
            ElementType myElementType = doc.GetElement(myElement.GetTypeId()) as ElementType;

            try
            {
                using (Transaction tr = new Transaction(doc, "trans"))
                {
                    tr.Start();
                    if (myElement.LookupParameter("Комментарии") != null)
                        myElement.GetParameters("Комментарии")[0].Set("Привет!");
                    if (myElementType.LookupParameter("Код по классификатору") != null)
                        myElementType.GetParameters("Код по классификатору")[0].Set("101");
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

            return Result.Succeeded;
        }
    }
}
