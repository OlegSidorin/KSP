namespace KSP
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.ApplicationServices;
    using System.Text;
    using System.Linq;
    using System.Reflection;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class AddADGClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string path = Assembly.GetExecutingAssembly().Location;
            string assemblyTablePath = Path.GetDirectoryName(path) + "\\res\\2021-05-30_Классификатор.txt";
            var resourceReference = ExternalResourceReference.CreateLocalResource(doc, ExternalResourceTypes.BuiltInExternalResourceTypes.AssemblyCodeTable,
                ModelPathUtils.ConvertUserVisiblePathToModelPath(assemblyTablePath), PathType.Absolute) as ExternalResourceReference;

            using (Transaction t = new Transaction(doc, "load classificator"))
            {
                t.Start();
                AssemblyCodeTable.GetAssemblyCodeTable(doc).LoadFrom(resourceReference, new KeyBasedTreeEntriesLoadResults());
                t.Commit();
            }

            return Result.Succeeded;
        }
        

    }


}
