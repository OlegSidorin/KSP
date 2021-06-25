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
    class AddGPParameterVK : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var app = commandData.Application.Application;
            string path = Assembly.GetExecutingAssembly().Location;

            string ADGFOPFile = Path.GetDirectoryName(path) + "\\res\\ADG_FOP.txt";
            app.SharedParametersFilename = ADGFOPFile;

            CategorySet catSet = commandData.Application.Application.Create.NewCategorySet();
            Category catGP; 
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeInsulations);
            catSet.Insert(catGP);
            AddParams("GP_Марка", catSet, app, doc);
            AddParams("GP_Завод изготовитель", catSet, app, doc);
            AddParams("GP_Наименование", catSet, app, doc);

            catSet.Clear();
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory);
            catSet.Insert(catGP);
            AddParams("GP_Материал", catSet, app, doc);
            AddParams("GP_Толщина_изоляции", catSet, app, doc);

            catSet.Clear();
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            catSet.Insert(catGP);
            AddParams("GP_Масса", catSet, app, doc);
            AddParams("GP_Напор", catSet, app, doc);
            AddParams("GP_Номинальная мощность", catSet, app, doc);


            catSet.Clear();
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            catSet.Insert(catGP);
            catGP = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory);
            catSet.Insert(catGP);
            AddParams("GP_Потери_давления", catSet, app, doc);
            AddParams("GP_Рабочее давление", catSet, app, doc);
            
            

            return Result.Succeeded;
        }
        public void AddParams(string prm, CategorySet catSet, Application app, Document doc)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Add Shared Parameters");
                    DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();
                    DefinitionGroup sharedParameterGroup = sharedParameterFile.Groups.get_Item("ADG");
                    Definition sharedParameterDefinition = sharedParameterGroup.Definitions.get_Item(prm);
                    ExternalDefinition externalDefinition =
                        sharedParameterGroup.Definitions.get_Item(prm) as ExternalDefinition;
                    Guid guid = externalDefinition.GUID;
                    InstanceBinding newIB = app.Create.NewInstanceBinding(catSet);
                    doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.INVALID);
                    //SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);
                    // InternalDefinition def = sp.GetDefinition();
                    // def.SetAllowVaryBetweenGroups(doc, true);
                    t.Commit();
                }

            }
            catch (Exception e)
            {
                TaskDialog.Show("Warning", e.ToString());
            }
        }

    }


}
