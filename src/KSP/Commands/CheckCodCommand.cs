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
            //string str = "";
            Vsego vsego = new Vsego();
            Vsego temp = new Vsego();

            int procent = 1;

            //Element myElement = doc.GetElement(uiDoc.Selection.GetElementIds().First());
            //ElementType myElementType = doc.GetElement(myElement.GetTypeId()) as ElementType;

            IList<Element> eitems;

            // АР, КР
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Columns).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            // ОВ, ВК
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sprinklers).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            // ЭОМ, СС
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTrayFitting).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lights).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements();
            vsego = GetVsego(vsego, eitems, doc);


            if (vsego.all != 0)
                procent = 100 * vsego.withCod / vsego.all;

            TaskDialog.Show("123", String.Format("всего типов {0}, заполнено {1}, процент {2}%", vsego.all, vsego.withCod, procent));

            return Result.Succeeded;
        }


        class Vsego
        {
            public int all;
            public int withCod;
            public Vsego()
            {
                all = 0;
                withCod = 0;
            }

        };

        Vsego GetVsego(Vsego v, IList<Element> eitems, Document doc)
        {
            string st = "";
            Vsego vsego = v;
            IList<ElementType> eTypes = new List<ElementType>();
            foreach (var item in eitems)
            {
                try
                {
                    using (Transaction t = new Transaction(doc, "t!"))
                    {
                        t.Start();
                        var itemType = doc.GetElement(item.GetTypeId()) as ElementType;
                        if (eTypes.Where(xxx => xxx.Name == itemType.Name).Count() == 0)
                        {
                            eTypes.Add(itemType);
                            st = itemType.Name;
                            vsego.all += 1;
                            if (itemType.GetParameters("Код по классификатору")[0].AsString() != "")
                                vsego.withCod += 1;
                            //str += Environment.NewLine;
                        }
                        t.Commit();
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Catch", "Фигня, потому что: в " + st + Environment.NewLine + ex.Message);
                }
                //str += Environment.NewLine;
            }
            return vsego;
        }
        
    }
}
