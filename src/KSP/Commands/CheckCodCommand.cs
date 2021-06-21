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
            string str = "";

            int procent = 1;

            //Element myElement = doc.GetElement(uiDoc.Selection.GetElementIds().First());
            //ElementType myElementType = doc.GetElement(myElement.GetTypeId()) as ElementType;

            IList<Element> eitems;
            

            List<Cat> bic = new List<Cat>();
            #region list bic

            bic.Add(new Cat(BuiltInCategory.OST_Ceilings, "Потолки"));
            bic.Add(new Cat(BuiltInCategory.OST_Roofs, "Крыши"));
            bic.Add(new Cat(BuiltInCategory.OST_Floors, "Перекрытия"));
            bic.Add(new Cat(BuiltInCategory.OST_Doors, "Двери"));
            bic.Add(new Cat(BuiltInCategory.OST_Windows, "Окна"));
            bic.Add(new Cat(BuiltInCategory.OST_Walls, "Стены"));
            bic.Add(new Cat(BuiltInCategory.OST_StairsRailing, "Ограждение"));
            bic.Add(new Cat(BuiltInCategory.OST_Stairs, "Лестницы"));
            bic.Add(new Cat(BuiltInCategory.OST_Columns, "Колонны"));
            bic.Add(new Cat(BuiltInCategory.OST_Ramps, "Пандус"));
            bic.Add(new Cat(BuiltInCategory.OST_CurtainWallMullions, "Импосты витража"));
            bic.Add(new Cat(BuiltInCategory.OST_CurtainWallPanels, "Панели витража"));
            bic.Add(new Cat(BuiltInCategory.OST_Furniture, "Мебель"));
            bic.Add(new Cat(BuiltInCategory.OST_FurnitureSystems, "Комплекты мебели"));
            bic.Add(new Cat(BuiltInCategory.OST_Roads, "Дорожки"));
            bic.Add(new Cat(BuiltInCategory.OST_Parking, "Парковка"));


            bic.Add(new Cat(BuiltInCategory.OST_StructuralColumns, "Несущие колонны"));
            bic.Add(new Cat(BuiltInCategory.OST_StructuralFraming, "Каркас несущий"));
            bic.Add(new Cat(BuiltInCategory.OST_StructuralFoundation, "Фундамент несущей конструкции"));
            bic.Add(new Cat(BuiltInCategory.OST_StructuralTruss, "Фермы"));
            bic.Add(new Cat(BuiltInCategory.OST_StructConnections, "Соединения несущих конструкций"));

            bic.Add(new Cat(BuiltInCategory.OST_PipeCurves, "Трубы"));
            bic.Add(new Cat(BuiltInCategory.OST_PipeAccessory, "Арматура трубопроводов"));
            bic.Add(new Cat(BuiltInCategory.OST_PipeFitting, "Соединительные детали трубопроводов"));
            bic.Add(new Cat(BuiltInCategory.OST_FlexPipeCurves, "Гибкие трубы"));
            bic.Add(new Cat(BuiltInCategory.OST_PlaceHolderPipes, "Трубопровод по осевой"));
            bic.Add(new Cat(BuiltInCategory.OST_FabricationPipework, "Трубы из базы данных производителя MEP"));
            bic.Add(new Cat(BuiltInCategory.OST_Sprinklers, "Спринклеры"));
            bic.Add(new Cat(BuiltInCategory.OST_DuctCurves, "Воздуховоды"));
            bic.Add(new Cat(BuiltInCategory.OST_DuctAccessory, "Арматура воздуховодов"));
            bic.Add(new Cat(BuiltInCategory.OST_DuctFitting, "Соединительные детали воздуховодов"));
            bic.Add(new Cat(BuiltInCategory.OST_DuctTerminal, "Воздухораспределители"));
            bic.Add(new Cat(BuiltInCategory.OST_FlexDuctCurves, "Гибкие воздуховоды"));
            bic.Add(new Cat(BuiltInCategory.OST_PlaceHolderDucts, "Воздуховоды по осевой"));
            bic.Add(new Cat(BuiltInCategory.OST_PlumbingFixtures, "Сантехнические приборы"));
            bic.Add(new Cat(BuiltInCategory.OST_MechanicalEquipment, "Оборудование"));
            bic.Add(new Cat(BuiltInCategory.OST_FabricationDuctwork, "Элементы воздуховодов из базы данных производителя MEP"));
            bic.Add(new Cat(BuiltInCategory.OST_FabricationHangers, "Подвески из базы данных производителя MEP"));


            bic.Add(new Cat(BuiltInCategory.OST_CableTray, "Кабельные лотки"));
            bic.Add(new Cat(BuiltInCategory.OST_CableTrayFitting, "Соединительные детали кабельных лотков"));
            bic.Add(new Cat(BuiltInCategory.OST_CableTrayRun, "Участки кабельного лотка"));
            bic.Add(new Cat(BuiltInCategory.OST_Conduit, "Короба"));
            bic.Add(new Cat(BuiltInCategory.OST_ConduitFitting, "Соединительные детали коробов"));
            bic.Add(new Cat(BuiltInCategory.OST_ConduitRun, "Участки короба"));
            bic.Add(new Cat(BuiltInCategory.OST_LightingDevices, "Осветительная аппаратура"));
            bic.Add(new Cat(BuiltInCategory.OST_FireAlarmDevices, "Системы пожарной сигнализации"));
            bic.Add(new Cat(BuiltInCategory.OST_DataDevices, "Датчики"));
            bic.Add(new Cat(BuiltInCategory.OST_CommunicationDevices, "Устройства связи"));
            bic.Add(new Cat(BuiltInCategory.OST_SecurityDevices, "Предохранительные устройства"));
            bic.Add(new Cat(BuiltInCategory.OST_NurseCallDevices, "Устройства вызова и оповещения"));
            bic.Add(new Cat(BuiltInCategory.OST_TelephoneDevices, "Телефонные устройства"));
            bic.Add(new Cat(BuiltInCategory.OST_Wire, "Провода"));
            bic.Add(new Cat(BuiltInCategory.OST_ElectricalCircuit, "Электрические цепи"));
            bic.Add(new Cat(BuiltInCategory.OST_SpecialityEquipment, "Специальное оборудование"));
            bic.Add(new Cat(BuiltInCategory.OST_LightingFixtures, "Осветительные приборы"));
            bic.Add(new Cat(BuiltInCategory.OST_ElectricalFixtures, "Силовые электроприборы"));
            bic.Add(new Cat(BuiltInCategory.OST_ElectricalEquipment, "Электрооборудование"));
            bic.Add(new Cat(BuiltInCategory.OST_Casework, "Шкафы"));


            bic.Add(new Cat(BuiltInCategory.OST_GenericModel, "Обобщенные модели"));
            bic.Add(new Cat(BuiltInCategory.OST_Entourage, "Антураж"));
            bic.Add(new Cat(BuiltInCategory.OST_Planting, "Озеленение"));


            #endregion

            foreach (var c in bic)
            {
                eitems = new FilteredElementCollector(doc).OfCategory(c.builtInCategory).WhereElementIsNotElementType().ToElements();
                if (eitems.Count != 0)
                {
                    temp = GetVsego(new Vsego(), eitems, doc);
                    str += "\n" + c.rusName + ": " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
                    vsego = GetVsego(vsego, eitems, doc);
                }
                
            }
            /*
            // АР, КР
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов стен: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов окон: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов дверей: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов потолка: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов перекрытий: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов несущих колонн: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов каркас несущий: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Columns).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов колонн: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов крыш: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            // ОВ, ВК
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов сантехприборов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов труб: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов соединителей труб: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов арматуры труб: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов воздуховодов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов гибких воздуховодов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов соединителей труб: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов арматуры воздуховодов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов мех оборудования: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sprinklers).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов спринклеров: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            // ЭОМ, СС
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов лотков: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTrayFitting).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов соединителей лотков: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов коробов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitFitting).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов соединителей коробов: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lights).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов светильников: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingDevices).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов осветительной аппаратуры: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CommunicationDevices).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов устройств связи: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов електрооборудования: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            eitems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SecurityDevices).WhereElementIsNotElementType().ToElements();
            temp = GetVsego(new Vsego(), eitems, doc);
            str += "\n" + "заполнено / не заполнено" + " типов приборов охранной сигнализации: " + temp.withCod.ToString() + " / " + (temp.all - temp.withCod).ToString();
            vsego = GetVsego(vsego, eitems, doc);
            */

            if (vsego.all != 0)
                procent = 100 * vsego.withCod / vsego.all;

            TaskDialog.Show("123", String.Format("всего типов {0}, заполнено {1}, / незаполнено {2},  процент {3}%" + str, vsego.all, vsego.withCod, vsego.all - vsego.withCod, procent));

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
    public class Cat
    {
        public BuiltInCategory builtInCategory;
        public string rusName;
        public Cat(BuiltInCategory c, string str)
        {
            builtInCategory = c;
            rusName = str;
        }
    }
}
