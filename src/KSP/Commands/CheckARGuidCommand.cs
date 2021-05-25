namespace KSP
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Text;
    using System.Linq;
    //using static KSP.CommonMethods;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class CheckARGuidCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            StringBuilder sb = new StringBuilder();
            DateTime dt = DateTime.Now;
            MyMethods myMethods = new MyMethods();
            SharedParameterFromGUID param = new SharedParameterFromGUID();
            string noData = myMethods.noData;
            string noParameter = myMethods.noParameter;

            List<MyParameter> myParameters = myMethods.AllParameters(doc);

            #region перевести в словарь ключ - значение
            //Dictionary<string, MyParameter> onlyShared = new Dictionary<string, MyParameter>();
            //foreach (var e in myParameters)
            //{
            //    string s = "";
            //    try
            //    {
            //        if (e.GuidValue != "")
            //        {
            //            s = e.GuidValue + "\n(" + e.Name + ")";
            //            onlyShared.Add(e.GuidValue, e);
            //        }

            //    }
            //    catch (Exception x)
            //    {
            //        //TaskDialog.Show("1", s + "\n : " + x.Message.ToString());
            //    }
            //}
            #endregion

            #region показать все параметры, внедренные в проект
            //foreach (var e in myParameters)
            //{
            //    sb = sb.Append("\n" + "<" + e.GuidValue + ">" + e.Name + "-" + e.isShared + "-" + e.isInstance);
            //    //sb = sb.Append("\n<" + e.Key + ">").Append("(" + e.Value.Name + ")").Append("-" + e.Value.isShared).Append("-" + e.Value.isInstance);
            //}

            //TaskDialog.Show("Warning", sb.ToString());
            #endregion

            // М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели
            IList<ProjectInfo> pInfo = new FilteredElementCollector(doc).OfClass(typeof(ProjectInfo)).Cast<ProjectInfo>().ToList();
            MyParameter value = new MyParameter("", "", false, false);
            if (pInfo != null)
            {
                string[] pSet = {
                    myMethods.SharedParameterFromGUIDName("bb06ccda-e560-4727-8ffe-10d95da2f467", myParameters, "МСК_Тип проекта"),
                    myMethods.SharedParameterFromGUIDName("efd17af1-dc28-4f05-8dfb-49995f260aba", myParameters, "МСК_Степень огнестойкости"),
                    myMethods.SharedParameterFromGUIDName("85aff05b-aa70-4d0f-abf8-9c9c80fd8a8b", myParameters, "МСК_Отметка нуля проекта"),
                    myMethods.SharedParameterFromGUIDName("bbce2a0f-230d-4cd6-b330-4df16e65dc6c", myParameters, "МСК_Отметка уровня земли"),
                    myMethods.SharedParameterFromGUIDName("03641d89-8e7d-4f95-b90e-6626b94e601e", myParameters, "МСК_Проектировщик"),
                    myMethods.SharedParameterFromGUIDName("19fd1e38-47f3-4e04-b710-9b6b6ca27a5b", myParameters, "МСК_Заказчик"),
                    myMethods.SharedParameterFromGUIDName("fac88977-9987-4081-8883-c5d9405e3fb3", myParameters, "МСК_Имя проекта"),
                    myMethods.SharedParameterFromGUIDName("249f628c-1b80-4ed3-8962-a48a84c3c294", myParameters, "МСК_Имя обьекта"),
                    "Номер проекта",
                    myMethods.SharedParameterFromGUIDName("3213141a-e8e6-4521-a8b4-0fa6ae8044f1", myParameters, "МСК_Корпус"),
                    myMethods.SharedParameterFromGUIDName("97311fa0-e904-4db1-9398-215889ef764b", myParameters, "МСК_Номер секции"),
                    myMethods.SharedParameterFromGUIDName("4f7092a4-9ac5-49dd-8250-e11d314e6202", myParameters, "МСК_Количество секций")
                };

                sb.Append("М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append("\n").Append(pSet[i]).Append("\t");
                }
                sb.Append("\n");

                foreach (var el in pInfo)
                {
                    for (int i = 0; i < pSet.Count(); i++)
                    {
                        string result = myMethods.GetParameterValue(doc, el, pSet[i], noData, noParameter);
                        sb.Append("\n").Append(result).Append("\t");
                        //MSKCounter(pSet[i], result, noData, noParameter);
                    }
                    sb.Append("\n");
                }
                sb.Append("\n");

            }


            //М_АР_00_Таблица для заполнения параметров уровней
            IList<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            if (levels != null)
            {
                string[] pSet = {
                    "Имя",
                    "Фасад",
                    myMethods.SharedParameterFromGUIDName("6c93ae32-6fe2-48a6-ba39-b1522f8912f7", myParameters, "МСК_Надземный"),
                    myMethods.SharedParameterFromGUIDName("3bca158e-04af-4032-badc-7f9f3e61b836", myParameters, "МСК_Базовый уровень"),
                    myMethods.SharedParameterFromGUIDName("741ca870-26ce-460c-8c82-56c23ec53ebe", myParameters, "МСК_Система пожаротушения"),
                    myMethods.SharedParameterFromGUIDName("6dd43fcd-977a-498c-8269-d3a40ac3dce3", myParameters, "МСК_Наличие АУПТ"),
                    myMethods.SharedParameterFromGUIDName("8ddb89d8-4145-480e-a813-c51afd9fd9c6", myParameters, "МСК_Уровень комфорта")
                };

                sb.Append("М_АР_00_Таблица для заполнения параметров уровней\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append("\n").Append(pSet[i]).Append("\t");
                }
                sb.Append("\n");

                foreach (var lvl in levels)
                {
                    sb.Append(lvl.Name).Append("\t");
                    for (int i = 1; i < pSet.Count(); i++)
                    {
                        string result = myMethods.GetParameterValue(doc, lvl, pSet[i], noData, noParameter);
                        sb.Append("\n").Append(result).Append("\t");
                        //if (i > 1)
                        //    MSKCounter(pSet[i], result, noData, noParameter);
                    }
                    sb.Append("\n");
                }
                sb.Append("\n");
            }




            TaskDialog.Show("Final", sb.ToString());
            return Result.Succeeded;
        }

       
    }
}
