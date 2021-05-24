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
            Methods mth = new Methods();
            GetParameterTypeAndNameFromGUID param = new GetParameterTypeAndNameFromGUID();
            string noData = mth.noData;
            string noParameter = mth.noParameter;


            // М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели
            IList<ProjectInfo> pInfo = new FilteredElementCollector(doc).OfClass(typeof(ProjectInfo)).Cast<ProjectInfo>().ToList();
            if (pInfo != null)
            {
                string[] pSet = {
                                        param.Name(doc, "bb06ccda-e560-4727-8ffe-10d95da2f467", "МСК_Тип проекта"),
                                        param.Name(doc, "efd17af1-dc28-4f05-8dfb-49995f260aba", "МСК_Степень огнестойкости"),
                                        param.Name(doc, "85aff05b-aa70-4d0f-abf8-9c9c80fd8a8b", "МСК_Отметка нуля проекта"),
                                        param.Name(doc, "bbce2a0f-230d-4cd6-b330-4df16e65dc6c", "МСК_Отметка уровня земли"),
                                        param.Name(doc, "03641d89-8e7d-4f95-b90e-6626b94e601e", "МСК_Проектировщик"),
                                        param.Name(doc, "19fd1e38-47f3-4e04-b710-9b6b6ca27a5b", "МСК_Заказчик"),
                                        param.Name(doc, "fac88977-9987-4081-8883-c5d9405e3fb3", "МСК_Имя проекта"),
                                        param.Name(doc, "249f628c-1b80-4ed3-8962-a48a84c3c294", "МСК_Имя обьекта"),
                                        "Номер проекта",
                                        param.Name(doc, "3213141a-e8e6-4521-a8b4-0fa6ae8044f1", "МСК_Корпус"),
                                        param.Name(doc, "97311fa0-e904-4db1-9398-215889ef764b", "МСК_Номер секции"),
                                        param.Name(doc, "4f7092a4-9ac5-49dd-8250-e11d314e6202", "МСК_Количество секций")
                    };
                string[] ptSet = {
                                        param.Type(doc, "bb06ccda-e560-4727-8ffe-10d95da2f467", "МСК_Тип проекта"),
                                        param.Type(doc, "efd17af1-dc28-4f05-8dfb-49995f260aba", "МСК_Степень огнестойкости"),
                                        param.Type(doc, "85aff05b-aa70-4d0f-abf8-9c9c80fd8a8b", "МСК_Отметка нуля проекта"),
                                        param.Type(doc, "bbce2a0f-230d-4cd6-b330-4df16e65dc6c", "МСК_Отметка уровня земли"),
                                        param.Type(doc, "03641d89-8e7d-4f95-b90e-6626b94e601e", "МСК_Проектировщик"),
                                        param.Type(doc, "19fd1e38-47f3-4e04-b710-9b6b6ca27a5b", "МСК_Заказчик"),
                                        param.Type(doc, "fac88977-9987-4081-8883-c5d9405e3fb3", "МСК_Имя проекта"),
                                        param.Type(doc, "249f628c-1b80-4ed3-8962-a48a84c3c294", "МСК_Имя обьекта"),
                                        "NS",
                                        param.Type(doc, "3213141a-e8e6-4521-a8b4-0fa6ae8044f1", "МСК_Корпус"),
                                        param.Type(doc, "97311fa0-e904-4db1-9398-215889ef764b", "МСК_Номер секции"),
                                        param.Type(doc, "4f7092a4-9ac5-49dd-8250-e11d314e6202", "МСК_Количество секций")
                    };
                sb.Append("М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append(pSet[i]).Append("<" + ptSet[i] + ">").Append("\t");
                }
                sb.Append("\n");

                foreach (var el in pInfo)
                {
                    for (int i = 0; i < pSet.Count(); i++)
                    {
                        string result = mth.GetParameterValue(doc, el, pSet[i], noData, noParameter);
                        sb.Append(result).Append("\t");
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
                                        param.Name(doc, "6c93ae32-6fe2-48a6-ba39-b1522f8912f7", "МСК_Надземный"),
                                        param.Name(doc, "3bca158e-04af-4032-badc-7f9f3e61b836", "МСК_Базовый уровень"),
                                        param.Name(doc, "741ca870-26ce-460c-8c82-56c23ec53ebe", "МСК_Система пожаротушения"),
                                        param.Name(doc, "6dd43fcd-977a-498c-8269-d3a40ac3dce3",  "МСК_Наличие АУПТ"),
                                        param.Name(doc, "8ddb89d8-4145-480e-a813-c51afd9fd9c6",  "МСК_Уровень комфорта")
                    };
                string[] ptSet = {
                                        "NS",
                                        "NS",
                                        param.Type(doc, "6c93ae32-6fe2-48a6-ba39-b1522f8912f7", "МСК_Надземный"),
                                        param.Type(doc, "3bca158e-04af-4032-badc-7f9f3e61b836", "МСК_Базовый уровень"),
                                        param.Type(doc, "741ca870-26ce-460c-8c82-56c23ec53ebe", "МСК_Система пожаротушения"),
                                        param.Type(doc, "6dd43fcd-977a-498c-8269-d3a40ac3dce3",  "МСК_Наличие АУПТ"),
                                        param.Type(doc, "8ddb89d8-4145-480e-a813-c51afd9fd9c6",  "МСК_Уровень комфорта")
                    };
                sb.Append("М_АР_00_Таблица для заполнения параметров уровней\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append(pSet[i]).Append("\t");
                }
                sb.Append("\n");

                foreach (var lvl in levels)
                {
                    sb.Append(lvl.Name).Append("\t");
                    for (int i = 1; i < pSet.Count(); i++)
                    {
                        string result = mth.GetParameterValue(doc, lvl, pSet[i], noData, noParameter);
                        sb.Append(result).Append("<" + ptSet[i] + ">").Append("\t");
                        //if (i > 1)
                        //    MSKCounter(pSet[i], result, noData, noParameter);
                    }
                    sb.Append("\n");
                }
                sb.Append("\n");
            }


            TaskDialog.Show("Warning", sb.ToString());
                return Result.Succeeded;
        }

       
    }
}
