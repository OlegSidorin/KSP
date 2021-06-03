namespace KSP
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Text;
    using System.Linq;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class CheckMSKKRwGCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            StringBuilder sb = new StringBuilder();
            DateTime dt = DateTime.Now;
            MyMethods m = new MyMethods();
            SharedParameterFromGUID param = new SharedParameterFromGUID();
            string noData = m.noData;
            string noParameter = m.noParameter;
            string noCategory = m.noCategory;


            m.countIfMSKCOdIs = 0;
            m.countAllMSKCod = 0;
            m.countIfParameterIs = 0;
            m.countAll = 0;
            m.readyOn = 0;
            string outputString = "";

			MyMSK myMSK = new MyMSK();

			int first, second;

			List<MyParameter> myParameters = m.AllParameters(doc);

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

	        // М_КР_Таблица для ДОБАВЛЕНИЯ параметров модели
            IList<ProjectInfo> pInfo = new FilteredElementCollector(doc).OfClass(typeof(ProjectInfo)).Cast<ProjectInfo>().ToList();
            if (pInfo != null)
            {
                string[] pSet = {
                    m.SharedParameterFromGUIDName("bb06ccda-e560-4727-8ffe-10d95da2f467", myParameters, "МСК_Тип проекта"),
                    m.SharedParameterFromGUIDName("efd17af1-dc28-4f05-8dfb-49995f260aba", myParameters, "МСК_Степень огнестойкости"),
                    m.SharedParameterFromGUIDName("85aff05b-aa70-4d0f-abf8-9c9c80fd8a8b", myParameters, "МСК_Отметка нуля проекта"),
                    m.SharedParameterFromGUIDName("bbce2a0f-230d-4cd6-b330-4df16e65dc6c", myParameters, "МСК_Отметка уровня земли"),
                    m.SharedParameterFromGUIDName("03641d89-8e7d-4f95-b90e-6626b94e601e", myParameters, "МСК_Проектировщик"),
                    m.SharedParameterFromGUIDName("19fd1e38-47f3-4e04-b710-9b6b6ca27a5b", myParameters, "МСК_Заказчик"),
                    m.SharedParameterFromGUIDName("fac88977-9987-4081-8883-c5d9405e3fb3", myParameters, "МСК_Имя проекта"),
                    m.SharedParameterFromGUIDName("249f628c-1b80-4ed3-8962-a48a84c3c294", myParameters, "МСК_Имя обьекта"),
                    "Номер проекта",
                    m.SharedParameterFromGUIDName("3213141a-e8e6-4521-a8b4-0fa6ae8044f1", myParameters, "МСК_Корпус"),
                    m.SharedParameterFromGUIDName("97311fa0-e904-4db1-9398-215889ef764b", myParameters, "МСК_Номер секции"),
                    m.SharedParameterFromGUIDName("4f7092a4-9ac5-49dd-8250-e11d314e6202", myParameters, "МСК_Количество секций")
                };

                sb.Append("М_КР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
                }
                sb.Append("\n");

                foreach (var el in pInfo)
                {
                    for (int i = 0; i < pSet.Count(); i++)
                    {
                        string result = m.GetParameterValue(doc, el, pSet[i]);
                        sb.Append(result).Append("\t");
                        m.MSKCounter(pSet[i], result);
                    }
                    sb.Append("\n");
                }
                sb.Append("\n");
            }

	        // М_КР_00_Таблица для заполнения параметров уровней
            IList<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            if (levels != null)
            {
                string[] pSet = {
                    "Имя",
                    "Фасад",
                    m.SharedParameterFromGUIDName("6c93ae32-6fe2-48a6-ba39-b1522f8912f7", myParameters, "МСК_Надземный"),
                    m.SharedParameterFromGUIDName("3bca158e-04af-4032-badc-7f9f3e61b836", myParameters, "МСК_Базовый уровень"),
                    m.SharedParameterFromGUIDName("741ca870-26ce-460c-8c82-56c23ec53ebe", myParameters, "МСК_Система пожаротушения"),
                    m.SharedParameterFromGUIDName("6dd43fcd-977a-498c-8269-d3a40ac3dce3", myParameters, "МСК_Наличие АУПТ"),
                    m.SharedParameterFromGUIDName("8ddb89d8-4145-480e-a813-c51afd9fd9c6", myParameters, "МСК_Уровень комфорта")
                };
                sb.Append("М_КР_00_Таблица для заполнения параметров уровней\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
                }
                sb.Append("\n");

                foreach (var lvl in levels)
                {
                    sb.Append(lvl.Name).Append("\t");
                    for (int i = 1; i < pSet.Count(); i++)
                    {
                        string result = m.GetParameterValue(doc, lvl, pSet[i]);
                        sb.Append(result).Append("\t");
                        if (i > 1)
                            m.MSKCounter(pSet[i], result);
                    }
                    sb.Append("\n");
                }
                sb.Append("\n");
            }

			IList<Element> elements = new List<Element>();


	        // М_КР_01.1.1_Таблица для заполнения параметров фундаментов (кроме плитных, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().Where(x => !x.Name.Contains("лита")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("0301c4f6-cfa3-44ae-97ee-9c917b9f308c", myParameters, "МСК_С подколонником"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_01.1.1_Таблица для заполнения параметров фундаментов (кроме плитных, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_01.2.1_Таблица для заполнения параметров фундаментов (плитных, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().Where(x => x.Name.Contains("лита")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("0301c4f6-cfa3-44ae-97ee-9c917b9f308c", myParameters, "МСК_С подколонником"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_01.2.1_Таблица для заполнения параметров фундаментов (плитных, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_01.3.1_Таблица для заполнения параметров свай (общая) (разные категории)
			elements = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => x.Category.Name != "Линии")
				.Where(x => x.Category.Name != "Элементы узлов")
				.Where(x => x.Category.Name == "свая" || x.Category.Name == "Свая")
				.ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					"Категория",
					"Семейство",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("0852d2c0-86ad-4a0c-bb43-03b2949d3a81", myParameters, "МСК_Защитный слой 2"),
					m.SharedParameterFromGUIDName("a4efaeb3-40d1-425b-8c3e-01da4dc09871", myParameters, "МСК_Способ погружения"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("321c229b-b3a5-4a47-93c1-caeedae09bf3", myParameters, "МСК_Плотность"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};

				sb.Append("М_КР_01.3.1_Таблица для заполнения параметров свай (общая) (разные категории)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParametersWCatNameWFamily(doc, pSet, elements));
			}


	        // М_КР_02.1_Таблица для заполнения параметров стен (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("0852d2c0-86ad-4a0c-bb43-03b2949d3a81", myParameters, "МСК_Защитный слой 2"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_02.1_Таблица для заполнения параметров стен (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}



	        // М_КР_03.1_Таблица для заполнения параметров армирования перекрытий (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_03.1_Таблица для заполнения параметров армирования перекрытий (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_04.1_Таблица для заполнения параметров колонн (стальных)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType()
				.Where(x => !x.Name.Contains("онолит")).Where(x => !x.Name.Contains("ж.б")).Where(x => !x.Name.Contains("Ж.б")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал")
				};
				sb.Append("М_КР_04.1_Таблица для заполнения параметров колонн (стальных)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_04.2.1_Таблица для заполнения параметров колонн (железобетонных, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType()
				.Where(x => (x.Name.Contains("онолит") || x.Name.Contains("ж.б") || x.Name.Contains("Ж.б"))).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("96588d70-4bff-4014-8058-048eeba2234b", myParameters, "МСК_Уровень ответственности"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_04.2.1_Таблица для заполнения параметров колонн (железобетонных, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_05.1_Таблица для заполнения параметров балок (стальных)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType()
				.Where(x => !x.Name.Contains("онолит")).Where(x => !x.Name.Contains("ж.б")).Where(x => !x.Name.Contains("Ж.б")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал")
				};
				sb.Append("М_КР_05.1_Таблица для заполнения параметров балок (стальных)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_05.2.1_Таблица для заполнения параметров балок (железобетонных, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType()
				.Where(x => (x.Name.Contains("онолит") || x.Name.Contains("ж.б") || x.Name.Contains("Ж.б"))).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("266a965f-d878-4a4a-baf5-f4a5cdcd69fd", myParameters, "МСК_Расход арматуры"),
					m.SharedParameterFromGUIDName("7b89dfac-4fea-43ab-b372-1d076b4624af", myParameters, "МСК_Защитный слой 1"),
					m.SharedParameterFromGUIDName("7c106ad8-d958-459c-99b7-46d092a178c7", myParameters, "МСК_Класс арматуры текст"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("334ccf69-6722-4ad2-95e1-a11c3e5ee322", myParameters, "МСК_Марка В"),
					m.SharedParameterFromGUIDName("b1d55c3a-7202-47fb-8593-a22cc05502e3", myParameters, "МСК_Водонепроницаемость W"),
					m.SharedParameterFromGUIDName("a6669024-a5bc-4323-bfaa-651492c9de6c", myParameters, "МСК_Морозостойкость F"),
					m.SharedParameterFromGUIDName("d7e61577-6b6f-4e94-9c31-af1cd8020bb7", myParameters, "МСК_Обозначение материал")
				};
				sb.Append("М_КР_05.2.1_Таблица для заполнения параметров балок (железобетонных, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_06.1.1_Таблица для заполнения параметров лестниц (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_КР_06.1.1_Таблица для заполнения параметров лестниц (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_06.2.1_Таблица для заполнения параметров лестниц (марши, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRuns).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_КР_06.2.1_Таблица для заполнения параметров лестниц (марши, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_07.1_Таблица для заполнения параметров пандусов (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ramps).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_КР_07.1_Таблица для заполнения параметров пандусов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_КР_08_Таблица для заполнения параметров сборок
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода",
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение")
				};
				sb.Append("М_КР_08_Таблица для заполнения параметров сборок\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));

			}


	        // М_КР(КЛ)_05_Таблица для классификации ограждений
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода"
				};
				sb.Append("М_КР(КЛ)_05_Таблица для классификации ограждений\n");
				sb.Append(m.RowHeader(pSet));
				foreach (Element el in elements)
				{
					try
					{
						FamilyInstance fi = el as FamilyInstance;
						if (null == fi.GetTypeId())
						{
							//
						}
						else
						{
							FamilySymbol fis = doc.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = m.GetParameterValue(doc, el, "МСК_Код по классификатору");
							m.MSKCounter("МСК_Код по классификатору", result);
							result += "\t" + myMSK.getMyMSK(result);
							sb.Append(result).Append("\t");
							sb.Append("\n");
						}

					}
					catch
					{
						// сюда придут FamilySimbol, они тоже категорией ost_mass
					}
				}
				sb.Append("\n");
			}


	        // М_КР(КЛ)_07_Таблица для классификации форм
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Mass).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода"
				};
				sb.Append("М_КР(КЛ)_07_Таблица для классификации форм\n");
				sb.Append(m.RowHeader(pSet));
				foreach (Element el in elements)
				{
					try
					{
						FamilyInstance fi = el as FamilyInstance;
						if (null == fi.GetTypeId())
						{
							//
						}
						else
						{
							FamilySymbol fis = doc.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = m.GetParameterValue(doc, el, "МСК_Код по классификатору");
							m.MSKCounter("МСК_Код по классификатору", result);
							result += "\t" + myMSK.getMyMSK(result);
							sb.Append(result).Append("\t");
							sb.Append("\n");
						}

					}
					catch
					{
						// сюда придут FamilySimbol, они тоже категорией ost_mass
					}
				}
				sb.Append("\n");
			}


	        // М_КР(КЛ)_08_Таблица для классификации крыш
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода"
				};
				sb.Append("М_КР(КЛ)_08_Таблица для классификации крыш\n");
				sb.Append(m.RowHeader(pSet));
				foreach (Element el in elements)
				{
					try
					{
						FamilyInstance fi = el as FamilyInstance;
						if (null == fi.GetTypeId())
						{
							//
						}
						else
						{
							FamilySymbol fis = doc.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = m.GetParameterValue(doc, el, "МСК_Код по классификатору");
							m.MSKCounter("МСК_Код по классификатору", result);
							result += "\t" + myMSK.getMyMSK(result);
							sb.Append(result).Append("\t");
							sb.Append("\n");
						}

					}
					catch
					{
						// сюда придут FamilySimbol, они тоже категорией ost_mass
					}
				}
				sb.Append("\n");
			}


	        // М_КР(КЛ)_09_Таблица для классификации остальных категорий
			elements = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => x.Category.Name != "Линии")
				.Where(x => x.Category.Name != "Элементы узлов")
				.Where(x => x.Category.Name != "Каркас несущий")
				.Where(x => x.Category.Name != "Несущие колонны")
				.Where(x => x.Category.Name != "Фундамент несущей конструкции")
				.Where(x => !x.Name.Contains("КПСП_Отверстие"))
				.ToList();
			List<string> elementsInStrings = new List<string>();
			//            IEnumerable<Element> unicElements = elements.Distinct();
			if (elements != null)
			{
				string[] pSet = {
					"Категория",
					"Семейство",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Значение Кода"
				};

				sb.Append("М_КР(КЛ)_09_Таблица для классификации остальных категорий\n");
				sb.Append(m.RowHeader(pSet));
				// сперва закинем элементы массива в строки, для последующей работы с ними
				foreach (Element el in elements)
				{
					try
					{
						FamilyInstance fi = el as FamilyInstance;

						if (null == fi.GetTypeId())
						{
							//
						}
						else
						{
							string str = "";
							FamilySymbol fis = doc.GetElement(fi.GetTypeId()) as FamilySymbol;
							//							sb.Append(fi.Category.Name).Append("\t");
							//							sb.Append(fis.FamilyName).Append("\t");
							//							sb.Append(el.Name).Append("\t");
							//							sb.Append("?\t");
							//							sb.Append(getParameterValue(doc,el, m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"), noData, noParameter)).Append("\t");
							//							sb.Append("\n");
							string result = m.GetParameterValue(doc, el, "МСК_Код по классификатору");
							m.MSKCounter("МСК_Код по классификатору", result);
							result += "\t" + myMSK.getMyMSK(result);
							str += fi.Category.Name + "\t" + fis.FamilyName + "\t" + result + "\t";
							elementsInStrings.Add(str);
						}

					}
					catch
					{
						// 
					}
				}

				// теперь найдем уникальные эелементы
				IEnumerable<string> distinctElements = elementsInStrings.Distinct().OrderBy(x => x);
				foreach (string str in distinctElements)
				{
					if (str.Contains(noParameter))
						m.countAll += 1;

					if (str.Contains(noData))
						m.countAll += 1;

					sb.Append(str).Append("\n");
				}
			}


			#region Заголовок файла
			StringBuilder sbStart = new StringBuilder();
			sbStart.Append("Отчет по модели: ").Append(m.CropFileName(doc.Title)).Append("\n");
			sbStart.Append("Составлен: ").Append(String.Format("{0:dd.MM.yyyy}г.", dt)).Append("\n");
			sbStart.Append("Заполнено параметров: " + m.countIfParameterIs + "\n");
			sbStart.Append("Всего параметров: " + m.countAll + "\n");
			if (m.countAll > 0)
				m.readyOn = (int)(Math.Round((double)m.countIfParameterIs / (double)m.countAll * 100, 0));
			sbStart.Append(String.Format("Информационное наполнение: {0}%\n", m.readyOn));
			sbStart.Append("\n");
			if (m.countAllMSKCod > 0)
			{
				first = (int)(Math.Round((double)m.countIfMSKCOdIs / (double)m.countAllMSKCod * 100, 0));
				outputString += "Заполнен классификатор: " + first.ToString() + "%\n";
			}
			else
			{
				first = 0;
				outputString += "Заполнен классификатор: 0%\n";
			}
			if ((m.countAll > 0) || (m.countAll != m.countAllMSKCod))
			{
				second = (int)(Math.Round((double)(m.countIfParameterIs - m.countIfMSKCOdIs) / (double)(m.countAll - m.countAllMSKCod) * 100, 0));
				outputString += "Заполнения параметров элементов модели: " + second.ToString() + "%\n";
			}
			else
			{
				second = 0;
				outputString += "Заполнения параметров элементов модели: 0%\n";
			}


			sbStart.Append(outputString).Append("\n");
			#endregion

			StringBuilder sbResult = new StringBuilder();

			sbResult.Append(sbStart.ToString()).Append(sb.ToString());

			#region Завершение и Условные обозначения
			StringBuilder sbEnd = new StringBuilder();
			sbEnd.Append("\n\nУсловные обозначения: \n");
			sbEnd.Append("!!(_) \t - не добавлен в модель общий параметр\n");
			sbEnd.Append("??(_)" + "\t - есть параметр с таким именем, но он не из ФОП или из другого ФОП\n");
			sbEnd.Append(noData + "\t - значение параметра не заполнено\n");
			sbEnd.Append(noParameter + "\t - не добавлен в модель общий параметр в экземпляре\n");
			sbEnd.Append(noCategory + "\t - не назначена категория, связанная с этим общим параметром\n");
			#endregion

			sbResult.Append(sbEnd.ToString());

			if (!Directory.Exists(m.workingDir))
				Directory.CreateDirectory(m.workingDir);

			string filePathToExcel = m.workingDir + m.CropFileName(doc.Title) + String.Format(" (МСК_Код на {0:00}% ост на {1:00}%)", first, second) + ".xlsx";
			//string filePathToTxt = CropFileName(document.Title) + String.Format(" (МСК на {0:00}пр)", readyOn);
			string excelSheet = m.CropFileName(doc.Title);
			// writeToFile(dirPath, filePathToTxt, sb2.ToString());
			try
			{
				m.WriteToExcel(filePathToExcel, excelSheet, sbResult.ToString(), noData, noParameter, 37);
			}
			catch
			{
				var rnd = new Random();
				filePathToExcel = m.workingDir + m.CropFileName(doc.Title) + String.Format(" (МСК_Код на {0:00}% ост на {1:00}%)", first, second) + "-v" + rnd.Next(99).ToString() + ".xlsx";
				m.WriteToExcel(filePathToExcel, excelSheet, sbResult.ToString(), noData, noParameter, 37);
			}


			TaskDialog.Show("Final", "Готово!");
			m.OpenFolder(m.workingDir);
			return Result.Succeeded;
		}
    }
}
