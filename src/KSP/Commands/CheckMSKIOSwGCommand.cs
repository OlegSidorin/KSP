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
	class CheckMSKIOSwGCommand : IExternalCommand
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
            foreach (var e in myParameters)
            {
                sb = sb.Append("\n" + "<" + e.GuidValue + ">" + e.Name + "-" + e.isShared + "-" + e.isInstance);
                //sb = sb.Append("\n<" + e.Key + ">").Append("(" + e.Value.Name + ")").Append("-" + e.Value.isShared).Append("-" + e.Value.isInstance);
            }

            //TaskDialog.Show("Warning", sb.ToString());
			#endregion

			string str12 = "";
			BindingMap bindings = doc.ParameterBindings;
			int n = bindings.Size;
			if (0 < n)
			{
				DefinitionBindingMapIterator it = bindings.ForwardIterator();
				while (it.MoveNext())
				{
					Definition d = it.Key as Definition;
					Binding b = it.Current as Binding;
					str12 += "\n" + d.Name + "<" + d.ParameterType + ">" + "-" + b.ToString() + "-" + b.GetType().Name;
				}
			}

			TaskDialog.Show("Warning", str12);

			// М_ИОС_Таблица для ДОБАВЛЕНИЯ параметров модели
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

				sb.Append("М_ИОС_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
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

	        // М_ИОС_00 Таблица для заполнения параметров уровней
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
				sb.Append("М_ИОС_00_Таблица для заполнения параметров уровней\n");
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

	        // М_ИОС_01.1_Таблица для заполнения параметров воздуховодов (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("00a4a4bb-3575-45e4-8d3c-ddf9be915d27", myParameters, "МСК_Форма воздуховода"),
					"Рабочее давление",
					"Диаметр",
					"Ширина",
					"Высота",
					m.SharedParameterFromGUIDName("e928970f-9364-478f-8deb-0cd53d9e0460", myParameters, "МСК_Заводская изоляция"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					m.SharedParameterFromGUIDName("381b467b-3518-42bb-b183-35169c9bdfb3", myParameters, "МСК_Толщина стенки"),
					m.SharedParameterFromGUIDName("e5d67137-a4c3-4735-8d0f-2827b321bdf2", myParameters, "МСК_Огнестойкость EI"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c4d7c302-4962-46e5-aa62-f2885e5b5e2b", myParameters, "МСК_Признак ЭЭ"),
					m.SharedParameterFromGUIDName("4fe8448b-9d89-4a81-8e91-59ea6ce5e292", myParameters, "МСК_Материал теплоизоляции")
				};
				sb.Append("М_ИОС_01.1_Таблица для заполнения параметров воздуховодов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_02.1_Таблица для заполнения параметров фитингов возудоводов (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("00a4a4bb-3575-45e4-8d3c-ddf9be915d27", myParameters, "МСК_Форма воздуховода"),
					"МСК_Коэффициент шероховатости", // нет такого в ФОП
					m.SharedParameterFromGUIDName("e928970f-9364-478f-8deb-0cd53d9e0460", myParameters, "МСК_Заводская изоляция"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					m.SharedParameterFromGUIDName("381b467b-3518-42bb-b183-35169c9bdfb3", myParameters, "МСК_Толщина стенки"),
					m.SharedParameterFromGUIDName("e5d67137-a4c3-4735-8d0f-2827b321bdf2", myParameters, "МСК_Огнестойкость EI"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c4d7c302-4962-46e5-aa62-f2885e5b5e2b", myParameters, "МСК_Признак ЭЭ"),
					m.SharedParameterFromGUIDName("4fe8448b-9d89-4a81-8e91-59ea6ce5e292", myParameters, "МСК_Материал теплоизоляции")
				};
				sb.Append("М_ИОС_02.1_Таблица для заполнения параметров фитингов возудоводов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_03.1_Таблица для заполнения параметров воздухораспредлительных устройств (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("3b692944-1516-4825-8a7a-7d995b9af60b", myParameters, "МСК_Форма устройства"),
					m.SharedParameterFromGUIDName("b9d989f7-25ef-48a2-8c8d-3fc08ca369f0", myParameters, "МСК_Тип решетки"),
					m.SharedParameterFromGUIDName("c4d7c302-4962-46e5-aa62-f2885e5b5e2b", myParameters, "МСК_Признак ЭЭ"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("9b3dbd60-5be3-4842-9dbe-cd644ef5f9e8", myParameters, "МСК_Описание"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть")
				};
				sb.Append("М_ИОС_03.1_Таблица для заполнения параметров воздухораспредлительных устройств (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_04.1_Таблица для заполнения параметров труб (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					"МСК_Толщина труб", // нет в ФОП
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c4d7c302-4962-46e5-aa62-f2885e5b5e2b", myParameters, "МСК_Признак ЭЭ"),
					m.SharedParameterFromGUIDName("4fe8448b-9d89-4a81-8e91-59ea6ce5e292", myParameters, "МСК_Материал теплоизоляции")
				};
				sb.Append("М_ИОС_04.1_Таблица для заполнения параметров труб (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_05.1_Таблица для заполнения параметров фитингов труб (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("ca816cbc-4f60-47c8-b0fe-14eddcc32b16", myParameters, "МСК_Потеря давления жидкости"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					"МСК_Толщина труб", // нет в ФОП
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c4d7c302-4962-46e5-aa62-f2885e5b5e2b", myParameters, "МСК_Признак ЭЭ"),
					m.SharedParameterFromGUIDName("4fe8448b-9d89-4a81-8e91-59ea6ce5e292", myParameters, "МСК_Материал теплоизоляции")
				};
				sb.Append("М_ИОС_05.1_Таблица для заполнения параметров фитингов труб (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_06.1_Таблица для заполнения параметров отопительных приборов (оборудование) (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().Where(x => !x.Name.Contains("асос")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("eb2bda2f-f2cc-4e82-9dbb-37a902a93550", myParameters, "МСК_Тип расположения"),
					m.SharedParameterFromGUIDName("946c4e27-a56c-422d-999c-778a150b950e", myParameters, "МСК_Масса"),
					m.SharedParameterFromGUIDName("5f949eba-755b-4eeb-970b-1e1ac1908ec2", myParameters, "МСК_Удельная теплоемкость"),
					m.SharedParameterFromGUIDName("be7d2b1b-1916-428f-87f0-d9ee8d4f1efe", myParameters, "МСК_Тепловая мощность"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("4ae57b46-9303-4c39-8b84-eda83fc303db", myParameters, "МСК_Тип отоп_приборов"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("87ce1509-068e-400f-afab-75df889463c7", myParameters, "МСК_Материал наименование"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть")
				};
				sb.Append("М_ИОС_06.1_Таблица для заполнения параметров отопительных приборов (оборудование) (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_ИОС_07.1_Таблица для заполнения параметров распределительных устройств (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("7bae0e4a-a125-4818-973a-00fd56bf853d", myParameters, "МСК_Полная мощность"),
					m.SharedParameterFromGUIDName("1d29e937-60db-46f8-9f4f-9fedd8d152d2", myParameters, "МСК_Мощность в режиме тушения пожара"),
					m.SharedParameterFromGUIDName("9726f0ae-c8af-4ada-acf3-53905f395bbd", myParameters, "МСК_Класс защиты"),
					m.SharedParameterFromGUIDName("e815c9f3-a9ea-46ea-bda1-4142888b481e", myParameters, "МСК_Степень защиты от удара"),
					m.SharedParameterFromGUIDName("f479c656-5e96-46db-b764-de312da43f69", myParameters, "МСК_Основное устройство"),
					m.SharedParameterFromGUIDName("b52ad9d6-3610-492d-9f18-5f8e3c592bb6", myParameters, "МСК_Уровень квалификации"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение")
				};
				sb.Append("М_ИОС_07.1_Таблица для заполнения параметров распределительных устройств (электрооборудование) (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        // М_ИОС_08.1_Таблица для заполнения параметров осветительных приборов (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingFixtures).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("7bae0e4a-a125-4818-973a-00fd56bf853d", myParameters, "МСК_Полная мощность"),
					m.SharedParameterFromGUIDName("1d29e937-60db-46f8-9f4f-9fedd8d152d2", myParameters, "МСК_Мощность в режиме тушения пожара"),
					m.SharedParameterFromGUIDName("ffad6d6f-a441-4a89-aafe-d20c29181fe6", myParameters, "МСК_Признак заземления"),
					m.SharedParameterFromGUIDName("9726f0ae-c8af-4ada-acf3-53905f395bbd", myParameters, "МСК_Класс защиты"),
					m.SharedParameterFromGUIDName("e815c9f3-a9ea-46ea-bda1-4142888b481e", myParameters, "МСК_Степень защиты от удара"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("585a869f-3470-4f86-a969-11c22f7770e4", myParameters, "МСК_Продолжительность автономной работы")
				};
				sb.Append("М_ИОС_08.1_Таблица для заполнения параметров осветительных приборов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_09.1_Таблица для заполнения параметров кабельных лотков (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("2cbe3df6-5763-4bdf-81a6-bbdd146dbd4f", myParameters, "МСК_Тип конструкции"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("34060eea-e06a-4cd2-8dc5-0e12dbeb7695", myParameters, "МСК_Материал перегородки лотка"),
					m.SharedParameterFromGUIDName("62a33040-534c-4a1e-b2c9-6931b7eb7158", myParameters, "МСК_Коррозионная защита")
				};
				sb.Append("М_ИОС_09.1_Таблица для заполнения параметров кабельных лотков (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_10.1_Таблица для заполнения параметров фитингов кабельных лотков (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTrayFitting).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("2cbe3df6-5763-4bdf-81a6-bbdd146dbd4f", myParameters, "МСК_Тип конструкции"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("34060eea-e06a-4cd2-8dc5-0e12dbeb7695", myParameters, "МСК_Материал перегородки лотка"),
					m.SharedParameterFromGUIDName("62a33040-534c-4a1e-b2c9-6931b7eb7158", myParameters, "МСК_Коррозионная защита")
				};
				sb.Append("М_ИОС_10.1_Таблица для заполнения параметров фитингов кабельных лотков (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС_11.1_Таблица заполнения параметров насосов (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().Where(x => x.Name.Contains("асос")).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("7bae0e4a-a125-4818-973a-00fd56bf853d", myParameters, "МСК_Полная мощность"),
					m.SharedParameterFromGUIDName("1d29e937-60db-46f8-9f4f-9fedd8d152d2", myParameters, "МСК_Мощность в режиме тушения пожара"),
					m.SharedParameterFromGUIDName("9726f0ae-c8af-4ada-acf3-53905f395bbd", myParameters, "МСК_Класс защиты"),
					m.SharedParameterFromGUIDName("e815c9f3-a9ea-46ea-bda1-4142888b481e", myParameters, "МСК_Степень защиты от удара"),
					m.SharedParameterFromGUIDName("fc98eb11-1f2e-4682-8a41-b16c9f8a1b82", myParameters, "МСК_Производительность"),
					m.SharedParameterFromGUIDName("d5d3c594-bfbf-486c-936f-903ce6f7e9dd", myParameters, "МСК_Напор"),
					m.SharedParameterFromGUIDName("01cac3a8-85ec-4e54-a3dc-19903e71d4b0", myParameters, "МСК_Скорость вращения"),
					m.SharedParameterFromGUIDName("afaf7fef-3f82-4916-9a3f-383c5c9f2d5d", myParameters, "МСК_Диаметр рабочего колеса"),
					m.SharedParameterFromGUIDName("eb2bda2f-f2cc-4e82-9dbb-37a902a93550", myParameters, "МСК_Тип расположения"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("e19f8a6a-09c8-4e6c-9649-0919ae5f6ad8", myParameters, "МСК_Тип устройства")
				};
				sb.Append("М_ИОС_11.1_Таблица заполнения параметров насосов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС(КЛ)_01_Таблица для классификации пространств 
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Имя",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_ИОС(КЛ)_01_Таблица для классификации пространств\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_ИОС(КЛ)_02_Таблица для классификации сборок
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
				};
				sb.Append("М_ИОС(КЛ)_02_Таблица для классификации сборок\n");
				sb.Append(m.RowHeader(pSet));
				foreach (Element el in elements)
				{
					AssemblyInstance ai = el as AssemblyInstance;
					try
					{
						if (null == ai.GetTypeId())
						{
							//
						}
						else
						{
							AssemblyType aiType = doc.GetElement(ai.GetTypeId()) as AssemblyType;
							sb.Append(aiType.FamilyName).Append("\t");
						}

					}
					catch
					{
						sb.Append(noParameter).Append("\t");
					}
					sb.Append(ai.AssemblyTypeName).Append("\t");
					sb.Append("?\t");
					sb.Append(m.GetParameterValue(doc, el, "МСК_Код по классификатору")).Append("\t");
					sb.Append("\n");
				}

				sb.Append("\n");

			}

	        // М_ИОС(КЛ)_03_Таблица для классификации форм
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Mass).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
				};
				sb.Append("М_ИОС(КЛ)_03_Таблица для классификации форм\n");
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
							sb.Append(m.GetParameterValue(doc, el, "МСК_Код по классификатору")).Append("\t");
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

	        // М_ИОС(КЛ)_04_Таблица для классификации остальных категорий
			elements = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => !x.Name.Contains("КПСП_Отверстие"))
				.ToList();
			List<string> elementsInStrings = new List<string>();
			//            IEnumerable<Element> unicElements = elements.Distinct();
			if (elements != null)
			{
				string[] pSet = {
					"Категория",
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
				};

				sb.Append("М_ИОС(КЛ)_04_Таблица для классификации остальных категорий\n");
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
							str += fi.Category.Name + "\t" + fis.FamilyName + "\t" + el.Name + "\t" + "?\t" + result + "\t";
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
                outputString += "Заполнен классификатор: " + (int)(Math.Round((double)m.countIfMSKCOdIs / (double)m.countAllMSKCod * 100, 0)) + "%\n";
            else
                outputString += "Заполнен классификатор: 0%\n";
            if ((m.countAll > 0) || (m.countAll != m.countAllMSKCod))
                outputString += "Заполнения параметров элементов модели: " + (int)(Math.Round((double)(m.countIfParameterIs - m.countIfMSKCOdIs) / (double)(m.countAll - m.countAllMSKCod) * 100, 0)) + "%\n";
            else
                outputString += "Заполнения параметров элементов модели: 0%\n";

            sbStart.Append(outputString).Append("\n");
            #endregion

            StringBuilder sbResult = new StringBuilder();

            sbResult.Append(sbStart.ToString()).Append(sb.ToString());

            #region Завершение и Условные обозначения
            StringBuilder sbEnd = new StringBuilder();
            sbEnd.Append("\n\nУсловные обозначения: \n");
            sbEnd.Append("!!(_) \t - не добавлен в модель общий параметр\n");
            sbEnd.Append(noData + "\t - значение параметра не заполнено\n");
            sbEnd.Append(noParameter + "\t - не добавлен в модель общий параметр в экземпляре\n");
            sbEnd.Append(noCategory + "\t - не назначена категория, связанная с этим общим параметром\n");
            #endregion

            sbResult.Append(sbEnd.ToString());

            //string dirPath = @outputFolder + @"\"; // для динамо
            if (!Directory.Exists(m.workingDir))
                Directory.CreateDirectory(m.workingDir);

            string filePathToExcel = m.workingDir + m.CropFileName(doc.Title) + String.Format(" (МСК на {0:00}%)", m.readyOn) + ".xlsx";
            //string filePathToTxt = CropFileName(document.Title) + String.Format(" (МСК на {0:00}пр)", readyOn);
            string excelSheet = m.CropFileName(doc.Title);
            // writeToFile(dirPath, filePathToTxt, sb2.ToString());
            m.WriteToExcel(filePathToExcel, excelSheet, sbResult.ToString(), noData, noParameter, 37);
            //return outputString;


            TaskDialog.Show("Final", "готово!");
			m.OpenFolder(m.workingDir);
            return Result.Succeeded;
        }
    }
}
