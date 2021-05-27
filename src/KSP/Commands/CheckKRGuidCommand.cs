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
    class CheckKRGuidCommand : IExternalCommand
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

			/*
	         * М_КР_01.1.1_Таблица для заполнения параметров фундаментов (кроме плитных, общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFoundation).Where(x => x.Location != null).Where(x => !x.Name.Contains("лита")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Метод изготовления",
										"МСК_Уровень ответственности",
										"МСК_Расход арматуры",
										"МСК_Защитный слой 1",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_С подколонником",
										"МСК_Класс арматуры текст",
										"МСК_Марка В",
										"МСК_Водонепроницаемость W",
										"МСК_Морозостойкость F",
										"МСК_Обозначение материал"
					};
				sb.Append("М_КР_01.1.1_Таблица для заполнения параметров фундаментов (кроме плитных, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	        * М_КР_01.2.1_Таблица для заполнения параметров фундаментов (плитных, общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFoundation).Where(x => x.Location != null).Where(x => x.Name.Contains("лита")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Метод изготовления",
										"МСК_Уровень ответственности",
										"МСК_Расход арматуры",
										"МСК_Защитный слой 1",
										"МСК_Назначение",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_С подколонником",
										"МСК_Класс арматуры текст",
										"МСК_Марка В",
										"МСК_Водонепроницаемость W",
										"МСК_Морозостойкость F",
										"МСК_Обозначение материал"
					};
				sb.Append("М_КР_01.2.1_Таблица для заполнения параметров фундаментов (плитных, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_01.3.1_Таблица для заполнения параметров свай (общая) (разные категории)

			elements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => x.Category.Name != "Линии")
				.Where(x => x.Category.Name != "Элементы узлов")
				.Where(x => x.Category.Name == "свая" || x.Category.Name == "Свая")
				.ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"Категория",
										"Семейство",
										"МСК_Метод изготовления",
										"МСК_Уровень ответственности",
										"МСК_Расход арматуры",
										"МСК_Защитный слой 1",
										"МСК_Защитный слой 2",
										"МСК_Способ погружения",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Назначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Класс арматуры текст",
										"МСК_Марка В",
										"МСК_Водонепроницаемость W",
										"МСК_Морозостойкость F",
										"МСК_Плотность",
										"МСК_Обозначение материал"
					};

				sb.Append("М_КР_01.3.1_Таблица для заполнения параметров свай (общая) (разные категории)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParametersWCatNameWFamily(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_02.1_Таблица для заполнения параметров стен (общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Метод изготовления",
										"МСК_Уровень ответственности",
										"МСК_Расход арматуры",
										"МСК_Защитный слой 1",
										"МСК_Защитный слой 2",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Класс арматуры текст",
										"МСК_Марка В",
										"МСК_Водонепроницаемость W",
										"МСК_Морозостойкость F",
										"МСК_Обозначение материал"
					};
				sb.Append("М_КР_02.1_Таблица для заполнения параметров стен (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}



	         * М_КР_03.1_Таблица для заполнения параметров армирования перекрытий (общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Floors).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Метод изготовления",
										"МСК_Уровень ответственности",
										"МСК_Расход арматуры",
										"МСК_Защитный слой 1",
										"МСК_Назначение",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Класс арматуры текст",
										"МСК_Марка В",
										"МСК_Водонепроницаемость W",
										"МСК_Морозостойкость F",
										"МСК_Обозначение материал"
					};
				sb.Append("М_КР_03.1_Таблица для заполнения параметров армирования перекрытий (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_04.1_Таблица для заполнения параметров колонн (стальных)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralColumns).Where(x => x.Location != null)
				.Where(x => !x.Name.Contains("онолит")).Where(x => !x.Name.Contains("ж.б")).Where(x => !x.Name.Contains("Ж.б")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Метод изготовления",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал"
					};
				sb.Append("М_КР_04.1_Таблица для заполнения параметров колонн (стальных)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_04.2.1_Таблица для заполнения параметров колонн (железобетонных, общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralColumns).Where(x => x.Location != null)
				.Where(x => (x.Name.Contains("онолит") || x.Name.Contains("ж.б") || x.Name.Contains("Ж.б"))).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Метод изготовления",
										 "МСК_Уровень ответственности",
										 "МСК_Расход арматуры",
										 "МСК_Защитный слой 1",
										 "МСК_Класс арматуры текст",
										 "МСК_Назначение",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Марка В",
										 "МСК_Водонепроницаемость W",
										 "МСК_Морозостойкость F",
										 "МСК_Обозначение материал"
					};
				sb.Append("М_КР_04.2.1_Таблица для заполнения параметров колонн (железобетонных, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_05.1_Таблица для заполнения параметров балок (стальных)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFraming).Where(x => x.Location != null)
				.Where(x => !x.Name.Contains("онолит")).Where(x => !x.Name.Contains("ж.б")).Where(x => !x.Name.Contains("Ж.б")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Метод изготовления",
										 "МСК_Наименование",
										 "МСК_Назначение",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал"
					};
				sb.Append("М_КР_05.1_Таблица для заполнения параметров балок (стальных)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_05.2.1_Таблица для заполнения параметров балок (железобетонных, общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFraming).Where(x => x.Location != null)
				.Where(x => (x.Name.Contains("онолит") || x.Name.Contains("ж.б") || x.Name.Contains("Ж.б"))).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Метод изготовления",
										 "МСК_Расход арматуры",
										 "МСК_Защитный слой 1",
										 "МСК_Класс арматуры текст",
										 "МСК_Наименование",
										 "МСК_Назначение",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Марка В",
										 "МСК_Водонепроницаемость W",
										 "МСК_Морозостойкость F",
										 "МСК_Обозначение материал"
					};
				sb.Append("М_КР_05.2.1_Таблица для заполнения параметров балок (железобетонных, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_06.1.1_Таблица для заполнения параметров лестниц (общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Признак несущей конструкции",
										 "МСК_Наружный",
										 "МСК_Наименование",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Секция"
					};
				sb.Append("М_КР_06.1.1_Таблица для заполнения параметров лестниц (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_06.2.1_Таблица для заполнения параметров лестниц (марши, общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRuns).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Признак несущей конструкции",
										 "МСК_Наружный",
										 "МСК_Наименование",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Секция"
					};
				sb.Append("М_КР_06.2.1_Таблица для заполнения параметров лестниц (марши, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_07.1_Таблица для заполнения параметров пандусов (общая)

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Ramps).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Наружный",
										 "МСК_Наименование",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Секция"
					};
				sb.Append("М_КР_07.1_Таблица для заполнения параметров пандусов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


	         * М_КР_08_Таблица для заполнения параметров сборок

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Assemblies).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Наименование",
										 "МСК_Наружный",
										 "МСК_Метод изготовления",
										 "МСК_Марка",
										 "МСК_Обозначение"
					};
				sb.Append("М_КР_08_Таблица для заполнения параметров сборок\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));

			}


	         * М_КР(КЛ)_05_Таблица для классификации ограждений

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Семейство",
										 "Тип",
										 "Ключевая пометка",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_КР(КЛ)_05_Таблица для классификации ограждений\n");
				sb.Append(RowHeader(pSet));
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
							FamilySymbol fis = document.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter);
							MSKCounter("МСК_Код по классификатору", result, noData, noParameter);
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


	         * М_КР(КЛ)_07_Таблица для классификации форм

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Mass).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Семейство",
										 "Тип",
										 "Ключевая пометка",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_КР(КЛ)_07_Таблица для классификации форм\n");
				sb.Append(RowHeader(pSet));
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
							FamilySymbol fis = document.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter);
							MSKCounter("МСК_Код по классификатору", result, noData, noParameter);
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


	         * М_КР(КЛ)_08_Таблица для классификации крыш

			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Roofs).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Семейство",
										 "Тип",
										 "Ключевая пометка",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_КР(КЛ)_08_Таблица для классификации крыш\n");
				sb.Append(RowHeader(pSet));
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
							FamilySymbol fis = document.GetElement(fi.GetTypeId()) as FamilySymbol;
							sb.Append(fis.FamilyName).Append("\t");
							sb.Append(el.Name).Append("\t");
							sb.Append("?\t");
							string result = GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter);
							MSKCounter("МСК_Код по классификатору", result, noData, noParameter);
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


	         * М_КР(КЛ)_09_Таблица для классификации остальных категорий

			elements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance))
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
										 "Тип",
										 "Ключевая пометка",
										 "МСК_Код по классификатору"
					};

				sb.Append("М_КР(КЛ)_09_Таблица для классификации остальных категорий\n");
				sb.Append(RowHeader(pSet));
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
							FamilySymbol fis = document.GetElement(fi.GetTypeId()) as FamilySymbol;
							//							sb.Append(fi.Category.Name).Append("\t");
							//							sb.Append(fis.FamilyName).Append("\t");
							//							sb.Append(el.Name).Append("\t");
							//							sb.Append("?\t");
							//							sb.Append(getParameterValue(doc,el, "МСК_Код по классификатору", noData, noParameter)).Append("\t");
							//							sb.Append("\n");
							string result = GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter);
							MSKCounter("МСК_Код по классификатору", result, noData, noParameter);
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
						countAll += 1;

					if (str.Contains(noData))
						countAll += 1;

					sb.Append(str).Append("\n");
				}
			}
			*/

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
            sbEnd.Append("!!(_) \t - не добавлен в модель данный общий параметр\n");
            sbEnd.Append(noParameter + "\t - не добавлен в модель общий параметр\n");
            sbEnd.Append(noData + "\t - значение параметра не заполнено\n");
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
            return Result.Succeeded;
        }
    }
}
