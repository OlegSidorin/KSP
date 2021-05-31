namespace KSP
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Text;
    using System.Linq;
	using static KSP.CommonMethods;

	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class CheckMSKKRCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			Document document = uiDoc.Document;

			StringBuilder sb = new StringBuilder();
			countIfParameterIs = 0;
			countIfMSKCOdIs = 0;
			countAllMSKCod = 0;
			countAll = 0;
			readyOn = 0;
			//string outputFolder = ""; // = output_folder; // Для динамо
			string outputString = "";

			DateTime dt = DateTime.Now;
			//sb.Append("Отчет составлен: ").Append(String.Format("{0:dd.MM.yyyy}г.", dt)).Append("\n\n");
			/*
	         * М_КР_Таблица для ДОБАВЛЕНИЯ параметров модели
	         */
			IList<ProjectInfo> pInfo = new FilteredElementCollector(document).OfClass(typeof(ProjectInfo)).Cast<ProjectInfo>().ToList();
			if (pInfo != null)
			{
				string[] pSet = {
										"МСК_Тип проекта",
										"МСК_Степень огнестойкости",
										"МСК_Отметка нуля проекта",
										"МСК_Отметка уровня земли",
										"МСК_Проектировщик",
										"МСК_Заказчик",
										"МСК_Имя проекта",
										"МСК_Имя обьекта",
										"Номер проекта",
										"МСК_Корпус",
										"МСК_Номер секции",
										"МСК_Количество секций"
					};

				sb.Append("М_КР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var el in pInfo)
				{
					for (int i = 0; i < pSet.Count(); i++)
					{
						string result = GetParameterValue(document, el, pSet[i], noData, noParameter);
						sb.Append(result).Append("\t");
						MSKCounter(pSet[i], result, noData, noParameter);
					}
					sb.Append("\n");
				}
				sb.Append("\n");
			}

			/*
	         * М_КР_00_Таблица для заполнения параметров уровней
	         */
			IList<Level> levels = new FilteredElementCollector(document).OfClass(typeof(Level)).Cast<Level>().ToList();
			if (levels != null)
			{
				string[] pSet = {
										"Имя",
										"Фасад",
										"МСК_Надземный",
										"МСК_Базовый уровень",
										"МСК_Система пожаротушения",
										"МСК_Наличие АУПТ",
										"МСК_Уровень комфорта"
					};
				sb.Append("М_КР_00_Таблица для заполнения параметров уровней\n");
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
						string result = GetParameterValue(document, lvl, pSet[i], noData, noParameter);
						sb.Append(result).Append("\t");
						if (i > 1)
							MSKCounter(pSet[i], result, noData, noParameter);
					}
					sb.Append("\n");
				}
				sb.Append("\n");
			}

			IList<Element> elements = new List<Element>();

			/*
	         * М_КР_01.1.1_Таблица для заполнения параметров фундаментов (кроме плитных, общая)
	         */
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

			/*
	        * М_КР_01.2.1_Таблица для заполнения параметров фундаментов (плитных, общая)
	        */
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

			/*
	         * М_КР_01.3.1_Таблица для заполнения параметров свай (общая) (разные категории)
	         */
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

			/*
	         * М_КР_02.1_Таблица для заполнения параметров стен (общая)
	         */
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


			/*
	         * М_КР_03.1_Таблица для заполнения параметров армирования перекрытий (общая)
	         */
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

			/*
	         * М_КР_04.1_Таблица для заполнения параметров колонн (стальных)
	         */
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

			/*
	         * М_КР_04.2.1_Таблица для заполнения параметров колонн (железобетонных, общая)
	         */
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

			/*
	         * М_КР_05.1_Таблица для заполнения параметров балок (стальных)
	         */
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

			/*
	         * М_КР_05.2.1_Таблица для заполнения параметров балок (железобетонных, общая)
	         */
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

			/*
	         * М_КР_06.1.1_Таблица для заполнения параметров лестниц (общая)
	         */
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

			/*
	         * М_КР_06.2.1_Таблица для заполнения параметров лестниц (марши, общая)
	         */
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

			/*
	         * М_КР_07.1_Таблица для заполнения параметров пандусов (общая)
	         */
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

			/*
	         * М_КР_08_Таблица для заполнения параметров сборок
	         */
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

			/*
	         * М_КР(КЛ)_05_Таблица для классификации ограждений
	         */
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

			/*
	         * М_КР(КЛ)_07_Таблица для классификации форм
	         */
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

			/*
	         * М_КР(КЛ)_08_Таблица для классификации крыш
	         */
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

			/*
	         * М_КР(КЛ)_09_Таблица для классификации остальных категорий
	         */
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

			StringBuilder sb2 = new StringBuilder();
			sb2.Append("Отчет составлен: ").Append(String.Format("{0:dd.MM.yyyy}г.", dt)).Append("\n");
			sb2.Append("Заполнено параметров: \t" + countIfParameterIs + "\n");
			sb2.Append("Всего параметров: \t" + countAll + "\n");
			if (countAll > 0)
				readyOn = (int)(Math.Round((double)countIfParameterIs / (double)countAll * 100, 0));
			sb2.Append(String.Format("Заполнение: \t{0}%\n", readyOn));
			sb2.Append("\n");

			if (countAllMSKCod > 0)
				outputString += "Прогресс классификации элементов:\n" + (int)(Math.Round((double)countIfMSKCOdIs / (double)countAllMSKCod * 100, 0)) + "%\n";
			else
				outputString += "Прогресс классификации элементов:\n0%\n";
			if ((countAll > 0) || (countAll != countAllMSKCod))
				outputString += "Прогресс заполнения параметров элементов модели:\n" + (int)(Math.Round((double)(countIfParameterIs - countIfMSKCOdIs) / (double)(countAll - countAllMSKCod) * 100, 0)) + "%\n";
			else
				outputString += "Прогресс заполнения параметров элементов модели:\n0%\n";

			sb2.Append(outputString).Append("\n");
			sb2.Append(sb.ToString());

			if (!Directory.Exists(workingDir))
			{
				Directory.CreateDirectory(workingDir);
			}
			//string dirPath = @outputFolder + @"\"; // для динамо

			string filePathToExcel = workingDir + CropFileName(document.Title) + String.Format(" (МСК на {0:00}%)", readyOn) + ".xlsx";
			//string filePathToTxt = CropFileName(document.Title) + String.Format(" (МСК на {0:00}пр)", readyOn);
			string excelSheet = CropFileName(document.Title);
			//writeToFile(dirPath, filePathToTxt, sb2.ToString());
			WriteToExcel(filePathToExcel, excelSheet, sb2.ToString(), noData, noParameter, 22);
			return Result.Succeeded;
        }

	}
}
