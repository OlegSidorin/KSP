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
    class CheckMSKIOSCommand : IExternalCommand
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
	         * М_ИОС_Таблица для ДОБАВЛЕНИЯ параметров модели
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
						string result = GetParameterValue(document, el, pSet[i], noData, noParameter);
						sb.Append(result).Append("\t");
						MSKCounter(pSet[i], result, noData, noParameter);
					}
					sb.Append("\n");
				}
				sb.Append("\n");
			}

			/*
	         * М_ИОС_00 Таблица для заполнения параметров уровней
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
	         * М_ИОС_01.1_Таблица для заполнения параметров воздуховодов (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_DuctCurves).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Форма воздуховода",
										"Рабочее давление",
										"Диаметр",
										"Ширина",
										"Высота",
										"МСК_Заводская изоляция",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Толщина стенки",
										"МСК_Огнестойкость EI",
										"МСК_Горючесть",
										"МСК_Признак ЭЭ",
										"МСК_Материал теплоизоляции"
					};
				sb.Append("М_ИОС_01.1_Таблица для заполнения параметров воздуховодов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС_02.1_Таблица для заполнения параметров фитингов возудоводов (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_DuctFitting).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Форма воздуховода",
										"МСК_Коэффициент шероховатости",
										"МСК_Заводская изоляция",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Толщина стенки",
										"МСК_Огнестойкость EI",
										"МСК_Горючесть",
										"МСК_Признак ЭЭ",
										"МСК_Материал теплоизоляции"
					};
				sb.Append("М_ИОС_02.1_Таблица для заполнения параметров фитингов возудоводов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС_03.1_Таблица для заполнения параметров воздухораспредлительных устройств (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_DuctTerminal).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Форма устройства",
										"МСК_Тип решетки",
										"МСК_Признак ЭЭ",
										"МСК_Наименование",
										"МСК_Описание",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Горючесть"
					};
				sb.Append("М_ИОС_03.1_Таблица для заполнения параметров воздухораспредлительных устройств (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС_04.1_Таблица для заполнения параметров труб (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_PipeCurves).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Толщина труб",
										"МСК_Горючесть",
										"МСК_Признак ЭЭ",
										"МСК_Материал теплоизоляции"
					};
				sb.Append("М_ИОС_04.1_Таблица для заполнения параметров труб (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС_05.1_Таблица для заполнения параметров фитингов труб (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_PipeFitting).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Потеря давления жидкости",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Толщина труб",
										"МСК_Горючесть",
										"МСК_Признак ЭЭ",
										"МСК_Материал теплоизоляции"
					};
				sb.Append("М_ИОС_05.1_Таблица для заполнения параметров фитингов труб (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС_06.1_Таблица для заполнения параметров отопительных приборов (оборудование) (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_MechanicalEquipment).Where(x => x.Location != null).Where(x => !x.Name.Contains("асос")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Тип расположения",
										"МСК_Масса",
										"МСК_Удельная теплоемкость",
										"МСК_Тепловая мощность",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Тип отоп_приборов",
										"МСК_Код материала",
										"МСК_Материал наименование",
										"МСК_Горючесть"
					};
				sb.Append("М_ИОС_06.1_Таблица для заполнения параметров отопительных приборов (оборудование) (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	        * М_ИОС_07.1_Таблица для заполнения параметров распределительных устройств (общая)
	        */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_ElectricalEquipment).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Полная мощность",
										"МСК_Мощность в режиме тушения пожара",
										"МСК_Класс защиты",
										"МСК_Степень защиты от удара",
										"МСК_Основное устройство",
										"МСК_Уровень квалификации",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение"
					};
				sb.Append("М_ИОС_07.1_Таблица для заполнения параметров распределительных устройств (электрооборудование) (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	        * М_ИОС_08.1_Таблица для заполнения параметров осветительных приборов (общая)
	        */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_LightingFixtures).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Полная мощность",
										"МСК_Мощность в режиме тушения пожара",
										"МСК_Признак заземления",
										"МСК_Класс защиты",
										"МСК_Степень защиты от удара",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Назначение",
										"МСК_Наружный",
										"МСК_Продолжительность автономной работы"
					};
				sb.Append("М_ИОС_08.1_Таблица для заполнения параметров осветительных приборов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	        * М_ИОС_09.1_Таблица для заполнения параметров кабельных лотков (общая)
	        */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_CableTray).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Тип конструкции",
										"МСК_Предел огнестойкости",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Материал перегородки лотка",
										"МСК_Коррозионная защита"
					};
				sb.Append("М_ИОС_09.1_Таблица для заполнения параметров кабельных лотков (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	        * М_ИОС_10.1_Таблица для заполнения параметров фитингов кабельных лотков (общая)
	        */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_CableTrayFitting).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Тип конструкции",
										"МСК_Предел огнестойкости",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Материал перегородки лотка",
										"МСК_Коррозионная защита"
					};
				sb.Append("М_ИОС_10.1_Таблица для заполнения параметров фитингов кабельных лотков (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	        * М_ИОС_11.1_Таблица заполнения параметров насосов (общая)
	        */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_MechanicalEquipment).Where(x => x.Location != null).Where(x => x.Name.Contains("асос")).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Полная мощность",
										"МСК_Мощность в режиме тушения пожара",
										"МСК_Класс защиты",
										"МСК_Степень защиты от удара",
										"МСК_Производительность",
										"МСК_Напор",
										"МСК_Скорость вращения",
										"МСК_Диаметр рабочего колеса",
										"МСК_Тип расположения",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Тип устройства"
					};
				sb.Append("М_ИОС_11.1_Таблица заполнения параметров насосов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС(КЛ)_01_Таблица для классификации пространств 
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_MEPSpaces).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Имя",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_ИОС(КЛ)_01_Таблица для классификации пространств\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_ИОС(КЛ)_02_Таблица для классификации сборок
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Assemblies).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Семейство",
										 "Тип",
										 "Ключевая пометка",
										 "МСК_Код по классификатору"
					};
				sb.Append("М_ИОС(КЛ)_02_Таблица для классификации сборок\n");
				sb.Append(RowHeader(pSet));
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
							AssemblyType aiType = document.GetElement(ai.GetTypeId()) as AssemblyType;
							sb.Append(aiType.FamilyName).Append("\t");
						}

					}
					catch
					{
						sb.Append(noParameter).Append("\t");
					}
					sb.Append(ai.AssemblyTypeName).Append("\t");
					sb.Append("?\t");
					sb.Append(GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter)).Append("\t");
					sb.Append("\n");
				}

				sb.Append("\n");

			}

			/*
	         * М_ИОС(КЛ)_03_Таблица для классификации форм
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
				sb.Append("М_ИОС(КЛ)_03_Таблица для классификации форм\n");
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
							sb.Append(GetParameterValue(document, el, "МСК_Код по классификатору", noData, noParameter)).Append("\t");
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
	         * М_ИОС(КЛ)_04_Таблица для классификации остальных категорий
	         */
			elements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance))
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
										 "МСК_Код по классификатору"
					};

				sb.Append("М_ИОС(КЛ)_04_Таблица для классификации остальных категорий\n");
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
			if (countAll != 0)
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

			//string dirPath = @outputFolder + @"\"; // для динамо
			if (!Directory.Exists(workingDir))
			{
				Directory.CreateDirectory(workingDir);
			}
			string filePathToExcel = workingDir + CropFileName(document.Title) + String.Format(" (МСК на {0:00}%)", readyOn) + ".xlsx";
			//string filePathToTxt = CropFileName(document.Title) + String.Format(" (МСК на {0:00}%)", readyOn);
			string excelSheet = "МСК_" + "Параметры";
			//writeToFile(dirPath, filePathToTxt, sb2.ToString());
			WriteToExcel(filePathToExcel, excelSheet, sb2.ToString(), noData, noParameter, 18);
			return Result.Succeeded;
        }

	}
}
