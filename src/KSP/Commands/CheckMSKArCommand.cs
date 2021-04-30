namespace KSP
{
    using System;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Text;
    using System.Linq;
	using static KSP.CommonMethods;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class CheckMSKArCommand : IExternalCommand
    {
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementset)
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
			 * М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели
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

				sb.Append("М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
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
	         * М_АР_00_Таблица для заполнения параметров уровней
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
						string result = GetParameterValue(document, lvl, pSet[i], noData, noParameter);
						sb.Append(result).Append("\t");
						if (i > 1)
							MSKCounter(pSet[i], result, noData, noParameter);
					}
					sb.Append("\n");
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.1_Таблица для заполнения параметров зон (Площадь застройки)
	         */
			IList<Area> areas = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas).Cast<Area>().ToList();
			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны"
					};
				sb.Append("М_АР_01.1_Таблица для заполнения параметров зон (Площадь застройки)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Площадь застройки")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.2_Таблица для заполнения параметров зон (Общая площадь здания)
	         */

			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны"
					};
				sb.Append("М_АР_01.2_Таблица для заполнения параметров зон (Общая площадь здания)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Общая площадь здания")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.3_Таблица для заполнения параметров зон (Пожарная безопасность)
	         */

			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны",
										"МСК_Степень огнестойкости",
										"МСК_Кконстр_ПО",
										"МСК_Кфунк_ПО"
					};
				sb.Append("М_АР_01.3_Таблица для заполнения параметров зон (Пожарная безопасность)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Пожарный отсек")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.4_Таблица для заполнения параметров зон (ОДИ. Квартиры МГН)
	         */

			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны"
					};
				sb.Append("М_АР_01.4_Таблица для заполнения параметров зон (ОДИ. Квартиры МГН)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Квартира МГН")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.5_Таблица для заполнения параметров зон (ОДИ. Зона санитарно-бытовая МГН)
	         */

			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны"
					};
				sb.Append("М_АР_01.5_Таблица для заполнения параметров зон (ОДИ. Зона санитарно-бытовая МГН)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Зона санитарно-бытовая МГН")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}

			/*
	         * М_АР_01.6_Таблица для заполнения параметров зон (Наземная автостоянка/Подземная автостоянка)
	         */

			if (areas != null)
			{
				string[] pSet = {
										"Имя",
										"МСК_Код по классификатору",
										"МСК_Доступность МГН",
										"МСК_Признак наружного пространства",
										"МСК_Описание зон и помещений",
										"Номер",
										"МСК_Секция",
										"МСК_Вид деятельности",
										"МСК_Назначение",
										"МСК_Тип зоны",
										"МСК_Вместимость"
					};
				sb.Append("М_АР_01.6_Таблица для заполнения параметров зон (Наземная автостоянка/Подземная автостоянка)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i]).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Наземная автостоянка/Подземная автостоянка")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = GetParameterValue(document, a, pSet[i], noData, noParameter);
							sb.Append(result).Append("\t");
							if (i > 1)
								MSKCounter(pSet[i], result, noData, noParameter);
						}
						sb.Append("\n");
					}
					else
					{
						//
					}
				}
				sb.Append("\n");
			}



			IList<Element> elements = new List<Element>();

			/*
	         * М_АР_02_Таблица для заполнения параметров помещений
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Rooms).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Имя",
										 "МСК_Код по классификатору",
										 "МСК_Доступность МГН",
										 "МСК_Признак наружного пространства",
										 "МCК_Категория",
										 "МСК_Подпор воздуха",
										 "МСК_Путь эвакуации",
										 "МСК_Система пожаротушения",
										 "МСК_Наличие АУПТ",
										 "МСК_Имя помещения",
										 "МСК_Описание зон и помещений",
										 "МСК_Номер помещения",
										 "МСК_Секция",
										 "МСК_Вид деятельности",
										 "МСК_Назначение",
										 "МСК_Тип помещения",
										 "МСК_Полезная площадь",
										 "МСК_Расчетная площадь",
										 "МСК_Многосветное помещение",
										 "МСК_Мокрое помещение",
										 "МСК_Номер квартиры",
										 "МСК_Тип квартиры",
										 "МСК_Число комнат",
										 "МСК_Расчетное количество людей с постоянным пребыванием",
										 "МСК_Расчетное количество мест для инвалидов",
										 "МСК_Тип дымоудаления",
										 "МСК_Зона безопасности МГН",
										 "МСК_Чистое помещение",
										 "МСК_Тип противопожарной преграды",
										 "МСК_Зона",
										 "Полы",
										 "МСК_Покрытие пола",
										 "МСК_Толщина покрытия пола",
										 "МСК_Отделка стен",
										 "МСК_Толщина отделки стен",
										 "МСК_Отделка потолка",
										 "МСК_Толщина отделки потолка"
					};
				sb.Append("М_АР_02_Таблица для заполнения параметров помещений\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_03.1_Таблица заполнения параметров стен (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls).Where(x => x.Location != null).ToList();
			List<Element> wallsNotOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = GetParameterValue(document, e, pFilter[0], noData, noParameter);
					if (!s.Contains("тделка"))
						wallsNotOtdelka.Add(e);
				}

			}
			if (wallsNotOtdelka != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Признак несущей конструкции",
										"МСК_Предел огнестойкости",
										"МСК_Противопожарная преграда",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Тип противопожарной преграды"
					};
				sb.Append("М_АР_03.1_Таблица заполнения параметров стен (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, wallsNotOtdelka, noData, noParameter));
			}

			/*
	         * М_АР_04.1_Таблица для заполнения параметров навесных фасадов, панелей и витражей (общая)
	         */
			elements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => (x.Category.Name == "Импосты витража") || (x.Category.Name == "Панели витража"))
				.ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"Категория",
										"Семейство",
										"МСК_Предел огнестойкости",
										"МСК_Признак горючести",
										"МСК_Скорость распространения пламени",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал"
					};

				sb.Append("М_АР_04.1_Таблица для заполнения параметров навесных фасадов, панелей и витражей (общая) (разные категории)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParametersWCatNameWFamily(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_05.1_Таблица заполнения параметров перекрытий (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Floors).Where(x => x.Location != null).ToList();
			List<Element> floorNotOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = GetParameterValue(document, e, pFilter[0], noData, noParameter);
					if (!s.Contains("тделка"))
						floorNotOtdelka.Add(e);
				}

			}
			if (floorNotOtdelka != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Предел огнестойкости",
										 "МСК_Признак несущей конструкции",
										 "МСК_Противопожарная преграда",
										 "МСК_Наружный",
										 "МСК_Вид деятельности",
										 "МСК_Наименование",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Тип противопожарной преграды"
					};
				sb.Append("М_АР_05.1_Таблица заполнения параметров перекрытий (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, floorNotOtdelka, noData, noParameter));
			}

			/*
	         * М_АР_06.1.1_Таблица заполнения параметров отделки (стены, общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls).Where(x => x.Location != null).ToList();
			List<Element> wallsOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = GetParameterValue(document, e, pFilter[0], noData, noParameter);
					if (s.Contains("тделка"))
						wallsOtdelka.Add(e);
				}

			}
			if (wallsOtdelka != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Признак горючести",
										"МСК_Воспламеняемость",
										"МСК_Горючесть",
										"МСК_Скорость распространения пламени",
										"МСК_Наружный",
										"МСК_Вид деятельности",
										"МСК_Наименование",
										"МСК_Обозначение," +
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Наименование работ краткое"
					};
				sb.Append("М_АР_06.1.1_Таблица заполнения параметров отделки (стены, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, wallsOtdelka, noData, noParameter));
			}

			/*
	         * М_АР_06.2.1_Таблица заполнения параметров отделки (перекрытия, общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Floors).Where(x => x.Location != null).ToList();
			List<Element> floorOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = GetParameterValue(document, e, pFilter[0], noData, noParameter);
					if (s.Contains("тделка"))
						floorOtdelka.Add(e);
				}

			}
			if (floorOtdelka != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Признак горючести",
										 "МСК_Воспламеняемость",
										 "МСК_Горючесть",
										 "МСК_Скорость распространения пламени",
										 "МСК_Наружный",
										 "МСК_Вид деятельности",
										 "МСК_Наименование",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Наименование работ краткое"
					};
				sb.Append("М_АР_06.2.1_Таблица заполнения параметров отделки (перекрытия, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, floorOtdelka, noData, noParameter));
			}

			/*
	         * М_АР_06.3.1_Таблица заполнения параметров отделки (потолки, общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Ceilings).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Признак горючести",
										 "МСК_Воспламеняемость",
										 "МСК_Горючесть",
										 "МСК_Скорость распространения пламени",
										 "МСК_Наружный",
										 "МСК_Вид деятельности",
										 "МСК_Наименование",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал"
					};
				sb.Append("М_АР_06.3.1_Таблица заполнения параметров отделки (потолки, общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


			/*
	         * М_АР_07.1_Таблица заполнения параметров колонн (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralColumns).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Признак несущей конструкции",
										"МСК_Предел огнестойкости",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал"
					};
				sb.Append("М_АР_07.1_Таблица заполнения параметров колонн (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_08.1_Таблица заполнения параметров дверей
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Doors).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Класс устойчивости ко взлому",
										 "МСК_Устойчивость к разрушающим воздействиям",
										 "МСК_Процент остекления",
										 "МСК_Доступность МГН",
										 "МСК_Предел огнестойкости",
										 "МСК_Путь эвакуации",
										 "МСК_Автоматическое открывание",
										 "МСК_Автоматическое закрывание",
										 "МСК_Наружный",
										 "МСК_Количество слоев стекол",
										 "МСК_Наименование газа-заполнителя камеры",
										 "МСК_Ламинирование",
										 "МСК_Армирование",
										 "МСК_Решетка",
										 "МСК_Наименование",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Код материала2",
										 "МСК_Материал2",
										 "МСК_Противопожарная преграда",
										 "МСК_Тип противопожарной преграды",
										 "МСК_Остекление",
										 "МСК_Тип открывания",
										 "МСК_Высота порога"

					};
				sb.Append("М_АР_08.1_Таблица заполнения параметров дверей\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


			/*
	         * М_АР_09.1_Таблица заполнения параметров окон (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Windows).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Площадь отстекления",
										"МСК_Путь эвакуации",
										"МСК_Автоматическое открывание",
										"МСК_Автоматическое закрывание",
										"МСК_Наружный",
										"МСК_Количество слоев стекол",
										"МСК_Наименование газа-заполнителя камеры",
										"МСК_Ламинирование",
										"МСК_Армирование",
										"МСК_Решетка",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Материал",
										"МСК_Тип окна",
										"МСК_Тип створок",
										"МСК_Легкосбрасываемые",
										"МСК_Противопожарная преграда",
										"МСК_Тип противопожарной преграды",
										"Высота нижнего бруса",
										"МСК_Светопропускание"

					};
				sb.Append("М_АР_09.1_Таблица заполнения параметров окон (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}


			/*
	         * М_АР_10.1.1_Таблица заполнения параметров лестниц (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										 "Тип",
										 "МСК_Код по классификатору",
										 "МСК_Предел огнестойкости",
										 "Текущее количество подступенков",
										 "Текущая высота подступенка",
										 "Текущая ширина проступи",
										 "МСК_Доступность МГН",
										 "МСК_Признак несущей конструкции",
										 "МСК_Путь эвакуации",
										 "МСК_Наружный",
										 "МСК_Наименование",
										 "МСК_Форма марша",
										 "МСК_Марка",
										 "МСК_Обозначение",
										 "МСК_Код материала",
										 "МСК_Материал",
										 "МСК_Вид деятельности",
										 "МСК_Секция"
					};
				sb.Append("М_АР_10.1.1_Таблица заполнения параметров лестниц (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_10.2.1_Таблица заполнения параметров лестничных маршей (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRuns).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Предел огнестойкости",
										"Текущее количество подступенков",
										"Текущее количество проступей",
										"Текущая ширина марша",
										"Текущая высота подступенка",
										"Текущая ширина проступи",
										"МСК_Доступность МГН",
										"МСК_Признак несущей конструкции",
										"МСК_Путь эвакуации",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Форма марша",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Вид деятельности",
										"МСК_Секция"
					};
				sb.Append("М_АР_10.2.1_Таблица заполнения параметров лестничных маршей (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_11.1_Таблица заполнения параметров пандусов (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Ramps).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"МСК_Предел огнестойкости",
										"МСК_Высота проезда",
										"МСК_Доступность МГН",
										"МСК_Признак несущей конструкции",
										"МСК_Путь эвакуации",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Форма пандуса",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Секция"
					};
				sb.Append("М_АР_11.1_Таблица заполнения параметров пандусов (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_12.1_Таблица для заполнения параметров ограждений (общая)
	         */
			elements = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing).Where(x => x.Location != null).ToList();
			if (elements != null)
			{
				string[] pSet = {
										"Тип",
										"МСК_Код по классификатору",
										"Высота ограждения",
										"МСК_Размер_Диаметр",
										"МСК_Наружный",
										"МСК_Наименование",
										"МСК_Марка",
										"МСК_Обозначение",
										"МСК_Код материала",
										"МСК_Материал",
										"МСК_Секция"
					};
				sb.Append("М_АР_12.1_Таблица для заполнения параметров ограждений (общая)\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			}

			/*
	         * М_АР_13_Таблица для заполнения параметров сборок
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
				sb.Append("М_АР_13_Таблица для заполнения параметров сборок\n");
				sb.Append(RowHeader(pSet));
				sb.Append(RowElementsParameters(document, pSet, elements, noData, noParameter));
			};


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
	         * М_АР(КЛ)_13_Таблица для классификации остальных категорий
	         */
			elements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance)).Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => x.Category.Name != "Линии")
				.Where(x => x.Category.Name != "Элементы узлов")
				.Where(x => x.Category.Name != "Двери")
				.Where(x => x.Category.Name != "Импосты витража")
				.Where(x => x.Category.Name != "Панели витража")
				.Where(x => x.Category.Name != "Окна")
				.Where(x => x.Category.Name != "Несущие колонны")
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

				sb.Append("М_АР(КЛ)_13_Таблица для классификации остальных категорий\n");
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

			//string dirPath = @outputFolder + @"\"; // для динамо
			string dirPath = "C:\\Users\\Sidorin_O\\Documents\\TEST\\";
			string filePathToExcel = dirPath + CropFileName(document.Title) + String.Format(" (МСК на {0:00}%)", readyOn) + ".xlsx";
			//string filePathToTxt = CropFileName(document.Title) + String.Format(" (МСК на {0:00}пр)", readyOn);
			string excelSheet = CropFileName(document.Title);
			// writeToFile(dirPath, filePathToTxt, sb2.ToString());
			WriteToExcel(filePathToExcel, excelSheet, sb2.ToString(), noData, noParameter, 37);
			//return outputString;

			return Result.Succeeded;
        }
	}
}
