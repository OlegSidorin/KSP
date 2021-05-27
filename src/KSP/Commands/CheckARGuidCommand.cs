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

            // М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели
            IList<ProjectInfo> pInfo = new FilteredElementCollector(doc).OfClass(typeof(ProjectInfo)).Cast<ProjectInfo>().ToList();
            MyParameter value = new MyParameter("", "", false, false);
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

                sb.Append("М_АР_Таблица для ДОБАВЛЕНИЯ параметров модели\n");
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

            //М_АР_00_Таблица для заполнения параметров уровней
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

                sb.Append("М_АР_00_Таблица для заполнения параметров уровней\n");
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

	        //М_АР_01.1_Таблица для заполнения параметров зон (Площадь застройки)
            IList<Area> areas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).Cast<Area>().ToList();
            if (areas != null)
            {
                string[] pSet = {
                    "Имя",
                    m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
                    m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
                    m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
                    m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
                    "Номер",
                    m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
                    m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
                    m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
                    m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны")
                };
                sb.Append("М_АР_01.1_Таблица для заполнения параметров зон (Площадь застройки)\n");
                for (int i = 0; i < pSet.Count(); i++)
                {
                    sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
                }
                sb.Append("\n");

                foreach (var a in areas)
                {
                    if (a.AreaScheme.Name == "Площадь застройки")
                    {
                        for (int i = 0; i < pSet.Count(); i++)
                        {
                            string result = m.GetParameterValue(doc, a, pSet[i]);
                            sb.Append(result).Append("\t");
                            if (i > 1)
                                m.MSKCounter(pSet[i], result);
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

	        // М_АР_01.2_Таблица для заполнения параметров зон (Общая площадь здания)
			if (areas != null)
			{
				string[] pSet = {
					"Имя",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					"Номер",
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны")
				};
				sb.Append("М_АР_01.2_Таблица для заполнения параметров зон (Общая площадь здания)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Общая площадь здания")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = m.GetParameterValue(doc, a, pSet[i]);
							sb.Append(result).Append("\t");
							if (i > 1)
								m.MSKCounter(pSet[i], result);
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

	        // М_АР_01.3_Таблица для заполнения параметров зон (Пожарная безопасность)
			if (areas != null)
			{
				string[] pSet = {
					"Имя",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					"Номер",
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны"),
					m.SharedParameterFromGUIDName("efd17af1-dc28-4f05-8dfb-49995f260aba", myParameters, "МСК_Степень огнестойкости"),
					m.SharedParameterFromGUIDName("6fb90410-b21f-4f49-a1d7-243e00552d36", myParameters, "МСК_Кконстр_ПО"),
					m.SharedParameterFromGUIDName("0c71b198-9fe8-4328-8a3e-f1a32f1a1220", myParameters, "МСК_Кфунк_ПО")
				};
				sb.Append("М_АР_01.3_Таблица для заполнения параметров зон (Пожарная безопасность)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Пожарный отсек")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = m.GetParameterValue(doc, a, pSet[i]);
							sb.Append(result).Append("\t");
							if (i > 1)
								m.MSKCounter(pSet[i], result);
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

	        //М_АР_01.4_Таблица для заполнения параметров зон (ОДИ. Квартиры МГН)
			if (areas != null)
			{
				string[] pSet = {
					"Имя",
                    m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					"Номер",
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны")
				};
				sb.Append("М_АР_01.4_Таблица для заполнения параметров зон (ОДИ. Квартиры МГН)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Квартира МГН")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = m.GetParameterValue(doc, a, pSet[i]);
							sb.Append(result).Append("\t");
							if (i > 1)
								m.MSKCounter(pSet[i], result);
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

	        //М_АР_01.5_Таблица для заполнения параметров зон (ОДИ. Зона санитарно-бытовая МГН)
			if (areas != null)
			{
				string[] pSet = {
					"Имя",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					"Номер",
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны")
				};
				sb.Append("М_АР_01.5_Таблица для заполнения параметров зон (ОДИ. Зона санитарно-бытовая МГН)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Зона санитарно-бытовая МГН")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = m.GetParameterValue(doc, a, pSet[i]);
							sb.Append(result).Append("\t");
							if (i > 1)
								m.MSKCounter(pSet[i], result);
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

	        // М_АР_01.6_Таблица для заполнения параметров зон (Наземная автостоянка/Подземная автостоянка)
			if (areas != null)
			{
				string[] pSet = {
					"Имя",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					"Номер",
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("7f615183-6e80-4b78-98e2-132a8e78ee06", myParameters, "МСК_Тип зоны"),
					m.SharedParameterFromGUIDName("68d23d66-ef3e-4c3e-9bd8-128a5718d4c6", myParameters, "МСК_Вместимость")
				};
				sb.Append("М_АР_01.6_Таблица для заполнения параметров зон (Наземная автостоянка/Подземная автостоянка)\n");
				for (int i = 0; i < pSet.Count(); i++)
				{
					sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
				}
				sb.Append("\n");

				foreach (var a in areas)
				{
					if (a.AreaScheme.Name == "Наземная автостоянка/Подземная автостоянка")
					{
						for (int i = 0; i < pSet.Count(); i++)
						{
							string result = m.GetParameterValue(doc, a, pSet[i]);
							sb.Append(result).Append("\t");
							if (i > 1)
								m.MSKCounter(pSet[i], result);
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

	        //М_АР_02_Таблица для заполнения параметров помещений
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Имя",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("df9939e1-a33a-47d4-8f36-6a1cca58117c", myParameters, "МСК_Признак наружного пространства"),
					m.SharedParameterFromGUIDName("50ffa05b-c914-4445-a91d-00b7734d7474", myParameters, "МCК_Категория"),
					m.SharedParameterFromGUIDName("ddef9515-cc48-40c7-9bbc-1348a645f84a", myParameters, "МСК_Подпор воздуха"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("741ca870-26ce-460c-8c82-56c23ec53ebe", myParameters, "МСК_Система пожаротушения"),
					m.SharedParameterFromGUIDName("6dd43fcd-977a-498c-8269-d3a40ac3dce3", myParameters, "МСК_Наличие АУПТ"),
					m.SharedParameterFromGUIDName("d7825a27-fb45-4b3b-9e99-9fbb8cab5d6f", myParameters, "МСК_Имя помещения"),
					m.SharedParameterFromGUIDName("83454a18-368f-467c-a5fc-24727da0f618", myParameters, "МСК_Описание зон и помещений"),
					m.SharedParameterFromGUIDName("83cd0930-944f-4dd6-82e4-938f33dc2ed8", myParameters, "МСК_Номер помещения"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("71769132-b752-4452-9cce-b9fc8ddac241", myParameters, "МСК_Назначение"),
					m.SharedParameterFromGUIDName("56eb1705-f327-4774-b212-ef9ad2c860b0", myParameters, "МСК_Тип помещения"),
					m.SharedParameterFromGUIDName("763294b4-d5f5-4ca8-82fa-ce4db5cddb8f", myParameters, "МСК_Полезная площадь"),
					m.SharedParameterFromGUIDName("9f744f48-f020-413c-afaa-3db9a7365bef", myParameters, "МСК_Расчетная площадь"),
					m.SharedParameterFromGUIDName("c838cb78-2d4f-43e9-b67d-18e1cc84f76f", myParameters, "МСК_Многосветное помещение"),
					m.SharedParameterFromGUIDName("84b9df7c-184a-43c8-8e02-568f135eddc0", myParameters, "МСК_Мокрое помещение"),
					m.SharedParameterFromGUIDName("10fb72de-237e-4b9c-915b-8849b8907695", myParameters, "МСК_Номер квартиры"),
					m.SharedParameterFromGUIDName("78e3b89c-eb68-4600-84a7-c523de162743", myParameters, "МСК_Тип квартиры"),
					m.SharedParameterFromGUIDName("b63e029c-b976-4322-92b3-f338aabc168c", myParameters, "МСК_Число комнат"),
					m.SharedParameterFromGUIDName("c7feec1c-9972-4277-9bb1-b4ad207f7bca", myParameters, "МСК_Расчетное количество людей с постоянным пребыванием"),
					m.SharedParameterFromGUIDName("6b14078b-442c-43c6-8d15-ac445f4180b1", myParameters, "МСК_Расчетное количество мест для инвалидов"),
					m.SharedParameterFromGUIDName("c29c8299-8dfa-4cef-a8e4-1bc39ca7eb2e", myParameters, "МСК_Тип дымоудаления"),
					m.SharedParameterFromGUIDName("409712c6-cfa3-46c8-82ab-ee6db3808003", myParameters, "МСК_Зона безопасности МГН"),
					m.SharedParameterFromGUIDName("e8d80956-05a5-4921-a896-1ac0417a2309", myParameters, "МСК_Чистое помещение"),
					"МСК_Тип противопожарной преграды",
					m.SharedParameterFromGUIDName("c78f0a7d-b68b-4d21-a247-1c8c6ced8bc5", myParameters, "МСК_Зона"),
					"Полы",
					m.SharedParameterFromGUIDName("75ba5bd0-cadf-4741-9cef-f60a79b041ac", myParameters, "МСК_Покрытие пола"),
					m.SharedParameterFromGUIDName("a0c3c626-03d9-4e4b-96fd-c681586e21aa", myParameters, "МСК_Толщина покрытия пола"),
					m.SharedParameterFromGUIDName("47133ec1-c86b-4706-a998-1e5da5b7b2f0", myParameters, "МСК_Отделка стен"),
					m.SharedParameterFromGUIDName("5fea80cd-7116-4513-869c-4e3bbc71de8c", myParameters, "МСК_Толщина отделки стен"),
					m.SharedParameterFromGUIDName("bcae7ced-045a-4390-ad69-dad245ee6608", myParameters, "МСК_Отделка потолка"),
					m.SharedParameterFromGUIDName("9f54a1c8-e4b1-40ce-8beb-6bc30daa4194", myParameters, "МСК_Толщина отделки потолка")
				};
				sb.Append("М_АР_02_Таблица для заполнения параметров помещений\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        //М_АР_03.1_Таблица заполнения параметров стен (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToList();
			List<Element> wallsNotOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = m.GetParameterValue(doc, e, pFilter[0]);
					if (!s.Contains("тделка"))
						wallsNotOtdelka.Add(e);
				}

			}
			if (wallsNotOtdelka != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("96901a85-e765-4879-957c-b8fc4ff2858b", myParameters, "МСК_Противопожарная преграда"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					"МСК_Тип противопожарной преграды"
				};
				sb.Append("М_АР_03.1_Таблица заполнения параметров стен (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, wallsNotOtdelka));
			}

	        //М_АР_04.1_Таблица для заполнения параметров навесных фасадов, панелей и витражей (общая)
			elements = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
				.Where(x => x.Category.CategoryType == CategoryType.Model)
				.Where(x => (x.Category.Name == "Импосты витража") || (x.Category.Name == "Панели витража"))
				.ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Категория",
					"Семейство",
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("8e6e80c1-5cde-4cb3-af65-e6db9a0a9174", myParameters, "МСК_Признак горючести"),
					m.SharedParameterFromGUIDName("c93ee364-c38e-4bf4-afae-a8d9a4aca50f", myParameters, "МСК_Скорость распространения пламени"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал")
				};

				sb.Append("М_АР_04.1_Таблица для заполнения параметров навесных фасадов, панелей и витражей (общая) (разные категории)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParametersWCatNameWFamily(doc, pSet, elements));
			}

	        // М_АР_05.1_Таблица заполнения параметров перекрытий (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToList();
			List<Element> floorNotOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = m.GetParameterValue(doc, e, pFilter[0]);
					if (!s.Contains("тделка"))
						floorNotOtdelka.Add(e);
				}

			}
			if (floorNotOtdelka != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					m.SharedParameterFromGUIDName("96901a85-e765-4879-957c-b8fc4ff2858b", myParameters, "МСК_Противопожарная преграда"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					"МСК_Тип противопожарной преграды"
				};
				sb.Append("М_АР_05.1_Таблица заполнения параметров перекрытий (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, floorNotOtdelka));
			}

	        // М_АР_06.1.1_Таблица заполнения параметров отделки (стены, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToList();
			List<Element> wallsOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = m.GetParameterValue(doc, e, pFilter[0]);
					if (s.Contains("тделка"))
						wallsOtdelka.Add(e);
				}

			}
			if (wallsOtdelka != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("8e6e80c1-5cde-4cb3-af65-e6db9a0a9174", myParameters, "МСК_Признак горючести"),
					m.SharedParameterFromGUIDName("669ec82a-2d66-4d29-8a1a-ec611c22e323", myParameters, "МСК_Воспламеняемость"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c93ee364-c38e-4bf4-afae-a8d9a4aca50f", myParameters, "МСК_Скорость распространения пламени"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("0a45214f-65cc-41a7-b908-0dd592c1a47f", myParameters, "МСК_Наименование работ краткое")
				};
				sb.Append("М_АР_06.1.1_Таблица заполнения параметров отделки (стены, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, wallsOtdelka));
			}

	        // М_АР_06.2.1_Таблица заполнения параметров отделки (перекрытия, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToList();
			List<Element> floorOtdelka = new List<Element>();
			if (elements != null)
			{
				string[] pFilter = {
						"МСК_Наименование работ краткое"
					};
				foreach (Element e in elements)
				{
					string s = m.GetParameterValue(doc, e, pFilter[0]);
					if (s.Contains("тделка"))
						floorOtdelka.Add(e);
				}

			}
			if (floorOtdelka != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("8e6e80c1-5cde-4cb3-af65-e6db9a0a9174", myParameters, "МСК_Признак горючести"),
					m.SharedParameterFromGUIDName("669ec82a-2d66-4d29-8a1a-ec611c22e323", myParameters, "МСК_Воспламеняемость"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c93ee364-c38e-4bf4-afae-a8d9a4aca50f", myParameters, "МСК_Скорость распространения пламени"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("0a45214f-65cc-41a7-b908-0dd592c1a47f", myParameters, "МСК_Наименование работ краткое")
					};
				sb.Append("М_АР_06.2.1_Таблица заполнения параметров отделки (перекрытия, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, floorOtdelka));
			}

	        //М_АР_06.3.1_Таблица заполнения параметров отделки (потолки, общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("8e6e80c1-5cde-4cb3-af65-e6db9a0a9174", myParameters, "МСК_Признак горючести"),
					m.SharedParameterFromGUIDName("669ec82a-2d66-4d29-8a1a-ec611c22e323", myParameters, "МСК_Воспламеняемость"),
					m.SharedParameterFromGUIDName("db32c5ee-af89-48b7-ad84-b6fa5e19d5c2", myParameters, "МСК_Горючесть"),
					m.SharedParameterFromGUIDName("c93ee364-c38e-4bf4-afae-a8d9a4aca50f", myParameters, "МСК_Скорость распространения пламени"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
				};
				sb.Append("М_АР_06.3.1_Таблица заполнения параметров отделки (потолки, общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}


	        //М_АР_07.1_Таблица заполнения параметров колонн (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал")
				};
				sb.Append("М_АР_07.1_Таблица заполнения параметров колонн (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

			
	        // М_АР_08.1_Таблица заполнения параметров дверей
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("8715f455-f9d7-4f39-b23f-7b3f437c830d", myParameters, "МСК_Класс устойчивости ко взлому"),
					m.SharedParameterFromGUIDName("61a4dfb6-d951-4135-a709-6a00fe82977e", myParameters, "МСК_Устойчивость к разрушающим воздействиям"),
					m.SharedParameterFromGUIDName("3fc9867b-acf5-4f2b-bc26-b773ce2c314a", myParameters, "МСК_Процент остекления"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("2eebaa27-2210-4552-aeb0-e37ef07eca4c", myParameters, "МСК_Автоматическое открывание"),
					m.SharedParameterFromGUIDName("906135ad-c889-4b71-9b94-ea640a4380d6", myParameters, "МСК_Автоматическое закрывание"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("420e30f6-9d02-41c5-881d-1a782f407277", myParameters, "МСК_Количество слоев стекол"),
					m.SharedParameterFromGUIDName("f1cb5818-900b-407a-85a5-1aac4c0da623", myParameters, "МСК_Наименование газа-заполнителя камеры"),
					m.SharedParameterFromGUIDName("7b3af2b7-5b5f-4322-a588-545c0e3f1cac", myParameters, "МСК_Ламинирование"),
					m.SharedParameterFromGUIDName("288f2154-a879-410e-80ca-aaebe7a6b385", myParameters, "МСК_Армирование"),
					m.SharedParameterFromGUIDName("00d2d2cd-e493-4607-bfad-2e0842697c5a", myParameters, "МСК_Решетка"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("1553f66a-0009-4011-b9ad-c2a6d55196af", myParameters, "МСК_Код материала2"),
					m.SharedParameterFromGUIDName("24ec3719-0732-4663-9045-d95f820b6db5", myParameters, "МСК_Материал2"),
					m.SharedParameterFromGUIDName("96901a85-e765-4879-957c-b8fc4ff2858b", myParameters, "МСК_Противопожарная преграда"),
					"МСК_Тип противопожарной преграды",
					m.SharedParameterFromGUIDName("a058b8f4-1715-4837-bbbf-849f18eb2c0b", myParameters, "МСК_Остекление"),
					m.SharedParameterFromGUIDName("ba2b1451-e57a-4431-868d-00b46c7e8a93", myParameters, "МСК_Тип открывания"),
					m.SharedParameterFromGUIDName("c6910d56-4529-4fbb-9ddb-35fe49aa5c41", myParameters, "МСК_Высота порога")

				};
				sb.Append("М_АР_08.1_Таблица заполнения параметров дверей\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_АР_09.1_Таблица заполнения параметров окон (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("85f01269-04b3-4437-891c-e3ceacfeeabc", myParameters, "МСК_Площадь отстекления"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("2eebaa27-2210-4552-aeb0-e37ef07eca4c", myParameters, "МСК_Автоматическое открывание"),
					m.SharedParameterFromGUIDName("906135ad-c889-4b71-9b94-ea640a4380d6", myParameters, "МСК_Автоматическое закрывание"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("420e30f6-9d02-41c5-881d-1a782f407277", myParameters, "МСК_Количество слоев стекол"),
					m.SharedParameterFromGUIDName("f1cb5818-900b-407a-85a5-1aac4c0da623", myParameters, "МСК_Наименование газа-заполнителя камеры"),
					m.SharedParameterFromGUIDName("7b3af2b7-5b5f-4322-a588-545c0e3f1cac", myParameters, "МСК_Ламинирование"),
					m.SharedParameterFromGUIDName("288f2154-a879-410e-80ca-aaebe7a6b385", myParameters, "МСК_Армирование"),
					m.SharedParameterFromGUIDName("00d2d2cd-e493-4607-bfad-2e0842697c5a", myParameters, "МСК_Решетка"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("64bd484d-564b-4237-a18f-609bfd25923c", myParameters, "МСК_Тип окна"),
					m.SharedParameterFromGUIDName("56263316-6377-4426-b151-561a2325ac4a", myParameters, "МСК_Тип створок"),
					m.SharedParameterFromGUIDName("21b1ff0b-c5c5-446f-98de-79ed5a8696fd", myParameters, "МСК_Легкосбрасываемые"),
					m.SharedParameterFromGUIDName("96901a85-e765-4879-957c-b8fc4ff2858b", myParameters, "МСК_Противопожарная преграда"),
					"МСК_Тип противопожарной преграды",
					"Высота нижнего бруса",
					m.SharedParameterFromGUIDName("fdb653e9-69ff-4ece-ac52-33f354f8cc61", myParameters, "МСК_Светопропускание")
				};
				sb.Append("М_АР_09.1_Таблица заполнения параметров окон (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_АР_10.1.1_Таблица заполнения параметров лестниц (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					"Текущее количество подступенков",
					"Текущая высота подступенка",
					"Текущая ширина проступи",
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("787b4306-ad82-429c-b05a-f3879192c6bf", myParameters, "МСК_Форма марша"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_АР_10.1.1_Таблица заполнения параметров лестниц (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

	        // М_АР_10.2.1_Таблица заполнения параметров лестничных маршей (общая)
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRuns).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					"Текущее количество подступенков",
					"Текущее количество проступей",
					"Текущая ширина марша",
					"Текущая высота подступенка",
					"Текущая ширина проступи",
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("787b4306-ad82-429c-b05a-f3879192c6bf", myParameters, "МСК_Форма марша"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("4fd103c9-5ea4-4999-b484-ef3b9ae9188d", myParameters, "МСК_Вид деятельности"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_АР_10.2.1_Таблица заполнения параметров лестничных маршей (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

			/*
	         * М_АР_11.1_Таблица заполнения параметров пандусов (общая)
	         */
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ramps).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("4d902dad-86e1-4713-841a-a196798dbdf9", myParameters, "МСК_Предел огнестойкости"),
					m.SharedParameterFromGUIDName("411832d7-f92f-49ba-aba7-31ec7754e27c", myParameters, "МСК_Высота проезда"),
					m.SharedParameterFromGUIDName("006382d2-f621-41ac-9d24-8cbbe94d45ea", myParameters, "МСК_Доступность МГН"),
					m.SharedParameterFromGUIDName("886ac3f7-ee92-4498-bd58-8248e37001e6", myParameters, "МСК_Признак несущей конструкции"),
					"МСК_Путь эвакуации",
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("f93fdb2f-2c1c-4881-8a0c-6ee69316c64f", myParameters, "МСК_Форма пандуса"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_АР_11.1_Таблица заполнения параметров пандусов (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

			/*
	         * М_АР_12.1_Таблица для заполнения параметров ограждений (общая)
	         */
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					"Высота ограждения",
					m.SharedParameterFromGUIDName("9b679ab7-ea2e-49ce-90ab-0549d5aa37ff", myParameters, "МСК_Размер_Диаметр"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение"),
					m.SharedParameterFromGUIDName("1eaa8759-8cc8-44f3-9f69-08b454a09cc1", myParameters, "МСК_Код материала"),
					m.SharedParameterFromGUIDName("8b5e61a2-b091-491c-8092-0b01a55d4f45", myParameters, "МСК_Материал"),
					m.SharedParameterFromGUIDName("ad2dcb70-c648-46da-b8cc-00c088dd1653", myParameters, "МСК_Секция")
				};
				sb.Append("М_АР_12.1_Таблица для заполнения параметров ограждений (общая)\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			}

			/*
	         * М_АР_13_Таблица для заполнения параметров сборок
	         */
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Тип",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору"),
					m.SharedParameterFromGUIDName("647b5bc9-6570-416c-93d3-bd0d159775f2", myParameters, "МСК_Наименование"),
					m.SharedParameterFromGUIDName("b20d4101-069a-4467-aeac-3f2151cee025", myParameters, "МСК_Наружный"),
					m.SharedParameterFromGUIDName("fb118351-20b2-4f02-bd2c-99b27abaf8b2", myParameters, "МСК_Метод изготовления"),
					m.SharedParameterFromGUIDName("fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", myParameters, "МСК_Марка"),
					m.SharedParameterFromGUIDName("e7edd112-da46-46c3-886c-934dad841efb", myParameters, "МСК_Обозначение")
				};
				sb.Append("М_АР_13_Таблица для заполнения параметров сборок\n");
				sb.Append(m.RowHeader(pSet));
				sb.Append(m.RowElementsParameters(doc, pSet, elements));
			};


			/*
	         * М_КР(КЛ)_07_Таблица для классификации форм
	         */
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
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
			elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToList();
			if (elements != null)
			{
				string[] pSet = {
					"Семейство",
					"Тип",
					"Ключевая пометка",
					m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
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
			elements = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Where(x => x.Category.CategoryType == CategoryType.Model)
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
										 m.SharedParameterFromGUIDName("08257ef9-5429-48b1-aa0b-de2d9988ab0b", myParameters, "МСК_Код по классификатору")
					};

				sb.Append("М_АР(КЛ)_13_Таблица для классификации остальных категорий\n");
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
							//							sb.Append(getParameterValue(doc,el, "МСК_Код по классификатору", noData, noParameter)).Append("\t");
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
