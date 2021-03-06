namespace KSP
{
    using Autodesk.Revit.DB;
    using OfficeOpenXml;
	using System;
	using System.Collections.Generic;
	using System.IO;
    using System.Linq;
    using System.Text;
	public static class CommonMethods
    {
		public static string user = @"C:\Users\" + Environment.UserName.ToString();
		public static string workingDir = user + @"\Documents\TEST\";
		public static int countIfParameterIs;
		public static int countIfMSKCOdIs;  // Счетчик: Прогресс классификации элементов % это пропорция заполнения параметра "МСК_Код по классификатору"
		public static int countAllMSKCod;   // а Прогресс заполнения параметров элементов модели % это все остальные
		public static int countAll;
		public static int readyOn;
		public static string noData = "Н/З";
		public static string noParameter = "Н/П";
		public static string CropFileName(string fileName)
		{
			string cropFileName = fileName;
			if (fileName.Contains("_отсоединено"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 12);
			if (fileName.ToLower().Contains("_sidorin_o"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 10);
			if (fileName.ToLower().Contains("_tyapkov_a"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 10);
			if (fileName.ToLower().Contains("_zavyalov_d"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 11);
			if (fileName.ToLower().Contains("_konovalov_d"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 12);
			if (fileName.ToLower().Contains("_altukhov_d"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 11);
			if (fileName.ToLower().Contains("_barinov_r"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 10);
			if (fileName.ToLower().Contains("_berdyugina_e"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 13);
			if (fileName.ToLower().Contains("_zvyagintsev_d"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 14);
			if (fileName.ToLower().Contains("_martynenko_d"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 13);
			if (fileName.Contains("-BIM"))
				cropFileName = cropFileName.Substring(0, cropFileName.Length - 4);
			return cropFileName;
		}
		private static string CurCell(int row, int col)
		{
			int let = 25;
			string curCell;
			if (col <= let)
				curCell = Convert.ToChar(65 + col).ToString() + (row).ToString();
			else
				curCell = Convert.ToChar(65).ToString() + Convert.ToChar(65 + col - let - 1).ToString() + (row).ToString();
			return curCell;
		}
		public static void WriteToExcel(string path, string excelSheet, string str, string noData, string noParameter, int columns)
		{
			using (var package = new ExcelPackage())
			{
				var sb = new StringBuilder(str);
				var csb = new StringBuilder();
				var file = new FileInfo(path);
				var sheet = package.Workbook.Worksheets.Add(excelSheet);
				var punkt = new List<int>();
				punkt.Add(0);
				int dRow = 1;
				int dCol = 0;
				string currentString = "";
				double currentDouble = -1;
				


				for (int i = 0; i < sb.Length; i++)
				{
					char ch = sb[i];
					if (ch.ToString() == "\n")
					{
						string currentCell = CurCell(dRow, dCol);
						currentString = csb.Replace("\n", "").Replace("\t", "").ToString();
						if (Double.TryParse(currentString, out currentDouble))
							sheet.Cells[currentCell].Value = currentDouble;
						else
						{
							sheet.Cells[currentCell].Value = currentString;
							if (currentString.Contains("М_ИОС") || currentString.Contains("М_КР") || currentString.Contains("М_АР"))
							{

								sheet.Cells[dRow, 1, dRow, columns].Style.Font.Bold = true;
								sheet.Cells[dRow, 1, dRow, columns].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								sheet.Cells[dRow, 1, dRow, columns].Style.Fill.BackgroundColor.SetColor(100, 105, 124, 231);
								sheet.Cells[dRow, 1, dRow, columns].Style.Font.Color.SetColor(System.Drawing.Color.LightYellow);
								sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Font.Bold = true;
								sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							}
						}
						csb.Clear();
						dRow += 1;
						dCol = 0;
					}
					if (ch.ToString() == "\t")
					{
						string currentCell = CurCell(dRow, dCol);
						currentString = csb.Replace("\n", "").Replace("\t", "").ToString();

						if (Double.TryParse(currentString, out currentDouble))
						{
							sheet.Cells[currentCell].Value = currentDouble;
							//sheet.Cells[currentCell].Style.Font.Italic = true;
							//sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
							//sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aqua);
							//sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.Green);


						}
						else
						{
							sheet.Cells[currentCell].Value = currentString;
							if (currentString == noData)
							{
								sheet.Cells[currentCell].Style.Font.Bold = true;
								sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.Red);
							}
							if (currentString == noParameter)
							{
								sheet.Cells[currentCell].Style.Font.Bold = true;
								sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
								sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
								sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
							}
						}
						csb.Clear();
						dCol += 1;
					}
					else
					{
						csb.Append(ch);
					}

				}


				/* справочно
				 * '\u0065' A (char)
					" " (ASCII 32 (0x20)), обычный пробел.
					"\t" (ASCII 9 (0x09)), символ табуляции.
					"\n" (ASCII 10 (0x0A)), символ перевода строки.
					"\r" (ASCII 13 (0x0D)), символ возврата каретки.
					"\0" (ASCII 0 (0x00)), NUL-байт.
					"\x0B" (ASCII 11 (0x0B)), вертикальная табуляция.
				*/

				//sheet.Cells["A1"].Value = sb.ToString();

				//sheet.Cells["A2"].Value = row.Count;
				//sheet.Cells["A3"].Value = col.Count;
				//sheet.Cells["A4"].Value = col.First();
				//col.RemoveAt(0);
				//sheet.Cells["A5"].Value = col.First();

				// Save to file
				package.SaveAs(file);
			}
		}
		public static void WriteToFile(string dir, string name, string txt)
		{
			string fileName = dir + "\\" + name + ".txt";
			if (File.Exists(fileName))
			{
				File.Delete(fileName);
			}
			using (StreamWriter writer = new StreamWriter(fileName, true))
			{
				writer.Write(txt);
			}

		}
		public static void MSKCounter(string parameter, string result, string noData, string noParameter)
		{
			if (parameter == "МСК_Код по классификатору")
				countAllMSKCod += 1;

			if ((result == noData) || (result == noParameter))
			{
				countAll += 1;
			}
			else
			{
				countAll += 1;
				countIfParameterIs += 1;
				if (parameter == "МСК_Код по классификатору")
					countIfMSKCOdIs += 1;
			}
		}
		public static string GetParameterValue(Document doc, Element el, string item, string noData, string noParameter)
		{
			string str = GetElementParameterValue(el, item, noData, noParameter).Replace("\t", " ").Replace("\n", " ");

			if (str == noParameter)
			{
				if (null == el.GetTypeId())
					return str;
				else
				{
					Element elType = doc.GetElement(el.GetTypeId());
					return GetElementParameterValue(elType, item, noData, noParameter).Replace("\t", " ").Replace("\n", " ");
				}

			}
			else
				return str;
		}
		public static string GetElementParameterValue(Element el, string item, string noData, string noParameter)
		{
			try
			{
				Parameter p = el.ParametersMap.get_Item(item);
				if (p.HasValue)
				{
					if (p.StorageType == StorageType.String)
					{
						if (p.AsString() == "")
							return String.Format("{0}", noData);
						else
							return String.Format("{0}", p.AsString());
					}
					else
						return String.Format("{0}", p.AsValueString());
				}
				else
					return String.Format("{0}", noData);
			}
			catch
			{
				return String.Format("{0}", noParameter);
			}
		}
		public static string RowHeader(string[] pSet)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < pSet.Count(); i++)
			{
				sb.Append(pSet[i]).Append("\t");
			}
			sb.Append("\n");
			return sb.ToString();
		}
		public static string RowElementsParameters(Document doc, string[] pSet, IList<Element> iList, string noData, string noParameter)
		{
			StringBuilder sb = new StringBuilder();
			foreach (var el in iList)
			{
				sb.Append(el.Name).Append("\t");
				for (int i = 1; i < pSet.Count(); i++)
				{
					string result = GetParameterValue(doc, el, pSet[i], noData, noParameter);
					sb.Append(result).Append("\t");
					MSKCounter(pSet[i], result, noData, noParameter);
				}
				sb.Append("\n");
			}
			sb.Append("\n");
			return sb.ToString();
		}
		public static string RowElementsParametersWCatNameWFamily(Document doc, string[] pSet, IList<Element> iList, string noData, string noParameter)
		{
			StringBuilder sb = new StringBuilder();
			foreach (var el in iList)
			{
				sb.Append(el.Name).Append("\t");
				sb.Append(el.Category.Name).Append("\t");
				FamilyInstance fi = el as FamilyInstance;
				try
				{
					if (null == fi.GetTypeId())
					{
						//
					}
					else
					{
						FamilySymbol fis = doc.GetElement(fi.GetTypeId()) as FamilySymbol;
						sb.Append(fis.FamilyName).Append("\t");
					}

				}
				catch
				{
					sb.Append(noParameter).Append("\t");
				}
				for (int i = 3; i < pSet.Count(); i++)
				{
					string result = GetParameterValue(doc, el, pSet[i], noData, noParameter);
					sb.Append(result).Append("\t");
					MSKCounter(pSet[i], result, noData, noParameter);
				}
				sb.Append("\n");
			}
			sb.Append("\n");
			return sb.ToString();
		}
	}
}
