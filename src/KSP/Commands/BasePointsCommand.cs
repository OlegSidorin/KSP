namespace KSP
{
    using System;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Linq;
    using System.Text;
    using OfficeOpenXml;
    using System.IO;
	using static KSP.CommonMethods;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class BasePointsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            ProjectLocation currentLocation = doc.ActiveProjectLocation;
            IList<BasePoint> bpList = new FilteredElementCollector(doc).OfClass(typeof(BasePoint)).Cast<BasePoint>().ToList();

            IList<ElementId> ids = new List<ElementId>();
            string str = "";
            string txtProjectPoint = "";
            string txtSurveyPoint = "";
            using (Transaction t = new Transaction(doc, "Some"))
            {
                t.Start();

                Double ang = 0;
                foreach (BasePoint basePoint in bpList)
                {
                    if (!basePoint.IsShared)
                        ang = -currentLocation.GetProjectPosition(basePoint.Position).Angle;
                }
                foreach (BasePoint basePoint in bpList)
                {
                    if (basePoint.IsShared)
                        txtSurveyPoint += BasePointsInfo(uiDoc, basePoint);
                    else
                        txtProjectPoint += BasePointsInfo(uiDoc, basePoint);

                    //CreateCross(uiDoc, new XYZ(basePoint.Position.X, basePoint.Position.Y, basePoint.Position.Z), 0, "");
                }

                //CreateCross(uiDoc, new XYZ(0, 0, 0), ang, "Zero");
                str = txtProjectPoint + "\n" + txtSurveyPoint;
				//writeToFile("C:\\Users\\Sidorin_O\\Documents\\TEST", doc.Title, str);
				if (!Directory.Exists(workingDir))
				{
					Directory.CreateDirectory(workingDir);
				}
				string filePathToExcel = workingDir + CropFileName(doc.Title) + " - базовые точки" + ".xlsx";
                string excelSheet = "БТ-" + CropFileName(doc.Title);


                WriteToExcel(filePathToExcel, excelSheet, str, "нет данных", "нет параметра");


                t.Commit();
            }
            return Result.Succeeded;
        }

		private string BasePointsInfo(UIDocument uiDoc, BasePoint bp)
		{
			Autodesk.Revit.DB.Document doc = uiDoc.Document;
			ProjectLocation currentLocation = doc.ActiveProjectLocation;

			string northSouthName = "С/Ю";
			string eastWestName = "В/З";
			string elevationName = "Отм";
			string trueNorthName = "Угол от ИС";
			string strFormat = "\t{0}\t{1}\t{2}\n";
			// для точки съемки Угла нет, а для базовой точки Долготы и Широты нет

			Double kRadToDegrees = 57.2957795;
			Double kFootsToMm = 304.80;

			string str = "";

			double northSouth, northSouthProject;
			double eastWest, eastWestProject;
			double elevation, elevationProject;
			double trueNorthProject;

			if (!bp.IsShared) // project point
			{
				str += "Файл:\t" + CropFileName(doc.Title) + ".rvt" + "\n";
				str += "Текущая площадка: " + doc.ActiveProjectLocation.Name + "\n\n";
				str += bp.Category.Name + "\n";
				//str += bp.Id + " \n";

				northSouthProject = currentLocation.GetProjectPosition(bp.Position).NorthSouth;
				northSouth = bp.Position.Y;
				eastWestProject = currentLocation.GetProjectPosition(bp.Position).EastWest;
				eastWest = bp.Position.X;
				elevationProject = currentLocation.GetProjectPosition(bp.Position).Elevation;
				elevation = bp.Position.Z;
				trueNorthProject = -currentLocation.GetProjectPosition(bp.Position).Angle;


				str += "Общие координаты" + " (SharedPosition):" + "\n"; // bp.SharedPosition
				str += String.Format(strFormat, northSouthName, northSouthProject * kFootsToMm, northSouthProject);
				str += String.Format(strFormat, eastWestName, eastWestProject * kFootsToMm, eastWestProject);
				str += String.Format(strFormat, elevationName, elevationProject * kFootsToMm, elevationProject);
				str += String.Format(strFormat, trueNorthName, trueNorthProject * kRadToDegrees, trueNorthProject);

				str += "Координаты в файле" + " (Position):" + "\n"; // bp.Position
				str += String.Format(strFormat, northSouthName, northSouth * kFootsToMm, northSouth);
				str += String.Format(strFormat, eastWestName, eastWest * kFootsToMm, eastWest);
				str += String.Format(strFormat, elevationName, elevation * kFootsToMm, elevation);

				//                IList<TextNoteType> tntList = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().ToList();
				//                TextNoteType tnt = tntList.FirstOrDefault();  
				//                TextNote tn = TextNote.Create(doc,uiDoc.ActiveView.Id, new XYZ(eastWest+1, northSouth-1, elevation), 0.7, str, tnt.Id);  
				return str;
			}
			else
			{
				str += bp.Category.Name + "\n";
				//str += bp.Id + " \n";

				northSouthProject = currentLocation.GetProjectPosition(bp.Position).NorthSouth;
				northSouth = bp.Position.Y;
				eastWestProject = currentLocation.GetProjectPosition(bp.Position).EastWest;
				eastWest = bp.Position.X;
				elevationProject = currentLocation.GetProjectPosition(bp.Position).Elevation;
				elevation = bp.Position.Z;

				str += "Общие координаты (SharedPosition):\n"; // bp.SharedPosition
				str += String.Format(strFormat, northSouthName, northSouthProject * kFootsToMm, northSouthProject);
				str += String.Format(strFormat, eastWestName, eastWestProject * kFootsToMm, eastWestProject);
				str += String.Format(strFormat, elevationName, elevationProject * kFootsToMm, elevationProject);

				str += "Координаты в файле (Position):\n"; // bp.Position
				str += String.Format(strFormat, northSouthName, northSouth * kFootsToMm, northSouth);
				str += String.Format(strFormat, eastWestName, eastWest * kFootsToMm, eastWest);
				str += String.Format(strFormat, elevationName, elevation * kFootsToMm, elevation);


				//                IList<TextNoteType> tntList = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().ToList();
				//                TextNoteType tnt = tntList.FirstOrDefault();  
				//                TextNote tn = TextNote.Create(doc,uiDoc.ActiveView.Id, new XYZ(eastWest+1, northSouth-1, elevation), 0.7, str, tnt.Id); 
				return str;

			}


		}
		private void WriteToExcel(string path, string excelSheet, string str, string noData, string noParameter)
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
				//int currentInt = -1;

				for (int i = 0; i < sb.Length; i++)
				{
					char ch = sb[i];
					if (ch.ToString() == "\n")
					{
						var currentCell = Convert.ToChar(65 + dCol).ToString() + (dRow).ToString();
						currentString = csb.Replace("\n", "").Replace("\t", "").ToString();
						if (Double.TryParse(currentString, out currentDouble))
							sheet.Cells[currentCell].Value = currentDouble;
						else
						{
							sheet.Cells[currentCell].Value = currentString;
							if (currentString.Contains("М_ИОС"))
							{

								sheet.Cells[dRow, 1, dRow, 20].Style.Font.Bold = true;
								sheet.Cells[dRow, 1, dRow, 20].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								sheet.Cells[dRow, 1, dRow, 20].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
								sheet.Cells[dRow, 1, dRow, 20].Style.Font.Color.SetColor(System.Drawing.Color.LightYellow);
								sheet.Cells[dRow + 1, 1, dRow + 1, 20].Style.Font.Bold = true;
								sheet.Cells[dRow + 1, 1, dRow + 1, 20].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, 20].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, 20].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								sheet.Cells[dRow + 1, 1, dRow + 1, 20].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							}
						}
						csb.Clear();
						dRow += 1;
						dCol = 0;
					}
					if (ch.ToString() == "\t")
					{
						var currentCell = Convert.ToChar(65 + dCol).ToString() + (dRow).ToString();
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
								sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
								sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.LightYellow);
							}
						}
						csb.Clear();
						dCol += 1;
					}
					else
					{
						csb.Append(ch);
					}

					//					if (ch.ToString() == "\t")
					//					{
					//						var currentCell = Convert.ToChar(65 + nextCol++).ToString() + (nextRow).ToString();
					//						var len = i - punkt.Last();
					//						if (len < 0) len = 0;
					//						currentString = sb.ToString(punkt.Last(), len);
					//						sheet.Cells[currentCell].Value = currentString;
					//						punkt.Add(i);
					//					}

				}


				//				for (int i = 0; i < rowNumber-1; i++) 
				//				{
				//					var currentCell = Convert.ToChar(66).ToString() + (i + 1).ToString();
				//					var len = row.ElementAt(i+1) - row.ElementAt(i);
				//					if (len < 0) len = 0;
				//					var isIn = 
				//					sheet.Cells[currentCell].Value = row.ElementAt(i).ToString() + ": " + sb.ToString(row.ElementAt(i), len);
				//				}

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
	}
}
