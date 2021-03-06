namespace KSP
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;
    using System.Text;
    using System.Linq;
    using System.Reflection;
    using OfficeOpenXml;
    using System.Diagnostics;

    public class MyMethods
    {
        public string user = @"C:\Users\" + Environment.UserName.ToString();
        public string workingDir = @"C:\Users\" + Environment.UserName.ToString() + @"\Documents\TESTS\";
        public string noData = " [НЗ] ";
        public string noParameter = " [НП] ";
        public string noCategory = " [НК] ";
        public int countAllMSKCod;
        public int countAll;
        public int countIfParameterIs;
        public int countIfMSKCOdIs;
        public int readyOn;
        

        public List<MyParameter> AllParameters(Document doc)
        {
            //List<MySharedParameter> mySharedParameters = new List<MySharedParameter>();
            var sharedParameters = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();
            //var parameters = new FilteredElementCollector(doc).OfClass(typeof(ParameterElement)).Cast<ParameterElement>().ToList();

            BindingMap bindingMap = doc.ParameterBindings;
            int size = bindingMap.Size;


            //foreach (var e in sharedParameters)
            //{
            //    InternalDefinition iDef = e.GetDefinition();
            //    Definition def = e.GetDefinition();
            //    if (def != null)
            //    {
            //        MySharedParameter msp = new MySharedParameter(e.GuidValue.ToString(), def.Name);
            //        mySharedParameters.Add(msp);
            //    }
            //}
                
            List<MyParameter> myParameters = new List<MyParameter>();
            BindingMap bindings = doc.ParameterBindings;
            int n = bindings.Size;
            if (0 < n)
            {
                DefinitionBindingMapIterator it = bindings.ForwardIterator();
                while (it.MoveNext())
                {
                    Definition d = it.Key as Definition;
                    InternalDefinition id = it.Key as InternalDefinition;
                    Binding b = it.Current as Binding;
                    MyParameter myParameter = 
                        new MyParameter("", "", "name", false, true); // Name, isShared, isInstance
                    myParameter.Name = d.Name;
                    myParameter.Id = id.Id.ToString();
                    if (b is InstanceBinding)
                        myParameter.isInstance = true;
                    else
                        myParameter.isInstance = false;
                    foreach (var sp in sharedParameters)
                    {
                        if (myParameter.Id == sp.Id.ToString())
                        {
                            myParameter.isShared = true;
                            myParameter.GuidValue = sp.GuidValue.ToString();
                        }   
                    }

                    //foreach (var p in parameters)
                    //{
                    //    if (myParameter.Id == p.Id.ToString())
                    //    {
                    //        myParameter.isShared = false;
                    //        myParameter.GuidValue = "---"; //p.GuidValue.ToString();
                    //    }
                    //}

                    myParameters.Add(myParameter);
                }
            }
            return myParameters;
        }
        public string SharedParameterFromGUIDName(string stguid, List<MyParameter> allParameters, string stname)
        {
            string s = "!!(" + stname + ")";
            foreach(var e in allParameters)
            {
                if (e.GuidValue == stguid)
                {
                    if (e.isInstance)
                        s = e.Name + "<I>";
                    else
                        s = e.Name + "<T>";
                } 
                else if (e.Name == stname)
                {
                    s = "??(" + e.Name + ")";
                }
                    
            }
            return s;
        }
        public bool SharedParameterFromGUIDIsInstance(string stguid, List<MyParameter> allParameters, string stname)
        {
            bool b = true;
            foreach (var e in allParameters)
            {
                if (e.GuidValue == stguid)
                    b = e.isInstance;
            }
            return b;
        }
        public string GetParameterValue(Document doc, Element el, string item)
        {
            if (item.Contains("!!")) 
            {
                return String.Format("{0}", noParameter);
            }
            else if (item.Contains("??"))
            {
                return String.Format("{0}", noCategory);
            }
            else if (item.Contains("Значение Кода"))
            {
                return String.Format("{0}", "");
            }
            else if (item.Contains("<I>"))
            {
                try
                {
                    Parameter p = el.ParametersMap.get_Item(item.Replace("<I>", ""));
                    if (p.HasValue)
                    {
                        if (p.StorageType == StorageType.String)
                        {
                            if (p.AsString() == "")
                                return String.Format("{0}", noData);
                            else
                                return String.Format("{0}", p.AsString().Replace("\n", " "));
                        }
                        else
                            return String.Format("{0}", p.AsValueString().Replace("\n", " "));
                    }
                    else
                        return String.Format("{0}", noData);
                }
                catch (Exception)
                {
                    //TaskDialog.Show("Warning", e.ToString());
                    return String.Format("{0}", noCategory);
                }
            }
            else if (item.Contains("<T>"))
            {
                Element elType = doc.GetElement(el.GetTypeId());
                try
                {
                    Parameter p = elType.ParametersMap.get_Item(item.Replace("<T>", ""));
                    if (p.HasValue)
                    {
                        if (p.StorageType == StorageType.String)
                        {
                            if (p.AsString() == "")
                                return String.Format("{0}", noData);
                            else
                                return String.Format("{0}", p.AsString().Replace("\n", " "));
                        }
                        else
                            return String.Format("{0}", p.AsValueString().Replace("\n", " "));
                    }
                    else
                        return String.Format("{0}", noData);
                }
                catch (Exception)
                {
                    //TaskDialog.Show("Warning", e.ToString());
                    return String.Format("{0}", noCategory);
                }
            }
            else
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
                                return String.Format("{0}", p.AsString().Replace("\n", " "));
                        }
                        else
                            return String.Format("{0}", p.AsValueString().Replace("\n", " "));
                    }
                    else
                        return String.Format("{0}", noData);
                }
                catch (Exception)
                {
                    //TaskDialog.Show("Warning", e.ToString());
                    Element elType = doc.GetElement(el.GetTypeId());
                    try
                    {
                        Parameter p = elType.ParametersMap.get_Item(item);
                        if (p.HasValue)
                        {
                            if (p.StorageType == StorageType.String)
                            {
                                if (p.AsString() == "")
                                    return String.Format("{0}", noData);
                                else
                                    return String.Format("{0}", p.AsString().Replace("\n", " "));
                            }
                            else
                                return String.Format("{0}", p.AsValueString().Replace("\n", " "));
                        }
                        else
                            return String.Format("{0}", noData);
                    }
                    catch (Exception)
                    {
                        //TaskDialog.Show("Warning", e1.ToString());
                        return String.Format("{0}", noParameter);
                    }
                }
            }
        }
        public string RowHeader(string[] pSet)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < pSet.Count(); i++)
            {
                sb.Append(pSet[i].Replace("<I>", "").Replace("<T>", "")).Append("\t");
            }
            sb.Append("\n");
            return sb.ToString();
        }
        public string RowElementsParameters(Document doc, string[] pSet, IList<Element> iList)
        {
            MyMSK myMSK = new MyMSK();
            StringBuilder sb = new StringBuilder();
            foreach (var el in iList)
            {
                sb.Append(el.Name).Append("\t");
                for (int i = 1; i < pSet.Count(); i++)
                {
                    string result = GetParameterValue(doc, el, pSet[i]);

                    if (pSet[i].Contains("МСК_Код по классификатору"))
                        sb.Append(result).Append("\t").Append(myMSK.getMyMSK(result));
                    else
                        sb.Append(result).Append("\t");

                    MSKCounter(pSet[i], result);

                }
                sb.Append("\n");
            }
            sb.Append("\n");
            return sb.ToString();
        }
        public string RowElementsParametersWCatNameWFamily(Document doc, string[] pSet, IList<Element> iList)
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
                    string result = GetParameterValue(doc, el, pSet[i]);
                    sb.Append(result).Append("\t");
                    MSKCounter(pSet[i], result);
                }
                sb.Append("\n");
            }
            sb.Append("\n");
            return sb.ToString();
        }
        public void MSKCounter(string parameter, string result)
        {
            if (parameter.Contains("МСК_Код по классификатору"))
                countAllMSKCod += 1;

            if ((result == noData) || (result == noParameter) || (result == noCategory))
            {
                countAll += 1;
            }
            else
            {
                countAll += 1;
                countIfParameterIs += 1;
                if (parameter.Contains("МСК_Код по классификатору"))
                    countIfMSKCOdIs += 1;
            }
        }
        string CurCell(int row, int col)
        {
            int let = 25;
            string curCell;
            if (col <= let)
                curCell = Convert.ToChar(65 + col).ToString() + (row).ToString();
            else
                curCell = Convert.ToChar(65).ToString() + Convert.ToChar(65 + col - let - 1).ToString() + (row).ToString();
            return curCell;
        }
        public void WriteToExcel(string path, string excelSheet, string str, string noData, string noParameter, int columns)
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
                List<int> groupMarkers = new List<int>();



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
                                groupMarkers.Add(dRow);
                            }
                            if (currentString.Contains("Отчет") || currentString.Contains("Условные"))
                            {
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Bold = true;
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
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
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                //sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                //sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightPink);
                            }
                            if (currentString == noCategory)
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                //sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                //sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCyan);
                            }
                            if (currentString.Contains("!!"))
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }
                            if (currentString.Contains("??"))
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightDown;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            }
                            if (currentString.Contains("МСК_Код по классификатору") || currentString.Contains("Значение Кода"))
                            {
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
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
                
                groupMarkers.Add(dRow - 7); // 7 это вычет из-за Условных обозначений

                int lastMarker = groupMarkers.Count;

                for (var n = 0; n < lastMarker - 1; n++)
                { 
                    for (var i = groupMarkers[n] + 2; i < groupMarkers[n + 1] - 1; i++)
                    {
                        sheet.Row(i).OutlineLevel = 1;
                        sheet.Row(i).Collapsed = true;
                    }
                }
                package.SaveAs(file);
            }
        }
        public void WriteToFile(string dir, string name, string txt)
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
        public string CropFileName(string fileName)
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
        public void OpenFolder(string folderPath)
        {
            
            if (Directory.Exists(folderPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = folderPath,
                    FileName = "explorer.exe",
                };
                var pr = Process.GetProcessesByName("explorer").Where(xxx => xxx.MainWindowTitle.Contains("TESTS")).ToList();
                string ss = "";
                if (pr != null)
                {
                    //var str = pr.Count.ToString();
                    //TaskDialog.Show("Warning", str);
                    foreach (var p in pr)
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }
                        
                    }
                }
                //TaskDialog.Show("Process", ss);
                Process.Start(startInfo);
            }
            else
            {
                TaskDialog.Show("Warning", String.Format("{0} не существует", folderPath));
            }
        }
    }
}
