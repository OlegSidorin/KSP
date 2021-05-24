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
    public class Methods
    {
        public string noData = " ! ";
        public string noParameter = " - ";
        public int countAllMSKCod;
        public int countAll;
        public int countIfParameterIs;
        public int countIfMSKCOdIs;

        public List<MyParameter> AllParameters(Document doc)
        {
            List<string> mySharedParameters = new List<string>();
            var sharedParameters = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).ToList();
            foreach (var e in sharedParameters)
                mySharedParameters.Add(e.Name);
            List<MyParameter> myParameters = new List<MyParameter>();
            BindingMap bindings = doc.ParameterBindings;
            int n = bindings.Size;
            if (0 < n)
            {
                DefinitionBindingMapIterator it = bindings.ForwardIterator();
                while (it.MoveNext())
                {
                    Definition d = it.Key as Definition;
                    Binding b = it.Current as Binding;
                    MyParameter myParameter = 
                        new MyParameter("name", false, true); // Name, isShared, isInstance
                    if (d is InternalDefinition)
                        myParameter.Name = d.Name;
                    else
                        myParameter.Name += d.Name;
                    if (b is InstanceBinding)
                        myParameter.isInstance = true;
                    else
                        myParameter.isInstance = false;
                    foreach (var e in mySharedParameters)
                    {
                        if (d.Name == e)
                            myParameter.isShared = true;
                    }

                    myParameters.Add(myParameter);
                }
            }
            return myParameters;
        }
        public string getParameterNameFromGuid(Document doc, string st, string name)
        {
            string str = "";
            string path = Assembly.GetExecutingAssembly().Location;
            try
            {
                using (Transaction t = new Transaction(doc, "1"))
                {
                    t.Start();
                    Guid.TryParse(st, out Guid guid);
                    SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);
                    InternalDefinition def = sp.GetDefinition();
                    //ParameterType pt = def.ParameterType;
                    str = def.Name.ToString();
                    t.Commit();
                }

            }
            catch (Exception e)
            {
                str = "!" + "(" + name + ")";
            }


            return str;
        }

        public string GetElementParameterValue(Element el, string item, string noData, string noParameter)
        {
            if (item.Contains("!")) 
            {
                return String.Format("{0}", noParameter);
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
                                return String.Format("{0}", p.AsString());
                        }
                        else
                            return String.Format("{0}", p.AsValueString());
                    }
                    else
                        return String.Format("{0}", noData);
                }
                catch (Exception e)
                {
                    TaskDialog.Show("Warning", e.ToString());
                    return String.Format("{0}", noParameter);
                }
            }
        }

        public string GetParameterValue(Document doc, Element el, string item, string noData, string noParameter)
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

        public void MSKCounter(string parameter, string result, string noData, string noParameter)
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
    }
}
