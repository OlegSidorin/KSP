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
    public class MyMethods
    {
        public string noData = " [НЗ] ";
        public string noParameter = " [НП] ";
        public string noCategory = " [НПкЭК] ";
        public int countAllMSKCod;
        public int countAll;
        public int countIfParameterIs;
        public int countIfMSKCOdIs;

        public List<MyParameter> AllParameters(Document doc)
        {
            List<MySharedParameter> mySharedParameters = new List<MySharedParameter>();
            var sharedParameters = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();
            foreach (var e in sharedParameters)
            {
                InternalDefinition iDef = e.GetDefinition();
                Definition def = e.GetDefinition();
                if (def != null)
                {
                    MySharedParameter msp = new MySharedParameter(e.GuidValue.ToString(), iDef.Name);
                    mySharedParameters.Add(msp);
                }
            }
                
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
                        new MyParameter("", "name", false, true); // Name, isShared, isInstance
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
                        if (myParameter.Name == e.Name)
                        {
                            myParameter.isShared = true;
                            myParameter.GuidValue = e.GuidValue;
                        }   
                    }

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

        public string GetParameterValue(Document doc, Element el, string item, string noData, string noParameter)
        {
            if (item.Contains("!!")) 
            {
                return String.Format("{0}", noParameter);
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
                                return String.Format("{0}", p.AsString());
                        }
                        else
                            return String.Format("{0}", p.AsValueString());
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
                                return String.Format("{0}", p.AsString());
                        }
                        else
                            return String.Format("{0}", p.AsValueString());
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
                                return String.Format("{0}", p.AsString());
                        }
                        else
                            return String.Format("{0}", p.AsValueString());
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
                                    return String.Format("{0}", p.AsString());
                            }
                            else
                                return String.Format("{0}", p.AsValueString());
                        }
                        else
                            return String.Format("{0}", noData);
                    }
                    catch (Exception)
                    {
                        //TaskDialog.Show("Warning", e1.ToString());
                        return String.Format("{0}", noCategory);
                    }
                }
            }
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
