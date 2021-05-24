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
    class GetParameterTypeAndNameFromGUID
    {
        public string Name(Document doc, string st, string name)
        {
            string str = "";
            try
            {
                using (Transaction t = new Transaction(doc, "1"))
                {
                    t.Start();
                    Guid.TryParse(st, out Guid guid);
                    SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);
                    InternalDefinition def = sp.GetDefinition();
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
        public string Type(Document doc, string st, string name)
        {
            string str = "";
            try
            {
                using (Transaction t = new Transaction(doc, "1"))
                {
                    t.Start();
                    Guid.TryParse(st, out Guid guid);
                    SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);
                    InternalDefinition def = sp.GetDefinition();
                    BindingMap bindings = doc.ParameterBindings;
                    Binding b = bindings.get_Item(def);
                    if (b is InstanceBinding)
                        str = "inst";
                    else
                        str = "type";
                    t.Commit();
                }

            }
            catch (Exception e)
            {
                str = "uknown";
            }
            return str;
        }
    }
}
