namespace KSP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    public class MyParameter
    {
        public string GuidValue { get; set; }
        public string Name { get; set; }
        public bool isShared { get; set; }
        public bool isInstance { get; set; }

        public MyParameter(string gV,  string n, bool iS, bool iI)
        {
            GuidValue = gV;
            Name = n;
            isShared = iS;
            isInstance = iI;
        }
    }

    public class MySharedParameter
    {
        public string GuidValue { get; set; }
        public string Name { get; set; }
        public MySharedParameter(string gV, string n)
        {
            GuidValue = gV;
            Name = n;
        }
    }

}
