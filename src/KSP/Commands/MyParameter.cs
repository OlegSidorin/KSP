namespace KSP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    public class MyParameter
    {
        public string Name { get; set; }
        public bool isShared { get; set; }
        public bool isInstance { get; set; }

        public MyParameter(string n, bool iS, bool iI)
        {
            Name = n;
            isShared = iS;
            isInstance = iI;
        }
    }
}
