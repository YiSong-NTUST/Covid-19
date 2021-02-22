using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpertroApp
{
    class JSON_Obj_FW
    {
        public string ID;
        public string ELC;
        public string GNV;
        public string AGN;
        public string ROI;
        public List<string> WL;


       public  string MyDictionaryToJson(Dictionary<int, List<int>> dict)
        {
            var entries = dict.Select(d =>
                string.Format("\"{0}\": [{1}]", d.Key, string.Join(",", d.Value)));
            return "{" + string.Join(",", entries) + "}";
        }




    }
}
