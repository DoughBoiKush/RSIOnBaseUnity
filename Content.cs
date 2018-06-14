using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSIOnBaseUnity
{
    class Content
    {
        public string file { get; set; }
        public string documentID { get; set; }
        public string documentType { get; set; }
        public string documentTypeGroup { get; set; }
        public List<string> fileTypes { get; set; }
        public Dictionary<string, string> keywords { get; set; }
    }
}
