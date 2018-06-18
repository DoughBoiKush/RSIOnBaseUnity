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
      
    class RSIKeywordType
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Required { get; set; }             
    }

    class RSIDocumentType
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public List<RSIKeywordType> KeywordTypes { get; set; }
    }

    class RSIDocumentTypeGroup
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public List<RSIDocumentType> DocumentTypes { get; set; }
    }
    
}
