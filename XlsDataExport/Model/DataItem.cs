using System.Collections.Generic;

namespace XlsDataExport.Model
{
    public class DataItem
    {
        public IList<HeaderItem> Header { get; set; }
        public IList<IDictionary<string, string>> Data { get; set; } 
    }
}