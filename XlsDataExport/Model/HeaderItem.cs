using System.Collections.Generic;

namespace XlsDataExport.Model
{
    public class HeaderItem
    {
        public string Title { get; set; } 
        public string Dataindex { get; set; } 
        public IList<HeaderItem> Childs { get; set; } 
    }
}