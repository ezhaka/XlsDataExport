using XlsDataExport.Model;

namespace XlsDataExport
{
    public interface IExcelDataWriter
    {
        void Write(DataItem data, string fileName);
    }
}