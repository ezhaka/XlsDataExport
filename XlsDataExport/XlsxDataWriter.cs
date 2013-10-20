using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using XlsDataExport.Model;

namespace XlsDataExport
{
    public class XlsxDataWriter : IExcelDataWriter
    {
        public void Write(DataItem data, string fileName)
        {
            FileInfo newFile = new FileInfo(fileName);

            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(fileName);
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Main");

                int headerHeight = 1;
                GetHeaderHeight(data.Header, 1, ref headerHeight);

                IDictionary<string, int> indexToColumn = new Dictionary<string, int>();
                int headerWidth = WriteHeader(worksheet, data.Header, 1, 1, headerHeight, indexToColumn);
                WriteData(worksheet, data.Data, headerHeight + 1, indexToColumn);
                worksheet.Cells[headerHeight, 1, headerHeight + data.Data.Count, headerWidth].AutoFilter = true;

                package.Save();
            }
        }

        private int WriteHeader(ExcelWorksheet worksheet, IList<HeaderItem> headerItems, int row, int column, int headerHeight, IDictionary<string, int> indexToColumn)
        {
            int widthsSum = 0;

            foreach (HeaderItem headerItem in headerItems)
            {
                worksheet.Cells[row, column].Value = headerItem.Title;

                int width = 1;
                ExcelRange headerRange;

                if (headerItem.Childs != null && headerItem.Childs.Any())
                {
                    width = WriteHeader(worksheet, headerItem.Childs, row + 1, column, headerHeight, indexToColumn);
                    headerRange = worksheet.Cells[row, column, row, column + width - 1];
                    ApplyHeaderCellStyle(headerRange);
                }
                else
                {
                    headerRange = worksheet.Cells[row, column, headerHeight, column];
                    ApplyHeaderCellStyle(headerRange);

                    indexToColumn.Add(headerItem.Dataindex, column);
                }

                column += width;
                widthsSum += width;
            }

            return widthsSum;
        }

        private void ApplyHeaderCellStyle(ExcelRange headerRange)
        {
            Color headerColor = Color.FromArgb(198, 224, 180);
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(headerColor);

            headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);

            headerRange.Merge = true;
        }

        private void GetHeaderHeight(IList<HeaderItem> headerItems, int row, ref int maxLevel)
        {
            foreach (HeaderItem headerItem in headerItems)
            {
                if (headerItem.Childs != null && headerItem.Childs.Any())
                {
                    GetHeaderHeight(headerItem.Childs, row + 1, ref maxLevel);
                }
            }

            if (row > maxLevel)
            {
                maxLevel = row;
            }
        }

        private void WriteData(ExcelWorksheet worksheet, IList<IDictionary<string, string>> data, int row, IDictionary<string, int> indexToColumn)
        {
            foreach (var dataItem in data)
            {
                foreach (KeyValuePair<string, string> indexToValue in dataItem)
                {
                    int column = indexToColumn[indexToValue.Key];
                    worksheet.Cells[row, column].Value = indexToValue.Value;
                }

                row++;
            }
        }
    }
}