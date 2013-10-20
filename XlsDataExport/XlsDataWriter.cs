using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using XlsDataExport.Model;

namespace XlsDataExport
{
    public class XlsDataWriter : IExcelDataWriter
    {
        public void Write(DataItem data, string fileName)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Main");
            IDictionary<string, int> indexToColumn = new Dictionary<string, int>();

            int headerHeight = 1;
            this.GetHeaderHeight(data.Header, 1, ref headerHeight);

            int headerWidth = WriteHeader(worksheet, data.Header, 0, 0, headerHeight, indexToColumn);
            this.WriteData(worksheet, data.Data, headerHeight, indexToColumn);

            worksheet.SetAutoFilter(new CellRangeAddress(headerHeight - 1, headerHeight + data.Data.Count, 0, headerWidth - 1));

            using (var fileData = new FileStream(fileName, FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }

        private int WriteHeader(ISheet worksheet, IList<HeaderItem> headerItems, int rowIndex, int columnIndex, int headerHeight, IDictionary<string, int> indexToColumn)
        {
            int widthsSum = 0;
            IRow row = worksheet.GetRow(rowIndex) ?? worksheet.CreateRow(rowIndex);

            foreach (HeaderItem headerItem in headerItems)
            {
                ICell cell = row.CreateCell(columnIndex);
                cell.SetCellValue(headerItem.Title);

                int width = 1;
                CellRangeAddress cellRange;

                if (headerItem.Childs != null && headerItem.Childs.Any())
                {
                    width = WriteHeader(worksheet, headerItem.Childs, rowIndex + 1, columnIndex, headerHeight, indexToColumn);

                    cellRange = new CellRangeAddress(rowIndex, rowIndex, columnIndex, columnIndex + width - 1);
                }
                else
                {
                    cellRange = new CellRangeAddress(rowIndex, headerHeight - 1, columnIndex, columnIndex);
                    indexToColumn.Add(headerItem.Dataindex, columnIndex);
                }

                worksheet.AddMergedRegion(cellRange);
                cell.CellStyle = this.GetHeaderCellStyle((HSSFWorkbook)worksheet.Workbook);
                this.AddBorderToRegion(cellRange, (HSSFSheet)worksheet);

                columnIndex += width;
                widthsSum += width;
            }

            return widthsSum;
        }

        private void WriteData(ISheet worksheet, IList<IDictionary<string, string>> data, int rowIndex, IDictionary<string, int> indexToColumn)
        {
            foreach (var dataItem in data)
            {
                IRow row = worksheet.CreateRow(rowIndex);

                foreach (KeyValuePair<string, string> indexToValue in dataItem)
                {
                    int column = indexToColumn[indexToValue.Key];
                    ICell cell = row.CreateCell(column);
                    cell.SetCellValue(indexToValue.Value);
                    cell.CellStyle = this.GetBodyCellStyle((HSSFWorkbook) worksheet.Workbook);
                }

                rowIndex++;
            }
        }

        private void GetHeaderHeight(IList<HeaderItem> headerItems, int rowIndex, ref int maxLevel)
        {
            foreach (HeaderItem headerItem in headerItems)
            {
                if (headerItem.Childs != null && headerItem.Childs.Any())
                {
                    GetHeaderHeight(headerItem.Childs, rowIndex + 1, ref maxLevel);
                }
            }

            if (rowIndex > maxLevel)
            {
                maxLevel = rowIndex;
            }
        }

        private void AddBorderToRegion(CellRangeAddress range, HSSFSheet worksheet)
        {
            HSSFWorkbook workbook = (HSSFWorkbook) worksheet.Workbook;

            HSSFRegionUtil.SetBorderBottom(BorderStyle.THIN, range, worksheet, workbook);
            HSSFRegionUtil.SetBorderLeft(BorderStyle.THIN, range, worksheet, workbook);
            HSSFRegionUtil.SetBorderRight(BorderStyle.THIN, range, worksheet, workbook);
            HSSFRegionUtil.SetBorderTop(BorderStyle.THIN, range, worksheet, workbook);
        }

        private ICellStyle GetBodyCellStyle(HSSFWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.BorderTop = BorderStyle.THIN;
            cellStyle.BorderBottom = BorderStyle.THIN;
            cellStyle.BorderLeft = BorderStyle.THIN;
            cellStyle.BorderRight = BorderStyle.THIN;

            return cellStyle;
        }

        private ICellStyle GetHeaderCellStyle(HSSFWorkbook workbook)
        {
            HSSFPalette palette = workbook.GetCustomPalette();
            const short fillForegroundColor = 10;
            palette.SetColorAtIndex(fillForegroundColor, 198, 224, 180);

            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            headerCellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            headerCellStyle.FillForegroundColor = fillForegroundColor;
            headerCellStyle.Alignment = HorizontalAlignment.CENTER;

            return headerCellStyle;
        }
    }
}