using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace NPOI_demo
{
    public class NPOIHelper
    {
        public static void ExportExcel(DataTable dt, string fileName, string headerText)
        {
            // 确定文件格式
            string fileExt = Path.GetExtension(fileName).ToLower();
            IWorkbook workbook;
            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                throw new ArgumentException("Unsupported file extension. The file must be either .xls or .xlsx.");
            }

            // 创建Excel样式
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            // 创建列头样式
            IFont headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerFont.FontHeightInPoints = 20;
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.SetFont(headerFont);
            headerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            // 计算需要的工作表数量
            int totalRows = dt.Rows.Count;
            int sheetsNeeded = (totalRows + 65535) / 65536;

            // 创建工作表并设置列头
            for (int sheetIndex = 0; sheetIndex < sheetsNeeded; sheetIndex++)
            {
                ISheet sheet = workbook.CreateSheet($"Sheet{sheetIndex + 1}");
                IRow headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue(headerText);
                headerRow.GetCell(0).CellStyle = headerStyle;
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dt.Columns.Count - 1));

                // 设置列头
                IRow headerColumnRow = sheet.CreateRow(1);
                int index = 0;
                foreach (DataColumn column in dt.Columns)
                {
                    headerColumnRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                    headerColumnRow.GetCell(column.Ordinal).CellStyle = headerStyle;
                    //sheet.SetColumnWidth(column.Ordinal, (Encoding.UTF8.GetBytes(column.ColumnName).Length + 1) * 256);
                    sheet.AutoSizeColumn(index);
                    index++;
                }
                index = 0;
                // 填充数据
                int startRow = sheetIndex * 65536;
                int endRow = Math.Min(startRow + 65535, totalRows);
                for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex - startRow + 2);
                    foreach (DataColumn column in dt.Columns)
                    {
                        ICell newCell = dataRow.CreateCell(column.Ordinal);
                        string drValue = dt.Rows[rowIndex][column].ToString();
                        SetCellValue(newCell, drValue, column.DataType.ToString(), dateStyle);
                    }
                }
            }

            // 写入文件
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            // 关闭工作簿资源
            workbook.Close();
        }

        private static void SetCellValue(ICell cell, string value, string dataType, ICellStyle dateStyle)
        {
            switch (dataType)
            {
                case "System.String":
                    if (isNumeric(value, out double numericResult))
                    {
                        cell.SetCellValue(numericResult);
                    }
                    else
                    {
                        cell.SetCellValue(value);
                    }
                    break;
                case "System.DateTime":
                    if (DateTime.TryParse(value, out DateTime date))
                    {
                        cell.SetCellValue(date);
                        cell.CellStyle = dateStyle;
                    }
                    break;
                case "System.Boolean":
                    if (bool.TryParse(value, out bool boolValue))
                    {
                        cell.SetCellValue(boolValue);
                    }
                    break;
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    if (int.TryParse(value, out int intValue))
                    {
                        cell.SetCellValue(intValue);
                    }
                    break;
                case "System.Decimal":
                case "System.Double":
                    if (double.TryParse(value, out double doubleValue))
                    {
                        cell.SetCellValue(doubleValue);
                    }
                    break;
                case "System.DBNull":
                    cell.SetCellValue("");
                    break;
                default:
                    cell.SetCellValue("");
                    break;
            }
        }

        private static bool isNumeric(string? drValue, out double result)
        {
            Regex rex = new Regex(@"^[-]?\d+[.]?\d*$");
            result = -1;
            if (rex.IsMatch(drValue))
            {
                result = double.Parse(drValue);
                return true;
            }
            return false;
        }
    }
}
