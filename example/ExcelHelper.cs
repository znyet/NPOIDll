using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Web;

namespace Common.Utils
{
    public class ExcelHelper
    {
        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        public static DataTable ToDataTable(IWorkbook workbook)
        {
            DataTable dt = new DataTable();
            ISheet sheet = workbook.GetSheetAt(0);
            //表头  
            IRow header = sheet.GetRow(sheet.FirstRowNum);
            List<int> columns = new List<int>();
            for (int i = 0; i < header.LastCellNum; i++)
            {
                object obj = GetValueType(header.GetCell(i));
                if (obj == null || obj.ToString() == string.Empty)
                {
                    dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                }
                else
                    dt.Columns.Add(new DataColumn(obj.ToString().Trim()));
                columns.Add(i);
            }
            //数据  
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                DataRow dr = dt.NewRow();
                bool hasValue = false;
                foreach (int j in columns)
                {
                    dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                    if (dr[j] != null && dr[j].ToString() != string.Empty)
                    {
                        dr[j] = dr[j].ToString().Trim();
                        hasValue = true;
                    }
                    else
                    {
                        dr[j] = "";
                    }
                }
                if (hasValue)
                {
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }


        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string file)
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                return ToDataTable(workbook);
            }
        }


        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(System.IO.Stream stream)
        {
            IWorkbook workbook = WorkbookFactory.Create(stream);
            return ToDataTable(workbook);
        }

        /// <summary>
        /// 导出xls
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="filename"></param>
        public static void ResponseExcelXls(IWorkbook workBook, string filename)
        {
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xls", filename));
            HttpContext.Current.Response.Clear();
            MemoryStream ms = new MemoryStream();
            workBook.Write(ms);
            ms.WriteTo(HttpContext.Current.Response.OutputStream);
            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// 导出xlsx
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="filename"></param>
        public static void ResponseExcelXlsx(IWorkbook workBook, string filename)
        {
            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xlsx", filename));
            HttpContext.Current.Response.Clear();
            MemoryStream ms = new MemoryStream();
            workBook.Write(ms);
            ms.WriteTo(HttpContext.Current.Response.OutputStream);
            HttpContext.Current.Response.End();
        }


    }
}
