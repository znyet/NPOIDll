using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace testNpoi
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e) //读取2003或2007 Excel
        {
            IWorkbook wb = WorkbookFactory.Create(FileUpload1.PostedFile.InputStream);

            ISheet sheet = wb.GetSheetAt(0);

            var rows = sheet.GetEnumerator();

            while (rows.MoveNext())
            {
                IRow row=rows.Current as IRow;
                List<ICell> cells = row.Cells;
                if (row != null && cells != null)
                {
                    Response.Write(row.RowNum);
                    foreach (var cell in cells)
                    {
                        Response.Write(cell.ToString());
                        Response.Write("<br/>");
                    }
                }
            }
        }

        protected void Button2_Click(object sender, EventArgs e) //导出2003格式
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue("the value");
                }
            }

            MemoryStream file = new MemoryStream();
            hssfworkbook.Write(file);

            string filename = "test.xls";
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            file.WriteTo(Response.OutputStream);
            Response.End();

        }

        protected void Button3_Click(object sender, EventArgs e)//导出2007格式
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue("the value");
                }
            }

            string filename = "test.xlsx";

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));

            using (var f = File.Create(@"c:\test.xlsx"))
            {
                workbook.Write(f);
            }
            Response.WriteFile(@"c:\test.xlsx");
            Response.Flush();
            Response.End();
        }
    }
}