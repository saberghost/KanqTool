using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Kanq.Common
{
    public class ExcelHelper
    {
        public static MemoryStream ExportExcel()
        {
            //创建Excel工作薄
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建单元格样式
            ICellStyle cellTitleStyle = workbook.CreateCellStyle();
            //设置单元格居中显示
            cellTitleStyle.Alignment = HorizontalAlignment.Center;
            //创建字体
            IFont font = workbook.CreateFont();
            //设置字体加粗显示
            font.IsBold = true;
            cellTitleStyle.SetFont(font);
            //创建Excel工作表
            ISheet sheet = workbook.CreateSheet("管理员");
            //创建Excel行
            IRow row = sheet.CreateRow(0);
            //创建Excel单元格
            ICell cell = row.CreateCell(0);
            //设置单元格值
            cell.SetCellValue("管理员管理");
            //设置单元格合并
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 3));
            cell.CellStyle = cellTitleStyle;
            //using (FileStream fs = new FileStream(@"C:\Users\Saber\Desktop\Demo.xlsx", FileMode.OpenOrCreate))
            //{
            //    workbook.Write(fs);
            //}
            MemoryStream ms;
            using (ms = new MemoryStream())
            {
                workbook.Write(ms);
            }
            return ms;
        }
    }
}
