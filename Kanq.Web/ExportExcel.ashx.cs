using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Kanq.Common;
using System.IO;

namespace Kanq.Web
{
    /// <summary>
    /// ExportExcel 的摘要说明
    /// </summary>
    public class ExportExcel : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            context.Response.AppendHeader("Content-Disposition", "attachment;filename=demo.xlsx");// 下载后的文件名
            MemoryStream ms = ExcelHelper.ExportExcel();
            context.Response.BinaryWrite(ms.ToArray());
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}