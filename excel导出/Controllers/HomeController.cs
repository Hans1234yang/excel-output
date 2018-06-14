using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.IO;
using NPOI.SS;
using NPOI.HSSF.UserModel;


namespace excel导出.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        public ActionResult Excel()
        {
            return View();
        }

        public ActionResult Excel2()
        {
            return View();
        }

        //第一种导出的方法  :采用表格的方式输出
        public ActionResult ExportToExcel()
        {
            var sbHtml = new StringBuilder();

            sbHtml.Append("<table border='1' cellspacing='0' cellpadding='0'>");
            sbHtml.Append("<tr>");

            var lstTitle = new List<string> { "编号","姓名","年龄","创建时间"};

            foreach(var item in lstTitle)
            {
                sbHtml.AppendFormat("<td style='font-size: 14px;text-align:center;background-color: #DCE0E2; font-weight:bold;' height='25'>{0}</td>",item);
            }
            sbHtml.Append("</tr>");

            for(int i=0;i<100;i++)
            {
                sbHtml.Append("<tr>");
                sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", i);
                sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>屌丝{0}号</td>", i);
                sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", new Random().Next(20, 30) + i);
                sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", DateTime.Now);
                sbHtml.Append("</tr>");
            }

            ///第一种 使用FileContentResult
            byte[] fileContents = Encoding.UTF8.GetBytes(sbHtml.ToString());
            return File(fileContents, "application/ms-excel", "fileContents.xls");
        }



        //第二种 导出的方式 NPIO方式
        public ActionResult ExportToExcel2()
        {
            //创建工作簿 
            NPOI.SS.UserModel.IWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            //创建 工作表 
            NPOI.SS.UserModel.ISheet sheet1 = workbook.CreateSheet("SheetTest");    ///通过 工作簿来创建一个 工作表

            //创建Cell 单元格 
            NPOI.SS.UserModel.ICell cell1;    //

            int i = 0;
            int rowLimit = 100;

            DateTime originalTime = DateTime.Now;
            for(i=0; i<rowLimit;i++)
            {
                cell1 = sheet1.CreateRow(i).CreateCell(0);
                cell1.SetCellValue("值"+i.ToString());
            }

            using (MemoryStream ms=new MemoryStream())
            {
                workbook.Write(ms);
                var buffer = ms.GetBuffer();

                ms.Close();
                return File(buffer, "application/ms-excel", "test.xlsx");
            }






        }
    }
}