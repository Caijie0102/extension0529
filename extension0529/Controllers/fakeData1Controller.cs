using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.WebPages;
using System.Threading;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using extension0529.Models;
using extension0529.Extension;



namespace extension0529.Controllers
{
    public class fakeData1Controller : Controller
    {
        // GET: fakeData1
        public ActionResult ExportToExcel()
        {
            List<fakeData> fakeDataList = new List<fakeData>
            {
    new fakeData { Name = "John Smith", Age = 25, Email = "john.smith@example.com" },
    new fakeData { Name = "Jane Doe", Age = 30, Email = "jane.doe@example.com" },
    new fakeData { Name = "Robert Johnson", Age = 35, Email = "robert.johnson@example.com" }
            };
            List<fakeData> fakeDataList2 = new List<fakeData>
            {
    new fakeData { Name = "喬瑟夫", Age = 25, Email = "john.smith@example.com" },
    new fakeData { Name = "珍妮", Age = 30, Email = "jane.doe@example.com" },
    new fakeData { Name = "強森", Age = 35, Email = "robert.johnson@example.com" },
    new fakeData { Name = "凱利", Age = 35, Email = "robert.johnson@example.com" }
            };
            var viewModel = new BigView1
            {
                Report1 = fakeDataList, //Report1為資料表名稱
                Report2= fakeDataList2, //Report2為資料表名稱
            };

            var fileName = "0529_" + Guid.NewGuid().ToString() + ".xlsx";
            var guidFileName = fakeDataList.ExportExcel(fileName);//2個 損毀
            var guidFileName1 = fakeDataList.ExportExcel<fakeData>(fileName, "ss"); //3個，且ss為我資料表的名稱
            var guidFileName2 = viewModel.ExportExcel<BigView1>(fileName);//2個

            return File(guidFileName2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            //return File(guidFileName, "application/pdf", "repoxrt.pdf");//excel偉裝成pdf
            //return Json(viewModel, JsonRequestBehavior.AllowGet);
            


        }


        //return Json(viewModel, JsonRequestBehavior.AllowGet);
    }
}