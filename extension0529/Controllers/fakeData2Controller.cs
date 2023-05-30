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
    public class fakeData2Controller : Controller
    {
        // GET: fakeData2
        public ActionResult ExportToExcel()
        {


            List<fakeData> fakeDataList = new List<fakeData>
            {
    new fakeData { Name = "John Smith", Age = 25, Email = "john.smith@example.com" },
    new fakeData { Name = "Jane Doe", Age = 30, Email = "jane.doe@example.com" },
    new fakeData { Name = "Robert Johnson", Age = 35, Email = "robert.johnson@example.com" }
            };

            var viewModel = new BigView1
            {
                Report1 = fakeDataList,
                Report2 = fakeDataList,
            };

            //return Json(fakeDataList, JsonRequestBehavior.AllowGet);


            var fileName = "0529_" + Guid.NewGuid().ToString() + ".xlsx";
            var guidFileName = fakeDataList.ExportExcel(fileName); //2個

            return File(guidFileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "repoxrt.xlsx");
           



        }
    }
}