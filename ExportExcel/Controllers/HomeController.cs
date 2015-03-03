using ExportExcel.Models;
using ExportExcel.Util;
using System.Collections.Generic;
using System.Web.Mvc;

namespace ExportExcel.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportExcel()
        {
            var listUsers = new List<User>
                                {
                                    new User
                                        {
                                            Name = "Tiago",
                                            LastName = "Avila",
                                            Cpf = "000.000.000-00"
                                        },
                                    new User
                                        {
                                            Name = "Avila",
                                            LastName = "Tiago",
                                            Cpf = "123.456.789-01"
                                        }
                                };
            return new DownloadExcelActionResult<User>("MyExcel", new[]{"Nome", "Sobre Nome", "Cpf"}, listUsers);
        }

    }
}
