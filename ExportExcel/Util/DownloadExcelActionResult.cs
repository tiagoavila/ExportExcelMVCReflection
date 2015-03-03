using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Reflection;
using ClosedXML.Excel;
using System.IO;

namespace ExportExcel.Util
{
    public class DownloadExcelActionResult<T> : ActionResult where T : class
    {
        public string[] Header { get; set; }
        public List<T> ListOfObjects { get; set; }
        public string FileName { get; set; }

        public DownloadExcelActionResult(string fileName, string[] header, List<T> listOfObjects)
        {
            FileName = fileName;
            Header = header;
            ListOfObjects = listOfObjects;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            var dataTableExcel = new DataTable("Planilha");

            foreach (var cell in Header)
            {
                dataTableExcel.Columns.Add(cell);
            }

            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            foreach (T item in ListOfObjects)
            {
                dataTableExcel.Rows.Add();
                for (int i = 0; i < propertyInfos.Count(); i++)
                {
                    var propertyValue = propertyInfos[i].GetValue(item, null);
                    dataTableExcel.Rows[dataTableExcel.Rows.Count - 1][i] = propertyValue;
                }
            }

            HttpContext curContext = HttpContext.Current;
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dataTableExcel);
                curContext.Response.Clear();
                curContext.Response.Buffer = true;
                curContext.Response.Charset = "";
                curContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                curContext.Response.AddHeader("content-disposition", "attachment;filename=" + FileName + ".xlsx");
                using (var myMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(myMemoryStream);
                    myMemoryStream.WriteTo(curContext.Response.OutputStream);
                    curContext.Response.Flush();
                    curContext.Response.End();
                }
            }
        }
    }
}