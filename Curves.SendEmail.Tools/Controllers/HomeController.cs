using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Configuration;
using System.Data;
using Curves.SendEmail.Tools.Models;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Web.Hosting;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace Curves.SendEmail.Tools.Controllers
{
    public class HomeController : Controller
    {
        public static List<ExcelModel> _list { get; set; }
        public static IEnumerable<SelectListItem> TemplateList { get; set; }
        public IEnumerable<string> GetTemplate()
        {
            return new List<string>
            {
                "A產品",
                "B產品"
            };
        }
        public ActionResult Index()
        {
            var templateList = GetTemplate();
            TemplateList = GetSelectListItems(templateList, "");
            return View();
        }

        [HttpPost]
        public JsonResult UploadFile()
        {
            try
            {
                string path = Server.MapPath("~/Uploads/");
                _list = new List<ExcelModel>();
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                HttpPostedFileBase file = Request.Files["inputFile[]"] as HttpPostedFileBase;
                string filePath = path + Path.GetFileName(file.FileName);
                file.SaveAs(filePath);
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //載入Excel檔案
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet sheet = ep.Workbook.Worksheets[1];
                        int startRowNumber = sheet.Dimension.Start.Row;
                        int endRowNumber = sheet.Dimension.End.Row;
                        int startColumn = sheet.Dimension.Start.Column;
                        int endColumn = sheet.Dimension.End.Column;
                        bool isHeader = true;
                        if (isHeader)
                        {
                            startRowNumber += 1;
                        }
                        for (int currentRow = startRowNumber; currentRow <= endRowNumber; currentRow++)
                        {
                            //for (int currentColumn = startColumn; currentColumn <= endColumn; currentColumn++)
                            //{

                            //}
                            ExcelRange range = sheet.Cells[currentRow, startColumn, currentRow, endColumn];//抓出目前的Excel列
                            if (range.Any(c => !string.IsNullOrEmpty(c.Text)) == false)
                            {
                                continue;//略過此列
                            }
                            //讀值
                            _list.Add(new ExcelModel
                            {
                                Email = sheet.Cells[currentRow, 1].Text,
                                Name = sheet.Cells[currentRow, 2].Text,
                                C = sheet.Cells[currentRow, 3].Text,
                                M = int.Parse(sheet.Cells[currentRow, 4].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                Box = int.Parse(sheet.Cells[currentRow, 5].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                B = int.Parse(sheet.Cells[currentRow, 6].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                In = int.Parse(sheet.Cells[currentRow, 7].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                N = int.Parse(sheet.Cells[currentRow, 8].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                All = int.Parse(sheet.Cells[currentRow, 9].Text, NumberStyles.Number, CultureInfo.InvariantCulture),
                                IsSend = false.ToString().ToLower()
                            });
                        }
                    }
                }
            }
            catch (Exception e)
            {
                return Json(new { status = false, message = e.ToString() }, JsonRequestBehavior.AllowGet);
            }
            string result = JsonConvert.SerializeObject(_list);
            return Json(new { status = true, message = result }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public async Task<JsonResult> SendEmail(string template, string sendmail, string lastdate, string email, string name, string c, string m, string box, string b, string i, string n, string all)
        {
            try
            {
                string emailTemplate;
                string url;
                switch (template)
                {
                    case "A產品":
                        emailTemplate = "AProduct";
                        url = "http://www.trible.io";
                        break;
                    default:
                        emailTemplate = "BProduct";
                        url = "http://www.trible.io";
                        break;
                }
                
                var message = await EMailTemplate(emailTemplate);
                message = message.Replace("@ViewBag.Name", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name));
                message = message.Replace("@ViewBag.C", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(c));
                message = message.Replace("@ViewBag.M", String.Format(new CultureInfo("en-US"), "{0:N2}", m));
                message = message.Replace("@ViewBag.Box", String.Format(new CultureInfo("en-US"), "{0:N2}", box));
                message = message.Replace("@ViewBag.B", String.Format(new CultureInfo("en-US"), "{0:N2}", b));
                message = message.Replace("@ViewBag.I", String.Format(new CultureInfo("en-US"), "{0:N2}", i));
                message = message.Replace("@ViewBag.N", String.Format(new CultureInfo("en-US"), "{0:N2}", n));
                message = message.Replace("@ViewBag.All", String.Format(new CultureInfo("en-US"), "{0:N2}", all));
                message = message.Replace("@ViewBag.Link", "<a href=\"" + url + "\" style='color:red'>出貨明細</a>");
                message = message.Replace("@ViewBag.LastDate", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastdate));
                if (email.Contains(','))
                {
                    List<string> emails = email.Split(',').ToList();
                    foreach (string e in emails)
                    {
                        await EmailServices.SendMailAsync(sendmail, e, "【" + template + "】" + m + "月份紅利計算、發票開立作業---" + c, message);
                    }
                }
                else
                {
                    await EmailServices.SendMailAsync(sendmail, email, "【" + template + "】" + m + "月份紅利計算、發票開立作業---" + c, message);
                }
            }
            catch (Exception)
            {
                return Json(new { status = false }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { status = true}, JsonRequestBehavior.AllowGet);
        }

        public async Task<string> EMailTemplate(string template)
        {
            var templateFilePath = HostingEnvironment.MapPath("~/Content/templates/" + template + ".html");
            StreamReader objStreamReaderFile = new StreamReader(templateFilePath);
            var body = await objStreamReaderFile.ReadToEndAsync();
            objStreamReaderFile.Close();
            return body;
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

        private IEnumerable<SelectListItem> GetSelectListItems(IEnumerable<string> elements, string paten)
        {
            var selectList = new List<SelectListItem>();

            foreach (var element in elements)
            {
                if (element == paten)
                {
                    selectList.Add(new SelectListItem
                    {
                        Value = element,
                        Text = element,
                        Selected = true
                    });
                }
                else
                {
                    selectList.Add(new SelectListItem
                    {
                        Value = element,
                        Text = element,
                    });
                }

            }
            return selectList;
        }
    }
}