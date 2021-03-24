using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;using System.Data.SqlClient;
using System.Web.Script.Serialization;
using System.Web.Configuration;

namespace WebApplication4.Controllers
{
    public class HomeController : Controller
    {

        public ActionResult Index()
        {
            if (Session["user"] == null)
            {
                Response.Redirect("/Home/Login");
                
            }
            return View();
        }

        public ActionResult Login()
        {
            return View();
        }
        [HttpPost]
        public JsonResult Create(string isim, string soyisim, string adres, string email)
        {
        
            SqlConnection abc = new SqlConnection(@"Data Source=.;Initial Catalog=tutorial;Integrated Security=True;");
            using (SqlCommand cmd = new SqlCommand(String.Format("INSERT INTO personalinformation(Name, LastName, Adress, Email) VALUES (@Name, @LastName, @Adress, @Email)"), abc))
            {
                cmd.Parameters.Add("@Name", SqlDbType.NVarChar, 50).Value = isim;
                cmd.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = soyisim;
                cmd.Parameters.Add("@Adress", SqlDbType.NVarChar, 50).Value = adres;
                cmd.Parameters.Add("@Email", SqlDbType.NVarChar, 50).Value = email;
                abc.Open();
                cmd.ExecuteNonQuery();
                abc.Close();
         
            }
       
            return Json(new { success = true });
        }
        public JsonResult LoginUser(string userName, string password)
        {

            SqlConnection cnn = new SqlConnection(@"Data Source=.;Initial Catalog=tutorial;Integrated Security=True;");
            cnn.Open();
            SqlCommand cmd = new SqlCommand(String.Format("select COUNT(*) from Users where userName='{0}' and password='{1}'",userName, password),cnn);
            var result = (Int32)cmd.ExecuteScalar();
            cnn.Close();
            if (result>0 )
            {
                Session["user"] = 1;
                return Json(new { success = true });

            }
            else
            {

                return Json(new { Success = false, Message = "Yanlış", Code = 100 }, JsonRequestBehavior.AllowGet);

            }

        }

        public void Logout()
        {
            Session.Remove("user");
            Response.Redirect("/Home");
        }
        public class Kisi
        {
                public string ad { get; set; }
                public string soyad { get; set; }
                public string adres { get; set; }
                public string email { get; set; }

         }


    public ActionResult Excel(string isim, string soyisim,  string adres, string email)
        {
            var kisi = new Kisi();
            kisi.ad = isim;
            kisi.soyad = soyisim;
            kisi.adres = adres;
            kisi.email = email;
            var kisiler = new List<Kisi>();

            if (Session["kisiler"] == null) 
            {
                
                kisiler.Add(kisi);
                Session["kisiler"] = kisiler;
                
            }
            else
            {
                kisiler = (List<Kisi>)(Session["kisiler"]);
                kisiler.Add(kisi);
                Session["kisiler"] = kisiler;
            }

           

            using (MemoryStream mem = new MemoryStream())
            {

                var spreadsheetDocument = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook);

                var workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

            
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                SheetData sheetData1 = new SheetData();


                var baslikSatiri = new Row();
                baslikSatiri.Append(CreateCell("İsim"));
                baslikSatiri.Append(CreateCell("Soyisim"));
                baslikSatiri.Append(CreateCell("Adres"));
                baslikSatiri.Append(CreateCell("E-mail"));

                sheetData1.Append(baslikSatiri);

                foreach (var item in kisiler)
                {
                    var tRow = new Row();
                    tRow.Append(CreateCell(item.ad));
                    tRow.Append(CreateCell(item.soyad));
                    tRow.Append(CreateCell(item.adres));
                    tRow.Append(CreateCell(item.email));

                    sheetData1.Append(tRow);
                }
                

                worksheetPart.Worksheet = new Worksheet(sheetData1);

         
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

    
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "sayfa1"
                };

                
                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                spreadsheetDocument.Close();

            

                string handle = Guid.NewGuid().ToString();

                mem.Position = 0;
                TempData[handle] = mem.ToArray();
                      
                return new JsonResult()
                {
                    Data = new { FileGuid = handle, FileName = "dosya.xlsx" }
                };
               
            }

           
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(text);
            return cell;
        }

        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-excel", fileName);
            }
            else
            {

                return new EmptyResult();
            }
        }

    }
}