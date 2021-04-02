using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LeituraPlanilhas.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedfile)
        {
            string path = Server.MapPath("~/Uploads/~");
            string filePath = string.Empty;
            string extension = string.Empty;
            DataTable dtSheet = new DataTable();

            if (postedfile != null)
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);


                }
                filePath = path + Path.GetFileName(postedfile.FileName);
                extension = Path.GetExtension(postedfile.FileName);
                postedfile.SaveAs(filePath);
            }
            string connectionString = string.Empty;
            //Codigo que cria a conexão com extensões e faz o upload da planilha
            switch (extension)
            {
                case ".xls": //para planilhas antigas

                    connectionString = ConfigurationManager.ConnectionStrings["Excel03ConsString"].ConnectionString;
                    break;
                case ".xlsx": // do 2007 em diante
                    connectionString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }

            //Aqui é o OLEDB COnnection
            connectionString = string.Format(connectionString, filePath);
            using (OleDbConnection connExcel = new OleDbConnection(connectionString))
            {
                using(OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        cmdExcel.Connection = connExcel;

                        //Pegar os dados da planilha
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string nomeAba = dtExcelSchema.Rows[3]["EXEC SCRIPT"].ToString();
                        connExcel.Close();
                        //Lendo a planilha

                        connExcel.Open();
                        cmdExcel.CommandText = "Select * from [" + nomeAba + "] ";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dtSheet);
                        connExcel.Close();

                    }
                }
            }
            return View(dtSheet);
        }

    }
}