using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Runtime.InteropServices;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Configuration;

namespace AuditSheet
{
    public partial class import_excel : System.Web.UI.Page
    {
        private static Excel.Application xlApp;
        private static Excel.Workbook xlWorkbook;
        private static Excel._Worksheet xlWorksheet;
        private static Excel.Range xlRange;

        private List<UserList> UserListArray;

        protected void Page_Load(object sender, EventArgs e)
        {
            UserListArray = new List<UserList>();

            if(Request.QueryString["excelfile"] == null)
            {
                JObject o = new JObject();
                o["status"] = 200;
                o["message"] = "Oops, the filename could not found.";
                Response.Write(o.ToString());
                return;
            }

            string path = @"C:\temp\" + Request.QueryString["excelfile"];

            HandleExcel(path);

        }


        private void HandleExcel(string ExcelSource)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(ExcelSource);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;

            int rows = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            for(int x = 0; x < rows; x++)
            {
                if(x > 1 )
                {
                    if (xlRange.Cells[x, 1] != null && xlRange.Cells[x, 1].Value2 != null)
                    {
                        UserList ul = new UserList();

                        ul.UserId = Convert.ToInt32(xlRange.Cells[x, 1].Value2);
                        ul.WorkerId = Convert.ToInt32(xlRange.Cells[x, 2].Value2);
                        ul.Role = xlRange.Cells[x, 3].Value2.ToString();
                        ul.Active = Convert.ToInt32(xlRange.Cells[x, 4].Value2);
                        ul.Username = xlRange.Cells[x, 5].Value2.ToString();
                        ul.PasswordHint = xlRange.Cells[x, 6].Value2.ToString();
                        ul.BusinessName = xlRange.Cells[x, 7].Value2.ToString();
                        ul.Firstname = xlRange.Cells[x, 8].Value2.ToString();
                        ul.Middlename = xlRange.Cells[x, 9].Value2.ToString();
                        ul.Lastname = xlRange.Cells[x, 10].Value2.ToString();
                        ul.Position = xlRange.Cells[x, 11].Value2.ToString();
                        ul.EmailAddress = xlRange.Cells[x, 12].Value2.ToString();
                        ul.MobileNumber = xlRange.Cells[x, 13].Value2.ToString();
                        ul.Address = xlRange.Cells[x, 14].Value2.ToString();
                        ul.City = xlRange.Cells[x, 15].Value2.ToString();
                        ul.State = xlRange.Cells[x, 16].Value2.ToString();
                        ul.Country = xlRange.Cells[x, 17].Value2.ToString();
                        ul.Zipcode = xlRange.Cells[x, 18].Value2.ToString();
                        ul.Type = xlRange.Cells[x, 19].Value2.ToString();
                        ul.Created_At = xlRange.Cells[x, 20].Value2.ToString();
                        ul.Updated_At = xlRange.Cells[x, 21].Value2.ToString();

                        UserListArray.Add(ul);
                    }
                }
            }

            ResponseUser res = new ResponseUser()
            {
                Status = 200,
                Message = "",
                Rows = rows,
                Users = UserListArray
            };

            string json = JsonConvert.SerializeObject(res);

            Response.Write(json);
        }
    }
}