using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
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
        private ResponseUser res;

        protected void Page_Load(object sender, EventArgs e)
        {
            UserListArray = new List<UserList>();

            if(Request.QueryString["excelfile"] == null)
            {
                JObject o = new JObject();
                o["status"] = 404;
                o["message"] = "Oops, the filename could not found.";
                Response.Write(o.ToString());
                return;
            }

            string path = @"C:\temp\" + Request.QueryString["excelfile"];

            HandleExcel(path);

        }

        private void HandleExcel(string ExcelSource)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(ExcelSource);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                Thread.Sleep(500);

                int rows = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row + 1;

                Thread.Sleep(500);

                for (int x = 0; x < rows; x++)
                {
                    Thread.Sleep(200);

                    if (x > 1)
                    {
                        if (xlRange.Cells[x, 1] != null && xlRange.Cells[x, 1].Value2 != null)
                        {
                            UserList ul = new UserList();

                            ul.UserId = check_cell(x, 1);
                            ul.WorkerId = check_cell(x, 2);
                            ul.Role = check_cell(x, 3);
                            ul.Active = check_cell(x, 4);
                            ul.Username = check_cell(x, 5);
                            ul.PasswordHint = check_cell(x, 6);
                            ul.BusinessName = check_cell(x, 7);
                            ul.Firstname = check_cell(x, 8);
                            ul.Middlename = check_cell(x, 9);
                            ul.Lastname = check_cell(x, 10);
                            ul.Position = check_cell(x, 11);
                            ul.EmailAddress = check_cell(x, 12);
                            ul.MobileNumber = check_cell(x, 13);
                            ul.Address = check_cell(x, 14);
                            ul.City = check_cell(x, 15);
                            ul.State = check_cell(x, 16);
                            ul.Country = check_cell(x, 17);
                            ul.Zipcode = check_cell(x, 18);
                            ul.Type = check_cell(x, 19);
                            ul.Created_At = check_cell(x, 20);
                            ul.Updated_At = check_cell(x, 21);

                            UserListArray.Add(ul);
                        }
                    }
                }

                res = new ResponseUser()
                {
                    Status = 200,
                    Message = "Success",
                    Rows = rows - 2,
                    Users = UserListArray
                };
            }
            catch(Exception ex)
            {
                res = new ResponseUser()
                {
                    Status = 500,
                    Message = ex.Message
                };
            }

            Thread.Sleep(500);

            string json = JsonConvert.SerializeObject(res);
            Response.Write(json);
        }

        private string check_cell(int row, int col)
        {
            try
            {
                if (xlRange.Cells[row, col] != null)
                {
                    return xlRange.Cells[row, col].Value2.ToString();
                }
            }
            catch(Exception ex)
            {
                string mx = ex.Message;
            }
            
            return "";
        }
    }
}