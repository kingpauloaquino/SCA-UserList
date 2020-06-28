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
    public partial class saleslave : System.Web.UI.Page
    {
        private static Excel.Application xlApp;
        private static Excel.Workbook xlWorkbook;
        private static Excel._Worksheet xlWorksheet;
        private static Excel.Range xlRange;

        private static string ApiUrl { get { return ConfigurationManager.AppSettings["SalesSlaveApiUrl"]; } }
        private static string ExcelSource { get { return ConfigurationManager.AppSettings["SalesSlaveExcelSource"]; } }
        private static string ExcelSavePath { get { return ConfigurationManager.AppSettings["SalesSlaveExcelSavePath"]; } }
        private static string ExcelLinkDownload { get { return ConfigurationManager.AppSettings["SalesSlaveExcelLinkDownload"]; } }
        private static string LotNumber { get; set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["from"] == null)
            {
                JObject o = new JObject();
                o["status"] = 401;
                o["message"] = "Please enter from date.";
                Response.Write(o.ToString());
                return;
            }

            if (Request.QueryString["to"] == null)
            {
                JObject o = new JObject();
                o["status"] = 402;
                o["message"] = "Please enter to date.";
                Response.Write(o.ToString());
                return;
            }

            Write(Request.QueryString["from"], Request.QueryString["to"]);
        }

        public void Write(string from, string to)
        {
            JObject o;

            try
            {
                SalesSlave a = SalesSlaveReports(from, to);
                
                if (a.Message != 200)
                {
                    o = new JObject();
                    o["status"] = 402;
                    o["message"] = "Oops, something went wrong.";
                    Response.Write(o.ToString());
                    return;
                }

                if (a.buyer_list.Count == 0)
                {
                    o = new JObject();
                    o["status"] = 403;
                    o["message"] = "Lot number no data record.";
                    Response.Write(o.ToString());
                    return;
                }

                if (a.Description.Count == 0)
                {
                    o = new JObject();
                    o["status"] = 403;
                    o["message"] = "Lot number no data record.";
                    Response.Write(o.ToString());
                    return;
                }

                if (!File.Exists(ExcelSource))
                {
                    o = new JObject();
                    o["status"] = 401;
                    o["message"] = "Excel source file did not found.";

                    Response.Write(o.ToString());
                    return;
                }

                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(ExcelSource);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                // encode all buyers in column 1
                SlaveHelper.encode_buyer_list(a.buyer_list, xlApp);

                // encode buyer bids
                SlaveHelper.encode_sales_slave(a, xlApp, xlWorksheet);
                
               string excel_name = Flush();
                
                o = new JObject();
                o["status"] = 200;
                o["message"] = "Success.";
                o["link"] = ExcelLinkDownload + excel_name.Replace(" ", "%20");
                Response.Write(o.ToString());
                return;
            }
            catch (Exception ex)
            {
                o = new JObject();
                o["status"] = 500;
                o["message"] = "Error Message: " + ex.Message;
                Response.Write(o.ToString());
                return;
            }
        }

        public string Flush()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            string datetime = DateTime.Now.ToString("yyMMddhhmmss");

            string excel_name = "Lot_Sales_Report-"+ datetime + ".xlsx"; // Guid.NewGuid() + ".xlsx"; // 

            string path_des = ExcelSavePath + excel_name;
            
            if (File.Exists(path_des))
            {
                File.Delete(path_des);
            }
            xlWorksheet.SaveAs(path_des);

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return excel_name;
        }

        public SalesSlave SalesSlaveReports(string from, string to)
        {
            System.Threading.Thread.Sleep(200);
            SalesSlave slaves = new SalesSlave();

            string url_path = ApiUrl + "/v3/lot_sales_report/795edd365fd0e371ceaaf1ddd559a85d/" + from + "/" + to;

            try
            {
                WebClient client = new WebClient();
                string value = client.DownloadString(url_path);
                slaves = JsonConvert.DeserializeObject<SalesSlave>(value);
            }
            catch (Exception ex)
            {
                slaves.Message = 500;
            }
            return slaves;
        }
    }
}