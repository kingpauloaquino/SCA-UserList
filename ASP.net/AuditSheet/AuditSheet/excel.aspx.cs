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
    public partial class excel : System.Web.UI.Page
    {
        private static Excel.Application xlApp;
        private static Excel.Workbook xlWorkbook;
        private static Excel._Worksheet xlWorksheet;
        private static Excel.Range xlRange;

        private static string ApiUrl { get { return ConfigurationManager.AppSettings["ApiUrl"]; } }
        private static string ExcelSource { get { return ConfigurationManager.AppSettings["ExcelSource"]; } }
        private static string ExcelSavePath { get { return ConfigurationManager.AppSettings["ExcelSavePath"]; } }
        private static string ExcelLinkDownload { get { return ConfigurationManager.AppSettings["ExcelLinkDownload"]; } }
        private static string ZipSavePath { get { return ConfigurationManager.AppSettings["ZipSavePath"]; } }
        private static string ZipLinkDownload { get { return ConfigurationManager.AppSettings["ZipLinkDownload"]; } }
        private static string LotNumber { get; set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["lot"] ==null) 
            {
                JObject o = new JObject();
                o["status"] = 402;
                o["message"] = "Please enter lot number.";
                Response.Write(o.ToString());
                return;
            }

            LotNumber = Request.QueryString["lot"];

            Write();
        }

        public void Write()
        {
            JObject o;

            try
            {
                Audit a = Auditlogs();
                int row = 2;
                
                if (a.Message != 200)
                {
                    o = new JObject();
                    o["status"] = 402;
                    o["message"] = "Lot number did not found.";
                    Response.Write(o.ToString());
                    return;
                }

                if (a.Result.Count == 0)
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

                string box_code = null;

                int row_count = 0;

                int cgs_count = 1;

                foreach (ResultV1 r in a.Result)
                {
                    System.Threading.Thread.Sleep(100);

                    if (box_code != null)
                    {
                        if (box_code != r.box_code)
                        {
                            cgs_count = 1;

                            Excel.Range lineTarget = (Excel.Range)xlWorksheet.Rows[row_count];
                            xlWorksheet.HPageBreaks.Add(lineTarget);

                            System.Threading.Thread.Sleep(500);
                        }
                    }
                    
                    box_code = r.box_code;

                    string comments = null;

                    if (r.comments !=null )
                    {
                        comments = r.comments.Replace("+", " ").Replace("%0A", " ").Replace("<br>", " ").Replace("<br >", " ").Replace("<br />", " ");
                    }
                    
                    xlApp.Cells[row, 1] = box_code;
                    xlApp.Cells[row, 2] = r.unique_unit_number;
                    xlApp.Cells[row, 3] = cgs_count.ToString();
                    xlApp.Cells[row, 4] = r.fullness_grade;
                    //xlApp.Cells[row, 5] = "5";
                    xlApp.Cells[row, 6] = comments;
                    //xlApp.Cells[row, 7] = "7";
                    //xlApp.Cells[row, 8] = "8";
                    //xlApp.Cells[row, 9] = "9";

                    for (int c = 1; c <= 9; c++)
                    {
                        setBorderCell(row, c);
                    }

                    Excel.Range line = (Excel.Range)xlWorksheet.Rows[row + 1];
                    line.Insert();

                    row++;
                    cgs_count++;
                    row_count = row;
                }

                string filename = Flush();

                o = new JObject();
                o["status"] = 200;
                o["message"] = "Success.";
                o["link"] = ExcelLinkDownload + LotNumber + "/" + filename;
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

        public void WriteV2()
        {
            AuditV2 a = AuditlogsV2();
           
            JObject o;

            if (a.Message != 200)
            {
                o = new JObject();
                o["status"] = 402;
                o["message"] = "Lot number did not found.";
                Response.Write(o.ToString());
                return;
            }

            if (a.Result.Count == 0)
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
            
            foreach (ResultV2 r in a.Result)
            {
                string box = r.box_code;
                List<BoxResults> results = r.box_result;

                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(ExcelSource);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int row = 2;

                foreach (BoxResults b in results)
                {
                    xlApp.Cells[row, 1] = box;
                    xlApp.Cells[row, 2] = b.unique_unit_number;
                    //xlApp.Cells[row, 3] = "3";
                    xlApp.Cells[row, 4] = b.fullness_grade;
                    //xlApp.Cells[row, 5] = "5";
                    xlApp.Cells[row, 6] = b.comments;
                    //xlApp.Cells[row, 7] = "7";
                    //xlApp.Cells[row, 8] = "8";
                    //xlApp.Cells[row, 9] = "9";
                    
                    for (int c = 1; c <= 9; c++)
                    {
                        setBorderCell(row, c);
                    }

                    Excel.Range line = (Excel.Range)xlWorksheet.Rows[row + 1];
                    line.Insert();
                    row++;
                }

                 Flush(box);
            }

            //string filename = Flush();

            bool result = ZipFile.Compress(ExcelSavePath + LotNumber, ZipSavePath, LotNumber);

            o = new JObject();
            o["status"] = 200;
            o["message"] = "Success.";
            o["link"] = ZipLinkDownload + LotNumber + ".zip";
            Response.Write(o.ToString());
            return;
        }

        public string Flush(string box = "")
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if(!Directory.Exists(ExcelSavePath + LotNumber))
            {
                Directory.CreateDirectory(ExcelSavePath + LotNumber);
            }
            
            string path_des = ExcelSavePath + LotNumber + @"\" + LotNumber + " Audit Sheet.xlsx";

            //string path_des = ExcelSavePath + LotNumber + @"\" + LotNumber + "-" + box + " Audit Sheet.xlsx";
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

            return LotNumber + "%20Audit%20Sheet.xlsx";
        }
        
        public Audit Auditlogs()
        {
            System.Threading.Thread.Sleep(200);
            Audit audit = new Audit();

            string url_path = ApiUrl + "/api/v1/audit_lots_for_lot_number/795edd365fd0e371ceaaf1ddd559a85d/" + LotNumber;
            
            try
            {
                WebClient client = new WebClient();
                string value = client.DownloadString(url_path);
                audit = JsonConvert.DeserializeObject<Audit>(value);
            }
            catch (Exception ex)
            {
                audit.Message = 500;
            }
            return audit;
        }

        public AuditV2 AuditlogsV2()
        {
            System.Threading.Thread.Sleep(200);
            AuditV2 audit = new AuditV2();

            string url_path = ApiUrl + "/api/v2/audit_lots_for_lot_number/795edd365fd0e371ceaaf1ddd559a85d/" + LotNumber;

            try
            {
                WebClient client = new WebClient();
                string value = client.DownloadString(url_path);
                audit = JsonConvert.DeserializeObject<AuditV2>(value);
            }
            catch (Exception ex)
            {
                audit.Message = 500;
            }
            return audit;
        }
        
        public void setBorderCell(int row, int cell)
        {
            Excel.Range cell1 = xlWorksheet.Cells[row, cell];
            cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell1.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }
        
    }
}