using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Excel = Microsoft.Office.Interop.Excel;

namespace AuditSheet
{
    public class SlaveHelper
    {
        public static int column_default { get; set;}

        public static void encode_buyer_list(List<Buyers> buyers, Excel.Application xlApp, int colCount = 17)
        {
            foreach (Buyers b in buyers)
            {
                xlApp.Cells[1, colCount] = b.user_name.ToUpper();

                colCount++;
            }
        }

        public static void encode_buyer_bid(Excel.Application xlApp, Excel._Worksheet xlWorksheet, List<Buyers> buyers, int row, string buyer_username, string buyer_bid_amount)
        {
            int buyer_id = SlaveHelper.get_buyer_id(buyers, buyer_username);

            int column_id = column_default + buyer_id;

            SlaveHelper.put_buyer_bid(xlApp, xlWorksheet, row, column_id, buyer_bid_amount);
        }

        public static void put_buyer_bid(Excel.Application xlApp, Excel._Worksheet xlWorksheet, int row, int col, string value)
        {
            xlApp.Cells[row, col] = value;

            ExcelHelper.setCellWidth(xlWorksheet, col);

            ExcelHelper.setCurrencyCell(xlWorksheet, row, col);
        }

        public static int get_buyer_id(List<Buyers> buyers, string buyer_name)
        {
            int column = 0;
            for (int i = 0; i < buyers.Count; i++)
            {
                column++;

                if (buyers[i].user_name.Contains(buyer_name))
                {
                    return column;
                }

            }

            return column;
        }

        // 

        public static void encode_sales_slave(SalesSlave slaves, Excel.Application xlApp, Excel._Worksheet xlWorksheet)
        {
            int row = 2;

            column_default = 16;

            foreach (Reports r in slaves.Description)
            {
                xlApp.Cells[row, 1] = r.lot_number;
                xlApp.Cells[row, 2] = r.date_posted;
                xlApp.Cells[row, 3] = r.date_awarded != null ? r.date_awarded : "NYA" ;
                xlApp.Cells[row, 4] = r.number_of_box;
                xlApp.Cells[row, 5] = r.number_of_units;
                xlApp.Cells[row, 6] = r.full_equi_units;
                xlApp.Cells[row, 7] = r.seller_username;
                xlApp.Cells[row, 8] = r.index_type;
                xlApp.Cells[row, 9] = r.zipcode;
                xlApp.Cells[row, 10] = r.number_of_views;
                xlApp.Cells[row, 11] = r.number_of_bids;
                xlApp.Cells[row, 12] = r.winning_buyer != null ? r.winning_buyer : "NYA";
                xlApp.Cells[row, 13] = r.winning_bid != null ? r.winning_bid : "NYA";
                xlApp.Cells[row, 14] = r.avg_per_feu;
                xlApp.Cells[row, 15] = r.consignee_zip_code;
                xlApp.Cells[row, 16] = r.sca_buyer_bid_appraisal;
                xlApp.Cells[row, 17] = r.sca_avg_per_feu;

                int t_colds = (18 + slaves.buyer_list.Count);

                for (int i = 18; i < t_colds; i++)
                {
                    xlApp.Cells[row, i] = "N/A";
                }

                foreach(Bids bid in r.buyer_bid)
                {
                    encode_buyer_bid(xlApp, xlWorksheet, slaves.buyer_list, row, bid.user_name, bid.buyer_bid);
                }

                row++;
            }
        }
    }
}