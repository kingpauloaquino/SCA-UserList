using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AuditSheet
{
    public class SlaveObject
    {
        
    }

    public class SalesSlave
    {
        public int Message { get; set; }
        public List<Buyers> buyer_list { get; set; }
        public List<Reports> Description { get; set; }
    }

    public class Buyers
    {
        public string user_name { get; set; }
    }

    public class Reports
    {
        public string box_status { get; set; }
        public string lot_uid { get; set; }
        public string member_id { get; set; }
        public string lot_number { get; set; }
        public string date_posted { get; set; }
        public string date_awarded { get; set; }
        public string number_of_box { get; set; }
        public string number_of_units { get; set; }
        public string full_equi_units { get; set; }
        public string seller_username { get; set; }
        public string index_type { get; set; }
        public string zipcode { get; set; }
        public string number_of_views { get; set; }
        public string number_of_bids { get; set; }
        public string winning_buyer { get; set; }
        public string winning_bid { get; set; }
        public string avg_per_feu { get; set; }
        public string consignee_zip_code { get; set; }
        public string sca_buyer_bid_appraisal { get; set; }
        public string sca_avg_per_feu { get; set; }
        public List<Bids> buyer_bid { get; set; }
    }

    public class Bids
    {
        public string user_name { get; set; }
        public string buyer_bid { get; set; }
    }
}