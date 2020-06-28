using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AuditSheet
{
    public class AuditObject
    { }

    // Audit version 1

    public class ResultV1
    {
        public string box_code { get; set; }
        public string unique_unit_number { get; set; }
        public string fullness_grade { get; set; }
        public string comments { get; set; }
    }

    public class Audit
    {
        public int Message { get; set; }
        public List<ResultV1> Result { get; set; }
    }

    // Audit version 2

    public class ResultV2
    {
        public string box_code { get; set; }
        public List<BoxResults> box_result { get; set; }
    }

    public class BoxResults
    {
        public string unique_unit_number { get; set; }
        public string fullness_grade { get; set; }
        public string comments { get; set; }
    }

    public class AuditV2
    {
        public int Message { get; set; }
        public List<ResultV2> Result { get; set; }
    }

    public class Pages
    {
        public int PageCount { get; set; }
        public int RowCount { get; set; }
    }
}