using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace MeExpert.Models
{
    public class ProposalDoc
    {
        public string position { get; set; }
        public string start_date { get; set; }
        public string end_date { get; set; }
        public string scop { get; set; }
        public string company { get; set; }
        public string contact_date { get; set; }
        public int working_date { get; set; }
        public string proposal_date { get; set; }
        public string scope_service { get; set; }
        public decimal act_salary { get; set; }
        public int period { get; set; }
        public string rev_no { get; set; }
        public string file_name { get; set; }
        public string candidate_name { get; set; }
        public double mr_1_2 { get; set; }
        public double mr_1_4 { get; set; }
        public double mr_1_6 { get; set; }
        public double mr_1_7 { get; set; }
        public double tt_1_1 { get; set; }
        public double tt_1_2 { get; set; }
        public double tt_1_4 { get; set; }
        public double tt_1_6 { get; set; }
        public double tt_1_7 { get; set; }
        public double ot_1_1 { get; set; }
        public double subtotal1 { get; set; }
        public double tt_2_1 { get; set; }
        public double tt_2_2 { get; set; }
        public double tt_2_3 { get; set; }
        public double tt_2_10 { get; set; }
        public double subtotal2 { get; set; }
        public double subtotal2_dh { get; set; }
        public double grandtotal { get; set; }
        public double grandtotal_dh { get; set; }
        public double mr_subtotal { get; set; }
    }
    public class FilePPS
    {
        public MemoryStream stream { get; set; }
        public string position { get; set; }
        public string rev_no { get; set; }
        public int proposal_id { get; set; }
    }
}