using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MeExpert.Models
{
    public class Report
    {
        public string no { get; set; }
        public string date_job { get; set; }
        public string position { get; set; }
        public string area { get; set; }
        public string client { get; set; }
        public string candidate { get; set; }
        public string remark { get; set; }
        public string cv_sent { get; set; }
        public int sent_days { get; set; }
        public string cv_return { get; set; }
        public string interview { get; set; }
        public string salary { get; set; }
        public string health { get; set; }
        public string submit { get; set; }
        public string starting_date { get; set; }
    }
}