using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MeExpert.Models
{
    public class Candidate
    {
        public int cand_id { get; set; }
        public int item_id { get; set; }
        public int no { get; set; }
        public int req_cand_id { get; set; }
        public string full_name { get; set; }
        public string first_name { get; set; }
        public string last_name { get; set; }
        public string mobile_no_1 { get; set; }
        public string mail { get; set; }
        public string gender { get; set; }
        public string apply_position { get; set; }
        public string cv_path { get; set; }
        public string apply_date { get; set; }
        public string exp_salary { get; set; }
        public string status { get; set; }
        public string emp_id { get; set; }
        public string req_no { get; set; }
        public string proposal { get; set; }
        public string age { get; set; }
        public string last_apply_position { get; set; }
        public string remarks { get; set; }
        public string mobile { get; set; }
        public string create_by { get; set; }
        public string create_date { get; set; }
        public string id_card { get; set; }
        public string birth_day { get; set; }
        public string mobile_no_2 { get; set; }
        public string toeic_score { get; set; }
        public string start_working { get; set; }
        public string employ_status { get; set; }
        public string cand_remarks { get; set; }
        public string create_time { get; set; }
        public List<Education> educations { get; set; }
    }
    public class Proposal
    {
        public int proposal_id { get; set; }
        public int succ_id { get; set; }
        public string file_name { get; set; }
        public string file_extenstion { get; set; }
        public string position { get; set; }
        public string create_by { get; set; }
        public string create_date { get; set; }
        public int rev_no { get; set; }
        public string start_date { get; set; }
        public string end_date { get; set; }
        public int working_date { get; set; }
        public string proposal_date { get; set; }
        public string scope_service { get; set; }
        public decimal act_salary { get; set; }
        public decimal pps_comp { get; set; }
        public decimal pps_license { get; set; }
        public int period { get; set; }
    }
    public class Education
    {
        public int item_id { get; set; }
        public int cand_id { get; set; }
        public string degree { get; set; }
        public string inst_name { get; set; }
        public string faculty { get; set; }
        public string major { get; set; }
        public string grade { get; set; }
        public string from_date { get; set; }
        public string to_date { get; set; }
    }
    public class ApplyRecord
    {
        public int cv_id { get; set; }
        public int cand_id { get; set; }
        public int item_id { get; set; }
        public string cv_no { get; set; }
        public string apply_position { get; set; }
        public string exp_salary { get; set; }
        public string industry { get; set; }
        public string keyword { get; set; }
        public string remark { get; set; }
        public string apply_date { get; set; }
        public string cv_file { get; set; }
    }
}