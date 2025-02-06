using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MeExpert.Models
{
    public class Req
    {
        public int req_id { get; set; }
        public string req_no { get; set; }
        public string title { get; set; }
        public int client_id { get; set; }
        public string client { get; set; }
        public string position { get; set; }
        public int position_id { get; set; }
        public string file_name { get; set; }
        public string file_extension { get; set; }
        public int headcounter { get; set; }
        public int succ_id { get; set; }
        public string req_status { get; set; }
        public string create_date { get; set; }
        public string create_by { get; set; }
        public string s_date { get; set; }
        public string close_date { get; set; }
        public string remark { get; set; }
        public string sel_cand { get; set; }
        public List<PpsReq> propoSals { get; set; }
    }
    public class PpsReq
    {
        public int proposal_id { get; set; }
        public string file_name { get; set; }
        public string create_by { get; set; }
    }
    public class Position
    {
        public int position_id { get; set; }
        public string position_name { get; set; }
    }
    public class Cand
    {
        public int cand_id { get; set; }
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
        public int cv_id { get; set; }
        public bool cv_file { get; set; }
    }
    public class CandOfReq 
    {
        public int req_cand_id { get; set; }
        public int cand_id { get; set; }
        public int req_id { get; set; }
        public int cv_id { get; set; }
        public string cand_name { get; set; }
        public string position { get; set; }
        public string enddate { get; set; }
        public string interviewer { get; set; }
        public int interview_id { get; set; }
        public string remark { get; set; }
        public string interview_remark { get; set; }
        public string interview_date { get; set; }
        public string status { get; set; }
        public string remark_status { get; set; }
        public string cv_path { get; set; }
        public bool send { get; set; }
        public string sent_date { get; set; }
        public string cv_return { get; set; }
        public string create_status { get; set; }
        public List<Interview> interviews { get; set; }
    }
    public class Interview {
        public int interview_id { get; set; }
        public int item_id { get; set; }
        public string remark { get; set; }
        public string interview_date { get; set; }
        public string interviewer_name { get; set; }
        public int cont_id { get; set; }
        public List<ContactPerson> contacts { get; set; }
    }
    public class ReqSucc {
        public int succ_id { get; set; }
        public int req_id { get; set; }
        public int cand_id { get; set; }
        public int cv_id { get; set; }
        public string cand_name { get; set; }
        public string emp_id { get; set; }
        public string enddate { get; set; }
        public string create_date { get; set; }
        public string s_date { get; set; }
        public string s_work_date { get; set; }
        public int req_cand_id { get; set; }
        public string position { get; set; }
    }
    public class FilterReq
    {
        public int year { get; set; }
    }
}