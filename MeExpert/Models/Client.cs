using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MeExpert.Models
{
    public class Client
    {
        public int no { get; set; }
        public int client_id { get; set; }
        public string client_name { get; set; }
        public string client_address { get; set; }
        public string client_tel { get; set; }
        public bool isitem { get; set; }
        public string create_by { get; set; }
        public string create_date { get; set; }
        public string cont_person { get; set; }
        public string cont_mobile { get; set; }
        public string remark { get; set; }
        public List<ContactPerson> arr_data { get; set; }
        public List<Remark> arr_rmk { get; set; }
    }
    public class ContactPerson
    {
        public int cont_id { get; set; }
        public int client_id { get; set; }
        public string cont_person_name { get; set; }
        public string cont_person_mobile { get; set; }
        public string cont_pos { get; set; }
    }
    public class Remark
    {
        public int item_id { get; set; }
        public int client_id { get; set; }
        public string remark { get; set; }
        public string create_by { get; set; }
    }
    public class SrvFile
    {
        public int srv_id { get; set; }
        public string file_name { get; set; }
        public int rev_no { get; set; }
        public string create_by { get; set; }
    }
}