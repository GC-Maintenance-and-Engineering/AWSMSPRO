using MeExpert.CommonProvide;
using MeExpert.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Windows.Forms;

namespace MeExpert.Controllers
{
    //[AuthenFilter]
    public class RequisitionController : Controller
    {
        // GET: Requisition
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Detail(int req_id,int client_id)
        {
            ViewBag.ReqID = req_id;
            ViewBag.ClientID = client_id;
            return View();
        }
        private string UserAuthen()
        {
            var usr_str = HttpContext.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            return usr_id[1];
        }
        public JsonResult GetReq(int req_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlDataAdapter da;
            SqlCommand cmd;
            List<Req> res = new List<Req>();
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                switch (req_id)
                {
                    case 0:
                        sql = $"select req_id,req_no,title,c.client_id,client_name,position_name,req_status,r.create_date,r.create_by from requisition as r inner join client as c on r.client_id = c.client_id " +
                              $"inner join position as p  on p.position_id = r.position_id order by r.create_date desc";
                        da = new SqlDataAdapter(sql, db.GetConn());
                        DataTable table = new DataTable();
                        da.Fill(table);
                        foreach (DataRow row in table.Rows)
                        {
                            string sel_cand = string.Empty;
                            sql = $"select s.cand_id,s.emp_id,c.first_name,c.last_name from req_success s " +
                                $"inner join candidate c on c.cand_id =s.cand_id where req_id=@req_id";
                            cmd = new SqlCommand(sql,db.GetConn());
                            cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = (int)row["req_id"];
                            dr = cmd.ExecuteReader();
                            while (dr.Read())
                            {
                                var emp = (dr["emp_id"] == DBNull.Value)? string.Empty : $"{dr["emp_id"]} : ";
                                sel_cand = $"{sel_cand} {emp} {dr["first_name"]} {dr["last_name"]},</br>";
                            }
                            dr.Close();
                            res.Add(new Req()
                            {
                                req_id = (int)row["req_id"],
                                req_no = $"{row["req_no"]}",
                                title = $"{row["title"]}",
                                client_id = (int)row["client_id"],
                                client = $"{row["client_name"]}",
                                position = $"{row["position_name"]}",
                                req_status = $"{row["req_status"]}",
                                sel_cand = sel_cand,
                                create_date = $"{Convert.ToDateTime($"{ row["create_date"] }", CultureInfo.CurrentCulture).ToString("yyyy'/'MM'/'dd")}",
                                create_by = $"{row["create_by"]},</br> {Convert.ToDateTime($"{row["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                            });
                        }
                        break;
                    default:
                        sql = $"select req_id, req_no, title, c.client_id, position_id, req_status, file_name, file_extension,headcounter,r.create_date,s_date,close_date,r.remark  from requisition as r inner join client as c on r.client_id = c.client_id where req_id=@req_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Parameters.Add("@req_id", SqlDbType.Int).Value =req_id;
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            res.Add(new Req()
                            {
                                req_id = (int)dr["req_id"],
                                req_no = $"{dr["req_no"]}",
                                title = $"{dr["title"]}",
                                client_id = (int)dr["client_id"],
                                position_id = (int)dr["position_id"],
                                req_status = $"{dr["req_status"]}",
                                file_name = (dr["file_name"] == DBNull.Value) ? string.Empty : $"{dr["file_name"]}",
                                file_extension = $"{dr["file_extension"]}",
                                headcounter = (int)dr["headcounter"],
                                create_date = Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                s_date = string.IsNullOrEmpty($"{dr["s_date"]}") ? string.Empty : Convert.ToDateTime($"{dr["s_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                close_date = string.IsNullOrEmpty($"{dr["close_date"]}") ? string.Empty : Convert.ToDateTime($"{dr["close_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                remark = string.IsNullOrEmpty($"{dr["remark"]}") ? string.Empty : $"{dr["remark"]}",
                            });
                        }
                        dr.Close();
                        break;
                }
              
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }

            return Json(res,JsonRequestBehavior.AllowGet);
        }
        public JsonResult Update(Req data,string type)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
           
            string rq_no = string.Empty;
            string sql = string.Empty;
            int count = 0;
            DateTime dt = DateTime.Now;
            string dtformat = dt.ToString("yyyy-MM-dd");
            int req_id = 0;
            try
            {
                db.OpenConn();
                switch (type)
                {
                    case "insert":
                        sql = $"select count(*) as count from requisition ";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            count = (int)dr["count"];
                        }
                        dr.Close();

                        count = count + 1;
                        rq_no = $"RQ-{DateTime.Now.Year.ToString()}-{DateTime.Now.Month.ToString()}-{String.Format("{0:D5}", count)}";

                        sql = $"insert into  requisition (req_no,title,client_id,position_id,file_name,file_extension,headcounter,req_status,create_by) " +
                            $" values(@req_no,@title, @client_id,@position_id,@file_name,@file_extension,@headcounter,@req_status,@create_by);select scope_identity()";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Parameters.Add("@req_no", SqlDbType.NVarChar, 50).Value = data.req_no;
                        cmd.Parameters.Add("@title", SqlDbType.NVarChar, 100).Value = data.title;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.Parameters.Add("@position_id", SqlDbType.Int).Value = data.position_id;
                        cmd.Parameters.Add("@file_name", SqlDbType.NVarChar, 50).Value = rq_no;
                        cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar, 5).Value = ".pdf";
                        cmd.Parameters.Add("@headcounter", SqlDbType.Int).Value = data.headcounter;
                        cmd.Parameters.Add("@req_status", SqlDbType.NVarChar, 10).Value = "Open";
                        cmd.Parameters.Add("@create_by", SqlDbType.NVarChar,13).Value = UserAuthen();
                        req_id = Convert.ToInt32(cmd.ExecuteScalar());
                        Management.InsertLog($"Requisition List", $" Create Requisition req_no={data.req_no}");
                        break;
                    case "update":
                        sql = $"update requisition set req_status=@req_status , headcounter=@headcounter, s_date=@s_date, remark=@remark," +
                            $" req_no=@req_no, title=@title, client_id=@client_id, position_id=@position_id, close_date=@close_date where req_id=@req_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                        cmd.Parameters.Add("@req_status", SqlDbType.NVarChar,10).Value = data.req_status;
                        cmd.Parameters.Add("@headcounter", SqlDbType.Int).Value = data.headcounter;
                        cmd.Parameters.Add("@req_no", SqlDbType.NVarChar,50).Value = data.req_no;
                        cmd.Parameters.Add("@title", SqlDbType.NVarChar,100).Value = data.title;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.Parameters.Add("@position_id", SqlDbType.Int).Value = data.position_id;
                        if (data.req_status == "Close")
                        {
                            cmd.Parameters.Add("@close_date", SqlDbType.Date).Value = dtformat;
                        }
                        else {
                            cmd.Parameters.AddWithValue("@close_date", DBNull.Value);
                        }
                        if (!string.IsNullOrEmpty(data.s_date))
                        {
                            cmd.Parameters.Add("@s_date", SqlDbType.Date).Value = data.s_date;
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@s_date", DBNull.Value);
                        }
                        cmd.Parameters.Add("@remark", SqlDbType.NVarChar,200).Value = string.IsNullOrEmpty(data.remark) ? string.Empty : data.remark;
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Requisition List", $"Update Requisition req_id={data.req_id}");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(req_id,JsonRequestBehavior.AllowGet);
        }
        //public ActionResult UploadFiles()
        //{
        //    string uname = Request["uploadername"];
        //    HttpFileCollectionBase files = Request.Files;
        //    var fullPath = string.Empty;
        //    try
        //    {
        //        for (int i = 0; i < files.Count; i++)
        //        {
        //            HttpPostedFileBase file = files[i];
        //            string fname;
        //            fname = uname + ".pdf";
        //            fullPath = Path.Combine(Server.MapPath("~/Uploads/"), fname);
        //            file.SaveAs(fullPath);
                   
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(fullPath);
        //        throw ex;
        //    }

        //    return Json("Hi, " + uname + ". Your files uploaded successfully", JsonRequestBehavior.AllowGet);
        //}
        public JsonResult CreatePosition(string posi_name)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            var posi_id = 0;
            string sql = string.Empty;
            bool chkname = false;
            try
            {
                db.OpenConn();
                sql = $"select * from position where position_name=@position_name";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@position_name", SqlDbType.NVarChar, 100).Value = posi_name;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    posi_id = (int)dr["position_id"];
                    chkname = true;
                }
                dr.Close();
                if (!chkname)
                {
                    sql = $"insert into position (position_name) values(@position_name);SELECT scope_identity();";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Parameters.Add("@position_name", SqlDbType.NVarChar, 100).Value = posi_name;
                    posi_id = Convert.ToInt32(cmd.ExecuteScalar());
                }
               
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return Json(posi_id,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult InsertCandOfReq(CandOfReq data)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            try
            {
                db.OpenConn();
                string sql = $"insert into req_cand (cand_id,req_id,create_by) values(@cand_id,@req_id,@create_by)";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                cmd.Parameters.Add("@Create_by", SqlDbType.Int).Value = UserAuthen();
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Candidate List",$" Create Candidate of Requisition req_id={data.req_id}, cand_id={data.cand_id}");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetCv(int req_cand_id,int req_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            SqlDataAdapter da;
            DataTable table;
            List<CandOfReq> res = new List<CandOfReq>();
            List<Interview> res_sub;
            List<ContactPerson> res_cont;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                switch (req_cand_id)
                {
                    case 0:
                        sql = $" select r.cand_id,r.req_cand_id,first_name,last_name,ap.apply_position,r.create_date,r.status,remark,sent_date,cv_return,r.remark as 'remark_status',cv_id ,rs.create_date as create_status  " +
                             $" from req_cand as r" +
                             $" inner join candidate as c on r.cand_id = c.cand_id" +
                             $" inner join  (select cand_id,cv_no, apply_position,create_time, row_number() " +
                             $" over(partition by cand_id order by create_time desc) as rn from apply_record ) as ap " +
                             $" on ap.cand_id = c.cand_id " +
                             $" inner join cv_details cv on cv.cv_no=ap.cv_no" +
                             $" left join req_success rs on rs.cand_id = r.cand_id  and rs.req_id=r.req_id" +
                             $" where ap.rn = 1 and r.req_id=@req_id" +
                             $" order by create_date ,cand_id asc";
                        da = new SqlDataAdapter(sql,db.GetConn());
                        da.SelectCommand.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                        table = new DataTable();
                        da.Fill(table);
                        foreach (DataRow row in table.Rows)
                        {
                            res_sub = new List<Interview>();
                            string sql_sub_list = $"select * from req_interview where req_cand_id = @req_cand_id order by interview_date asc";
                            da = new SqlDataAdapter(sql_sub_list, db.GetConn());
                            da.SelectCommand.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = (int)row["req_cand_id"];
                            DataTable tb = new DataTable();
                            da.Fill(tb);
                            foreach(DataRow item in tb.Rows)
                            {
                                res_cont = new List<ContactPerson>();
                                string sql_cont = $"select c.cont_person_name,c.position from interviewer i inner join contact_person c on i.cont_id=c.cont_id " +
                                    $"where interview_id=@interview_id";
                                cmd = new SqlCommand(sql_cont,db.GetConn());
                                cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = (int)item["interview_id"];
                                dr = cmd.ExecuteReader();
                                while (dr.Read())
                                {
                                    res_cont.Add(new ContactPerson() { 
                                        cont_person_name = $"{dr["cont_person_name"]}",
                                    });
                                }
                                dr.Close();
                                res_sub.Add(new Interview()
                                {
                                    interview_id = (int)item["interview_id"],
                                    interview_date = Convert.ToDateTime($"{item["interview_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                    contacts = res_cont,
                                }) ;
                            }
                            res.Add(new CandOfReq()
                            {
                                cv_id = (int)row["cv_id"],
                                cand_id = (int)row["cand_id"],
                                req_cand_id = (int)row["req_cand_id"],
                                cand_name = $"{row["first_name"]} {row["last_name"]}",
                                position = $"{row["apply_position"]}",
                                status = $"{row["status"]}",
                                create_status = (row["create_status"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["create_status"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                remark_status = string.IsNullOrEmpty($"{row["remark_status"]}") ? string.Empty : $"{row["remark_status"]}",
                                sent_date = (row["sent_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["sent_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                cv_return = (row["cv_return"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["cv_return"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                interviews = res_sub,
                            });
                           
                        }
                        break;
                    default:
                        sql = $"select c.req_id,c.cand_id,c.req_cand_id,interview_id,interview_remark,interview_date,status,remark,c.status," +
                            $" rs.enddate,c.remark as 'remark_status',c.sent_date,c.cv_return,cd.first_name,cd.last_name, cv.cv_id" +
                            $" from req_cand as c " +
                            $" left join candidate as cd on cd.cand_id = c.cand_id" +
                            $" left join req_interview as v  on v.req_cand_id = c.req_cand_id" +
                            $" outer apply (select top 1 enddate from req_success as s where s.req_id = c.req_id) as rs" +
                            $"  outer apply(select top 1 apply_position,apply_date,exp_salary,cv_no from apply_record as a"+ 
                            $" where c.cand_id = a.cand_id) as ap"+
                            $" inner join cv_details cv on cv.cv_no = ap.cv_no" +
                            $" where c.req_cand_id = @req_cand_id order by interview_date asc";
                        da = new SqlDataAdapter(sql,db.GetConn());
                        da.SelectCommand.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                        table = new DataTable();
                        da.Fill(table);

                        foreach (DataRow row in table.Rows)
                        {
                            res_sub = new List<Interview>();
                            if ((row["interview_id"]) != DBNull.Value && (int)row["interview_id"] > 0) 
                            {
                                sql = $"select * from interviewer where interview_id=@interview_id order by item_id asc";
                                cmd = new SqlCommand(sql, db.GetConn());
                                cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = (int)row["interview_id"];
                                dr = cmd.ExecuteReader();
                                while (dr.Read())
                                {
                                    res_sub.Add(new Interview()
                                    {
                                        item_id = (int)dr["item_id"],
                                        cont_id = (int)dr["cont_id"],
                                    });
                                }
                                dr.Close();
                            }

                            res.Add(new CandOfReq()
                            {
                                cand_id = (int)row["cand_id"],
                                interview_id = (row["interview_id"] == DBNull.Value) ? 0 : (int)row["interview_id"],
                                remark = (row["interview_remark"] == DBNull.Value) ? string.Empty : $"{row["interview_remark"]}",
                                interview_date = (row["interview_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["interview_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                status = (row["status"] == DBNull.Value) ? string.Empty : $"{row["status"]}",
                                enddate = (row["enddate"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["enddate"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                sent_date = (row["sent_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["sent_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                cv_return = (row["cv_return"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{row["cv_return"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                                remark_status = (row["remark_status"] == DBNull.Value) ? string.Empty : $"{row["remark_status"]}",
                                //cv_path = GetCvPath((int)row["cand_id"]),
                                cv_id = (int)row["cv_id"],
                                cand_name = $"{row["first_name"]} {row["last_name"]}",
                                interviews = res_sub,
                            });
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult UpdateInterview(List<CandOfReq> data,string sent_date,string cv_return, string remark_status,int req_cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            SqlTransaction transaction;
            string sql = string.Empty;
            List<Interview> res_sub;
            CultureInfo culture = new CultureInfo("en-US");
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            int intview_id = 0;
            try
            {
                if (data != null)
                {
                    //delete interview
                    List<CandOfReq> intview_list = new List<CandOfReq>();
                    sql = $"select interview_id from req_interview where req_cand_id=@req_cand_id";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Transaction = transaction;
                    cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = data[0].req_cand_id;
                    dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        intview_list.Add(new CandOfReq()
                        {
                            interview_id = (int)dr["interview_id"],
                        });
                    }
                    dr.Close();
                    var res = intview_list.Where(tab => !data.Any(up => up.interview_id == tab.interview_id));
                    var list_res = res.ToList();
                    if (list_res.Count > 0)
                    {
                        sql = "delete interviewer where interview_id in ({0});delete req_interview where  interview_id in ({0})";
                        string[] paramArr = list_res.Select((x, i) => "@list_in" + i).ToArray();
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.CommandText = string.Format(sql, string.Join(",", paramArr));
                        for (int i = 0; i < list_res.Count; i++)
                        {
                            cmd.Parameters.Add(new SqlParameter("@list_in" + i, list_res[i].interview_id));
                        }
                        cmd.ExecuteNonQuery();
                    }

                    for (int i = 0; i < data.Count; i++)
                    {
                        //insert and update interview
                        if (data[i].interview_id == 0)
                        {
                            //insert
                            sql = $"insert req_interview (req_cand_id,interview_remark,interview_date) values(@req_cand_id,@interview_remark,@interview_date);SELECT scope_identity();";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = data[i].req_cand_id;
                            cmd.Parameters.Add("@interview_remark", SqlDbType.NVarChar, 500).Value = (string.IsNullOrEmpty(data[i].interview_remark))? string.Empty : data[i].interview_remark;
                            cmd.Parameters.Add("@interview_date", SqlDbType.DateTime).Value = Convert.ToDateTime(data[i].interview_date, culture);
                            intview_id = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                        else
                        {
                            //update
                            sql = $"update req_interview set req_cand_id=@req_cand_id,interview_remark=@interview_remark,interview_date=@interview_date where interview_id=@interview_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = data[i].req_cand_id;
                            cmd.Parameters.Add("@interview_remark", SqlDbType.NVarChar, 500).Value = (string.IsNullOrEmpty(data[i].interview_remark)) ? string.Empty : data[i].interview_remark;
                            cmd.Parameters.Add("@interview_date", SqlDbType.DateTime).Value = Convert.ToDateTime(data[i].interview_date, culture);
                            cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = data[i].interview_id;
                            cmd.ExecuteNonQuery();
                            intview_id = data[i].interview_id;

                            //delete interviwer
                            List<Interview> intvwer = new List<Interview>();
                            sql = $"select item_id from interviewer where interview_id=@interview_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = intview_id;
                            dr = cmd.ExecuteReader();
                            while (dr.Read())
                            {
                                intvwer.Add(new Interview()
                                {
                                    item_id = (int)dr["item_id"],
                                });
                            }
                            dr.Close();
                            if (intvwer.Count > 0)
                            {
                                if (data[i].interviews == null)
                                {
                                    sql = "delete interviewer where interview_id = @interview_id";
                                    cmd = new SqlCommand(sql, db.GetConn());
                                    cmd.Transaction = transaction;
                                    cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = intview_id;
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    var res_intvwer = intvwer.Where(tab_wer => !data[i].interviews.Any(up_wer => up_wer.item_id == tab_wer.item_id));
                                    var list_res_intvwer = res_intvwer.ToList();
                                    if (list_res_intvwer.Count > 0)
                                    {
                                        sql = "delete interviewer where item_id in ({0})";
                                        string[] paramArrWer = list_res_intvwer.Select((x, n) => "@list_in_wer" + n).ToArray();
                                        cmd = new SqlCommand(sql, db.GetConn());
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = string.Format(sql, string.Join(",", paramArrWer));
                                        for (int n = 0; n < list_res_intvwer.Count; n++)
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@list_in_wer" + n, list_res_intvwer[n].item_id));
                                        }
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }

                        }

                        //insert and update interviwer by interview id 
                        res_sub = new List<Interview>();
                        if (data[i].interviews != null)
                        {
                            for (int j = 0; j < data[i].interviews.Count; j++)
                            {
                                if (data[i].interviews[j].item_id == 0)
                                {
                                    sql = $"insert into interviewer (interview_id,cont_id) values(@interview_id,@cont_id)";
                                    cmd = new SqlCommand(sql, db.GetConn());
                                    cmd.Transaction = transaction;
                                    cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = intview_id;
                                    //cmd.Parameters.Add("@interviewer_name", SqlDbType.NVarChar, 100).Value = data[i].interviews[j].interviewer_name;
                                    cmd.Parameters.Add("@cont_id", SqlDbType.Int).Value = data[i].interviews[j].cont_id;
                                    cmd.ExecuteNonQuery();
                                }
                                //else
                                //{
                                //    sql = $"update interviewer set interview_id=@interview_id,interviewer_name=@interviewer_name where item_id=@item_id";
                                //    cmd = new SqlCommand(sql, db.GetConn());
                                //    cmd.Transaction = transaction;
                                //    cmd.Parameters.Add("@item_id", SqlDbType.Int).Value = data[i].interviews[j].item_id;
                                //    cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = intview_id;
                                //    cmd.Parameters.Add("@interviewer_name", SqlDbType.NVarChar, 100).Value = data[i].interviews[j].interviewer_name;
                                //    cmd.ExecuteNonQuery();
                                //}
                            }
                        }
                    }
                    Management.InsertLog($"Requisition List", $"Update Interview req_cand_id={ data[0].req_cand_id}");
                }
                sql = $"update req_cand set remark=@remark,sent_date=@sent_date,cv_return=@cv_return where req_cand_id = @req_cand_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Transaction = transaction;
                if (string.IsNullOrEmpty(sent_date))
                    cmd.Parameters.AddWithValue("@sent_date", DBNull.Value);
                else
                    cmd.Parameters.Add("@sent_date", SqlDbType.Date).Value = sent_date;  
                if (string.IsNullOrEmpty(cv_return))
                    cmd.Parameters.AddWithValue("@cv_return", DBNull.Value);
                else
                    cmd.Parameters.Add("@cv_return", SqlDbType.Date).Value = cv_return;
                cmd.Parameters.Add("@remark", SqlDbType.NVarChar,200).Value = remark_status;
                cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetReqSucc(int req_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<ReqSucc> res = new List<ReqSucc>();
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select succ_id,s.cand_id,emp_id,c.first_name,c.last_name,enddate,s.create_date,rc.req_cand_id, r.s_date, s.s_work_date ,p.position_name, cv.cv_id " +
                    $"from req_success as s " +
                    $"inner join candidate as c on s.cand_id = c.cand_id " +
                    $"inner join requisition as r on s.req_id = r.req_id " +
                    $"inner join req_cand as rc on s.req_id = rc.req_id and s.cand_id = rc.cand_id " +
                    $"inner join position as p on p.position_id = r.position_id " +
                    $"outer apply(select top 1 apply_position,apply_date,exp_salary,cv_no from apply_record as a " +
                    $"where a.cand_id = s.cand_id) as ap " +
                    $"inner join cv_details cv on cv.cv_no = ap.cv_no " +
                    $"where s.req_id = @req_id order by create_date asc";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new ReqSucc() { 
                        cand_id = (int)dr["cand_id"],
                        cv_id = (int)dr["cv_id"],
                        succ_id = (int)dr["succ_id"],
                        position = (dr["position_name"] == DBNull.Value) ? string.Empty : $"{dr["position_name"]}",
                        cand_name = $"{dr["first_name"]} {dr["last_name"]}",
                        emp_id = $"{dr["emp_id"]}",
                        enddate =  (dr["enddate"] == DBNull.Value) ? string.Empty: Convert.ToDateTime($"{dr["enddate"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy") ,
                        create_date = (dr["create_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        req_cand_id = (dr["req_cand_id"] == DBNull.Value) ? 0 : (int)dr["req_cand_id"],
                        s_date = (dr["s_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["s_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        s_work_date = (dr["s_work_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["s_work_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                    });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult UpdateSucc(ReqSucc data,string type,string req_cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            string sql = string.Empty;
            int succ_count = 0;
            int headcounter = 0;
            SqlTransaction transaction;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            try
            {
                
                switch (type)
                {
                    case "insert":
                        sql = $"insert into req_success (req_id,cand_id,create_by) values(@req_id,@cand_id,@create_by)";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                        cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                        //cmd.Parameters.Add("@enddate", SqlDbType.Date).Value = data.enddate;
                        cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                        cmd.ExecuteNonQuery();
                        
                        sql = $"update req_cand set status = 'pass' where req_cand_id=@req_cand_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Requisition Success List", $" Create Requisition req_id={data.req_id}, cand_id={data.cand_id}");

                        sql = $"select count(succ_id) as count from req_success where req_id=@req_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            succ_count = (int)dr["count"];
                        }
                        dr.Close();

                        sql = $"select headcounter from requisition where req_id=@req_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            headcounter = (int)dr["headcounter"];
                        }
                        dr.Close();
                        if (headcounter == succ_count)
                        {
                            sql = $"update requisition set req_status=@req_status, close_date=@close_date where req_id=@req_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = data.req_id;
                            cmd.Parameters.Add("@req_status", SqlDbType.NVarChar,10).Value = "Close";
                            cmd.Parameters.Add("@close_date", SqlDbType.Date).Value = DateTime.Now.ToString("yyyy-MM-dd");
                            cmd.ExecuteNonQuery();
                        }

                        break;
                    case "delete":
                        sql = $"delete proposal_list where succ_id=@succ_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                        cmd.ExecuteNonQuery();

                        sql = $"delete req_success where succ_id=@succ_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                        cmd.ExecuteNonQuery();

                        sql = $"update req_cand set status=@status where req_cand_id=@req_cand_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                        cmd.Parameters.AddWithValue("@status", DBNull.Value);
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Requisition Success List", $" Delete Requisition succ_id={data.succ_id}");
                        break;
                    case "update":
                        sql = $"update req_success set s_work_date =@s_work_date where succ_id =@succ_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                        cmd.Parameters.Add("@s_work_date", SqlDbType.Date).Value = data.s_work_date;
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Requisition Success List", $" Update Requisition succ_id={data.succ_id}");
                        break;
                    case "updateempid":
                        sql = $"update req_success set emp_id=@emp_id where succ_id=@succ_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                        if (string.IsNullOrEmpty(data.emp_id))
                            cmd.Parameters.AddWithValue("@emp_id", DBNull.Value);
                        else
                            cmd.Parameters.Add("@emp_id", SqlDbType.NVarChar, 13).Value = data.emp_id;
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Requisition Success List", $" Update Emp ID succ_id={data.succ_id}");
                        break;
                }
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult UpdateSuccDate(int req_id, string enddate, string s_work_date)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            
            try
            {
                db.OpenConn();
                string sql = $"update req_success set enddate=@enddate where req_id=@req_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                cmd.Parameters.Add("@enddate", SqlDbType.Date).Value = enddate;
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Requisition Success List", $"Update Requisition Success Date req_id={req_id}");
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult DelCand(int req_cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            SqlTransaction transaction;
            List<int> inview_list = new List<int>();
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            try
            {
               
                string sql = $"select interview_id from req_interview where req_cand_id=@req_cand_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    inview_list.Add((int)dr["interview_id"]);
                }
                dr.Close();
                if (inview_list.Count > 0)
                {
                    sql = "delete interviewer where interview_id in ({0})";
                    string[] pramArr = inview_list.Select((x, i) => "@inview_list" + i).ToArray();
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Transaction = transaction;
                    cmd.CommandText = string.Format(sql, string.Join(",", pramArr));
                    for (int i = 0; i < inview_list.Count; i++)
                    {
                        cmd.Parameters.Add(new SqlParameter("@inview_list" + i, inview_list[i]));
                    }
                    cmd.ExecuteNonQuery();
                }

                sql = $"delete req_interview where req_cand_id=@req_cand_id;delete req_cand where req_cand_id=@req_cand_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = req_cand_id;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                Management.InsertLog($"Candidate in Requisition", $"Delete Candidate in Requisition req_cand_id={req_cand_id} ");
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public string GetCvPath(int cand_id)
        {
            string cvpath = string.Empty;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;

            try
            {
                db.OpenConn();
                string sql = $"select top 1 a.cv_no,apply_date ,c.srv_doc_path,doc_extension " +
                    $"from apply_record as a " +
                    $"inner join cv_details as c on a.cv_no = c.cv_no " +
                    $"where cand_id = @cand_id order by apply_date desc";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    string srv_doc_path = $"{dr["srv_doc_path"]}";
                    cvpath = DeCodeBase64(srv_doc_path);
                }
                
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return cvpath;
        }
        public string DeCodeBase64(string encodePath)
        {
            string decodeStr = null;
            byte[] data = System.Convert.FromBase64String(encodePath);
            decodeStr = System.Text.ASCIIEncoding.ASCII.GetString(data);
            return decodeStr;
        }
        public JsonResult GetReqNo(string req_no) 
        {
            bool res = false;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;

            try
            {
                db.OpenConn();
                string sql = $"select req_no from requisition where req_no=@req_no";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_no",SqlDbType.NVarChar,50).Value = req_no;
                dr = cmd.ExecuteReader();
                if (dr.Read()) {
                    res = true;
                }
                dr.Close();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        public ActionResult DownloadProposal(int proposal_id)
        {
            List<FilePPS> filePPs = new List<FilePPS>();
            MemoryStream stream = new MemoryStream();
            string position = string.Empty;
            string rev_no = string.Empty;
            filePPs = CreateWordByOpenXML.GetDetailProposal(proposal_id);
            stream = filePPs[0].stream;
            position = filePPs[0].position.ToLower();
            rev_no = filePPs[0].rev_no;

            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"proposal-{position}-rev{rev_no}.docx");
         }
        [HttpPost]
        public JsonResult UpdateProposal(Proposal data)
        {
            string file_name = string.Empty;
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            int rev = 0;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select top 1 rev_no from proposal_list where succ_id=@succ_id and  pps_official_file is null order by rev_no desc";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    rev = (int)dr["rev_no"];
                    rev++;
                }
                dr.Close();

                DateTime dateTime = DateTime.UtcNow.Date;
                //var a = dateTime.ToString("ddMMyy");
                //$"PPS-{DateTime.Now.Year.ToString()}-{DateTime.Now.Month.ToString()}-{String.Format("{0:D5}", count)}-{data.position}";
                file_name = $"proposal-{data.position.ToLower()}-rev{rev}";
                    
                sql = $"insert into proposal_list (succ_id,rev_no,file_name,start_date,end_date,working_date,proposal_date,scope_service,act_salary,period,pps_comp,pps_license,create_by) " +
                    $"values(@succ_id,@rev_no,@file_name,@start_date,@end_date,@working_date,@proposal_date,@scope_service,@act_salary,@period,@pps_comp,@pps_license,@create_by)";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                cmd.Parameters.Add("@rev_no", SqlDbType.Int).Value = rev;
                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar, 50).Value = file_name;
                cmd.Parameters.Add("@start_date", SqlDbType.Date).Value = data.start_date;
                cmd.Parameters.Add("@end_date", SqlDbType.Date).Value = data.end_date;
                cmd.Parameters.Add("@working_date", SqlDbType.Int).Value = data.working_date;
                cmd.Parameters.Add("@proposal_date", SqlDbType.Date).Value = data.proposal_date;
                cmd.Parameters.Add("@scope_service", SqlDbType.NVarChar, 500).Value = data.scope_service;
                cmd.Parameters.Add("@act_salary", SqlDbType.Decimal).Value = data.act_salary;
                cmd.Parameters.Add("@pps_comp", SqlDbType.Decimal).Value = data.pps_comp;
                cmd.Parameters.Add("@pps_license", SqlDbType.Decimal).Value = data.pps_license;
                cmd.Parameters.Add("@period", SqlDbType.Int).Value = data.period;
                cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Proposal", $"Create Proposal file_name={file_name} succ_id={ data.succ_id}");
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetProposalBySuccID(int succ_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Proposal> res = new List<Proposal>();
            try
            {
                db.OpenConn();
                string sql = $"select * from proposal_list where succ_id = @succ_id and pps_official_file is null";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = succ_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Proposal() { 
                        proposal_id = (int)dr["proposal_id"],
                        rev_no = (int)dr["rev_no"],
                        file_name = $"{dr["file_name"]}",
                        start_date = (dr["start_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["start_date"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        end_date = (dr["end_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["end_date"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        working_date = (dr["working_date"] == DBNull.Value) ? 0 : (int)dr["working_date"],
                        proposal_date = (dr["proposal_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["proposal_date"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        scope_service = (dr["scope_service"] == DBNull.Value) ? string.Empty : $"{dr["scope_service"]}",
                        act_salary = (dr["act_salary"] == DBNull.Value) ? 0 : (decimal)dr["act_salary"],
                        pps_comp = (dr["pps_comp"] == DBNull.Value) ? 0 : (decimal)dr["pps_comp"],
                        pps_license = (dr["pps_license"] == DBNull.Value) ? 0 : (decimal)dr["pps_license"],
                        create_by = $"{dr["create_by"]}, {Convert.ToDateTime($"{dr["create_date"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                    });
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult SaveFileStream()
        {
            string req_id = Request["req_id"];
            string req_no = Request["req_no"];
            HttpFileCollectionBase files = Request.Files;
            string fullPath = string.Empty;
         
            try
            {
                if (files.Count != 0)
                {
                    for (int i=0; i <files.Count;i++)
                    {
                        //var file = files[i];
                        //Stream file_stream = file.InputStream;
                        //byte[] filestream = ConvertFileSream.ConvertFileStream(file_stream);
                        //ConvertFileSream.SaveReqFile(filestream,Convert.ToInt32(req_id));

                        //upload file
                        HttpPostedFileBase file = files[i];
                        string file_name = files[i].FileName;
                        float file_size = files[i].ContentLength;
                        string doc_extension = Path.GetExtension(file_name);
                        string doc_name = "REQ" + "-" + req_no.Replace(@"/", "-") + doc_extension;
                        fullPath = Path.Combine(Server.MapPath("~/Uploads/REQ/"), doc_name);
                        file.SaveAs(fullPath);

                        //update database
                        ConvertFileSream.SaveReqFile(doc_name, doc_extension, Convert.ToInt32(req_id));
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public ActionResult OpenReqFile(string req_id)
        {
            int int_req_id = Management.EncodeToNumber(req_id);
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            string path = string.Empty;
            string file_name = string.Empty;
            string file_extension = string.Empty;
            string type = "application/pdf";
            byte[] res_file = new byte[0];
            try
            {
                db.OpenConn();
                sql = $"select file_name,file_extension from requisition where req_id=@req_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = int_req_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    file_name = (dr["file_name"] == DBNull.Value) ? string.Empty : $"{dr["file_name"]}";
                    file_extension = $"{dr["file_extension"]}";
                }
                dr.Close();
                if (!string.IsNullOrEmpty(file_name))
                {
                    path = Path.Combine(Server.MapPath("~/Uploads/REQ/"), file_name);
                    res_file = System.IO.File.ReadAllBytes(path);
                    type = ReadFileType.GetTypeFile(file_extension);
                }
               
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return File(res_file, type);
        }
        public bool ConvertReqFile() {
            bool res = false;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            SqlDataAdapter da;
            SqlTransaction transaction;
            string sql = string.Empty;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            try
            {
               
                sql = $"select file_name,req_id from requisition where req_file IS Null and file_name IS NOT NULL";
                da = new SqlDataAdapter(sql,db.GetConn());
                da.SelectCommand.Transaction = transaction;
                DataTable table = new DataTable();
                da.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    byte[] fileStream;
                    string path = "http://gcgpgcmeapp01/MSPRO/Uploads/" + row["file_name"]+".pdf";

                    var client = new WebClient();
                    client.UseDefaultCredentials = true;
                    client.Credentials = new NetworkCredential("98007756", "Jane100121*");
                    //var request = (HttpWebRequest)WebRequest.Create(path);
                    //var response = (HttpWebResponse)request.GetResponse();
                    //if (response.CharacterSet != null)
                    //{
                        var content = client.DownloadData(path);
                        var stream = new MemoryStream(content);
                        byte[] file;
                        if (stream.Length != 0)
                        {
                            file = new byte[stream.Length];
                            stream.Read(file, 0, file.Length);
                            fileStream = file;

                            sql = $"update requisition set req_file=@req_file, file_name=@file_name where req_id=@req_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = (int)row["req_id"];
                            cmd.Parameters.Add("@req_file", SqlDbType.VarBinary).Value = fileStream;
                            cmd.Parameters.AddWithValue("@file_name", DBNull.Value);
                            cmd.Transaction = transaction;
                            cmd.ExecuteNonQuery();
                        }
                   // }
                   
                }
                transaction.Commit();
                res = true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return res;
        }
        public JsonResult DelReqByReqID(int req_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataAdapter da;
            SqlCommand cmd;
            SqlTransaction transaction;
            string sql = string.Empty;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            try
            {
                //del proposal
                sql = $"select succ_id from req_success where req_id=@req_id";
                da = new SqlDataAdapter(sql,db.GetConn());
                da.SelectCommand.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                da.SelectCommand.Transaction = transaction;
                DataTable table = new DataTable();
                da.Fill(table);
                foreach (DataRow row in table.Rows) {
                    sql = $"delete proposal_list where succ_id=@succ_id";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = (int)row["succ_id"];
                    cmd.Transaction = transaction;
                    cmd.ExecuteNonQuery();
                }
               
                //del succ
                sql = $"delete req_success where req_id=@req_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();


                
                sql = $"select req_cand_id from req_cand where req_id=@req_id";
                da = new SqlDataAdapter(sql, db.GetConn());
                da.SelectCommand.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                da.SelectCommand.Transaction = transaction;
                table = new DataTable();
                da.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    sql = $"select interview_id from req_interview where req_cand_id=@req_cand_id";
                    da = new SqlDataAdapter(sql, db.GetConn());
                    da.SelectCommand.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = (int)row["req_cand_id"];
                    da.SelectCommand.Transaction = transaction;
                    table = new DataTable();
                    da.Fill(table);
                    foreach (DataRow row_intw in table.Rows)
                    {
                        //del interviewer
                        sql = $"delete interviewer where interview_id=@interview_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Parameters.Add("@interview_id", SqlDbType.Int).Value = (int)row_intw["interview_id"];
                        cmd.Transaction = transaction;
                        cmd.ExecuteNonQuery();
                    }
                    //del interview
                    sql = $"delete req_interview where req_cand_id=@req_cand_id";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Parameters.Add("@req_cand_id", SqlDbType.Int).Value = (int)row["req_cand_id"];
                    cmd.Transaction = transaction;
                    cmd.ExecuteNonQuery();
                }
                //del req_cand
                sql = $"delete req_cand where req_id=@req_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();

                //del req
                sql = $"delete requisition where req_id=@req_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Requisition", $" Delete req_id={req_id} ");
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetPositionList()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Position> res = new List<Position>();

            try
            {
                db.OpenConn();
                string sql = $"select * from position";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Position()
                    {
                        position_id = (int)dr["position_id"],
                        position_name = $"{dr["position_name"]}",
                    });
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(res,JsonRequestBehavior.AllowGet) ;
        }
        #region Report Req
        public List<Report> GetReport(int year)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Report> res = new List<Report>();
            int n = 0;
            bool check_req = false;
            int req_id = 0;
            try
            {
                db.OpenConn();
                string sql = $"select  r.req_id,r.s_date,position_name,client_name,cd.first_name,last_name,cd.cand_remarks,rc.sent_date,rc.cv_return ,i.interview_date,rs.s_work_date " +
                $"from requisition r inner join position p on p.position_id = r.position_id  " +
                $"inner join client c on c.client_id = r.client_id " +
                $"inner join req_cand rc on rc.req_id = r.req_id  " + 
                $"inner join candidate cd on cd.cand_id = rc.cand_id " +
                $"outer apply(select top 1 interview_date from req_interview where req_cand_id = rc.req_cand_id order by interview_date asc) as i " +
                $"left join req_success rs on rs.req_id = r.req_id and rs.cand_id = rc.cand_id " +
                $"where year(r.create_date) = @year order by s_date desc";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@year",SqlDbType.Int).Value = year;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    //---------check-po---------//
                    if (req_id == (int)dr["req_id"] && n != 0)
                    {
                        check_req = true;
                    }
                    else
                    {
                        n++;
                        check_req = false;
                        req_id = (int)dr["req_id"];
                    }
                    
                    int cv_days = 0;
                    if (!Convert.IsDBNull(dr["s_date"]) && !Convert.IsDBNull(dr["sent_date"]))
                    {
                        cv_days = ((DateTime)dr["sent_date"] - (DateTime)dr["s_date"]).Days;
                    }
                    res.Add(new Report()
                    {
                        no = (check_req) ? string.Empty : n.ToString(),
                        date_job = (check_req) ? string.Empty : Convert.IsDBNull(dr["s_date"]) ? string.Empty : Convert.ToDateTime($"{dr["s_date"]}", CultureInfo.CurrentCulture).ToString("dd/MMM/yyyy"),
                        position = (check_req) ? string.Empty : Convert.IsDBNull(dr["position_name"]) ? string.Empty : $"{dr["position_name"]}",
                        area = string.Empty,
                        client = (check_req) ? string.Empty : $"{dr["client_name"]}",
                        candidate = $"{dr["first_name"]} {dr["last_name"]}",
                        remark = Convert.IsDBNull(dr["cand_remarks"]) ? string.Empty : $"{dr["cand_remarks"]}",
                        cv_sent = Convert.IsDBNull(dr["sent_date"]) ? string.Empty : Convert.ToDateTime($"{dr["sent_date"]}", CultureInfo.CurrentCulture).ToString("dd/MMM/yyyy"),
                        sent_days = cv_days,
                        cv_return = Convert.IsDBNull(dr["cv_return"]) ? string.Empty : Convert.ToDateTime($"{dr["cv_return"]}", CultureInfo.CurrentCulture).ToString("dd/MMM/yyyy"),
                        interview = Convert.IsDBNull(dr["interview_date"]) ? string.Empty : Convert.ToDateTime($"{dr["interview_date"]}", CultureInfo.CurrentCulture).ToString("dd/MMM/yyyy"),
                        salary = string.Empty,
                        health = string.Empty,
                        submit = string.Empty,
                        starting_date = (check_req) ? string.Empty : Convert.IsDBNull(dr["s_work_date"]) ? string.Empty : Convert.ToDateTime($"{dr["s_work_date"]}", CultureInfo.CurrentCulture).ToString("dd/MMM/yyyy"),
                    }); 
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return res;
        }
        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
        public ActionResult ExportExcel(int year)
        {
            try
            {
                List<Report> req_list = GetReport(year);
                DataTable table = ToDataTable(req_list);


                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var mmrStream = new MemoryStream();
                var exc = new ExcelPackage(mmrStream);
                var worksheet = exc.Workbook.Worksheets.Add("Requisition Report(" + year+ ")");
               
                worksheet.Cells["A2"].LoadFromDataTable(table, false);
                worksheet.Cells["A1:AN1"].Style.Font.Bold = true;
                //---change column name---//
                worksheet.Cells["A1"].Value = "No.";
                worksheet.Cells["B1"].Value = "Date job in";
                worksheet.Cells["C1"].Value = "Position";
                worksheet.Cells["D1"].Value = "Area";
                worksheet.Cells["E1"].Value = "User/Name";
                worksheet.Cells["F1"].Value = "Candidate's Name";
                worksheet.Cells["G1"].Value = "Remark";
                worksheet.Cells["H1"].Value = "CV Sent";
                worksheet.Cells["I1"].Value = "Sent Cvs By(Days)";
                worksheet.Cells["J1"].Value = "CV Return";
                worksheet.Cells["K1"].Value = "Interview";
                worksheet.Cells["L1"].Value = "Salary Negotiation";
                worksheet.Cells["M1"].Value = "Health Check Up";
                worksheet.Cells["N1"].Value = "Submit Proposal";
                worksheet.Cells["O1"].Value = "Starting Date";

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                for (int col = 1; col <= end.Column; col++)
                {
                    for (int row = 1; row <= end.Row; row++)
                    {
                        if (col ==2  || col == 8 || col == 9 || col == 10 || col == 11 || col == 15)
                        {
                            string act_date = (worksheet.Cells[row, col].Value != null) ? worksheet.Cells[row, col].Value.ToString() : string.Empty;
                            if (act_date != string.Empty)
                            {
                                DateTime temp;
                                if (DateTime.TryParse(act_date, out temp))
                                {
                                    worksheet.Cells[row, col].Value = DateTime.ParseExact(act_date, "dd/MMM/yyyy", CultureInfo.InvariantCulture).ToOADate();
                                    worksheet.Cells[row, col].Style.Numberformat.Format = "dd/MM/yyyy";
                                    worksheet.Column(col).Width = 12;
                                }
                            }
                        }

                    }
                }
                worksheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(9).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.DefaultColWidth = 20;
                worksheet.Column(3).AutoFit();
                worksheet.Column(5).AutoFit();
                worksheet.Column(6).AutoFit();
                worksheet.Column(7).AutoFit();
                worksheet.Calculate();

                Session["DownloadExcel_FileReport"] = exc.GetAsByteArray();
                return Json("", JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public ActionResult ExcelDownload()
        {
            if (Session["DownloadExcel_FileReport"] != null)
            {
                byte[] data = Session["DownloadExcel_FileReport"] as byte[];
                return File(data, "application/octet-stream", "Requisition-Report.xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }
        #endregion

    }

}