//using DocumentFormat.OpenXml.Drawing;
using MeExpert.CommonProvide;
using MeExpert.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace MeExpert.Controllers
{
    public class CandidateController : Controller
    {
        // GET: Candidate
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ManageCandidate()
        {
            return View();
        }
        public ActionResult Detail(int cand_id)
        {
            ViewBag.CandID = cand_id;
            return View();
        }
        public ActionResult ProposalDetail(string file_name)
        {
            ViewBag.FileName = file_name;
            return View();
        }
        public JsonResult Get()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Cand> res = new List<Cand>();
            var cont = new RequisitionController();
            try
            {
                db.OpenConn();
                string sql = $"select r.cand_id,c.first_name,c.last_name,ap.apply_position,ap.apply_date,rs.succ_id,rs.emp_id,rq.req_no,p.position_name,rs.s_work_date,pps.file_name as proposal ,cv.cv_id " +
                    $" from req_cand as r inner" +
                    $" join candidate as c on c.cand_id = r.cand_id" +
                    $" outer apply(select top 1 apply_position,apply_date,exp_salary,cv_no from apply_record as a" +
                    $" where c.cand_id = a.cand_id) as ap" +
                    $" left join(select cand_id, succ_id, create_date, emp_id, req_id, s_work_date, row_number() over(partition by cand_id order by create_date desc) as rn from req_success) as rs" +
                    $" on c.cand_id = rs.cand_id" +
                    $" inner join requisition as rq on rq.req_id = rs.req_id" +
                    $" inner join position as p on p.position_id = rq.position_id" +
                    $" outer apply(select top 1 file_name,succ_id from proposal_list as pp where rs.succ_id = pp.succ_id) as pps " +
                    $" inner join cv_details cv on cv.cv_no=ap.cv_no" +
                    $" where rs.succ_id is not null and rs.rn = 1" +
                    $" group by r.cand_id ,c.first_name,c.last_name,ap.apply_position,ap.apply_date,rs.succ_id,rs.emp_id,rq.req_no,p.position_name,rs.s_work_date,pps.file_name,cv.cv_id,rs.create_date order by rs.create_date desc";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string pass = "";
                    if (dr["succ_id"] != DBNull.Value)
                    {
                        pass = "Pass";
                    }
                    res.Add(new Cand()
                    {
                        cand_id = (int)dr["cand_id"],
                        full_name = $"{dr["first_name"]} {dr["last_name"]}",
                        apply_position = $"{dr["position_name"]}",
                        apply_date = string.IsNullOrEmpty($"{dr["s_work_date"]}") ? string.Empty : $"{Convert.ToDateTime($"{dr["s_work_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                        //cv_path = cont.GetCvPath((int)dr["cand_id"]),
                        cv_id = (int)dr["cv_id"],
                        emp_id = $"{dr["emp_id"]}",
                        status = pass,
                        req_no = $"{dr["req_no"]}",
                        proposal = (dr["proposal"] == DBNull.Value) ? string.Empty : $"{dr["proposal"]}",
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
            return Json(res, JsonRequestBehavior.AllowGet);
        }
        private string UserAuthen()
        {
            var usr_str = HttpContext.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            return usr_id[1];
        }
        [HttpPost]
        public ActionResult UploadFiles()
        {
            string uname = string.Empty;
            uname = Request["uploadername"];
            HttpFileCollectionBase files = Request.Files;
            var fullPath = string.Empty;
            try
            {
                if (files.Count != 0)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        string fname;
                        fname = uname + ".pdf";
                        fullPath = Path.Combine(Server.MapPath("~/Uploads/Proposal/"), fname);
                        file.SaveAs(fullPath);
                    }
                }
                else
                {
                    throw new Exception("");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(fullPath);
                throw ex;
            }

            return Json("Hi, " + uname + ". Your files uploaded successfully", JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult UpdateProposal(Proposal data)
        {
            string file_name = string.Empty;
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            int count = 0;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select count(*) as count from proposal_list ";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    count = (int)dr["count"];
                }
                dr.Close();

                count = count + 1;

                DateTime dateTime = DateTime.UtcNow.Date;
                //var a = dateTime.ToString("ddMMyy");
                file_name = $"PPS-{DateTime.Now.Year.ToString()}-{DateTime.Now.Month.ToString()}-{String.Format("{0:D5}", count)}-{data.position}";
                sql = $"insert into proposal_list (succ_id,file_name,file_extension,create_by) values(@succ_id, @file_name, @file_extension, @create_by)";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = data.succ_id;
                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar, 50).Value = file_name;
                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar, 5).Value = $".pdf";
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
            return Json(file_name, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetReqByCandID(int id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Req> res = new List<Req>();
            string sql = string.Empty;
            List<PpsReq> res_sub;
            SqlDataAdapter da;
            DataTable table;
            try
            {
                db.OpenConn();
                sql = $"select r.req_id,req_no,title,p.position_name,r.req_status,s.succ_id, r.client_id from requisition as r " +
                            $"inner join position as  p on r.position_id = p.position_id inner join req_cand as c on c.req_id = r.req_id " +
                            $"inner join req_success as s on s.cand_id = c.cand_id and s.req_id = r.req_id " +
                            $"where c.cand_id =  @id";
                da = new SqlDataAdapter(sql, db.GetConn());
                da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id;
                table = new DataTable();
                da.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    res_sub = new List<PpsReq>();
                    sql = $"select proposal_id,file_name,create_by,create_date from proposal_list where succ_id = @succ_id";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = (int)row["succ_id"];
                    dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        res_sub.Add(new PpsReq()
                        {
                            proposal_id = (int)dr["proposal_id"],
                            file_name = $"{dr["file_name"]}",
                            create_by = $"{dr["create_by"]}, {Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                        });
                    }
                    dr.Close();
                    res.Add(new Req()
                    {
                        req_id = (int)row["req_id"],
                        client_id = (int)row["client_id"],
                        req_no = $"{row["req_no"]}",
                        title = $"{row["title"]}",
                        position = $"{row["position_name"]}",
                        req_status = $"{row["req_status"]}",
                        succ_id = (int)row["succ_id"],
                        propoSals = res_sub
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
            return Json(res, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetCand()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Cand> res = new List<Cand>();
            string sql = string.Empty;
            var cont = new RequisitionController();
            int no = 1;
            try
            {
                db.OpenConn();
                sql = $"select c.cand_id,c.first_name,c.last_name,c.gender,c.mobile_no_1,c.mail,rc.apply_position,rc.remarks,rc.create_by,rc.create_time, DATEDIFF(yy, c.birth_day, getdate()) as age, cv.cv_id from candidate as c " +
                    $"outer apply(select top 1 apply_position, remarks, cv_no,create_by, create_time from apply_record as r where r.cand_id = c.cand_id order by create_time desc ) as rc left join cv_details cv on cv.cv_no=rc.cv_no order by c.create_time desc";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Cand()
                    {
                        no = no++,
                        cand_id = (int)dr["cand_id"],
                        full_name = $"{dr["first_name"]} {dr["last_name"]}",
                        gender = $"{dr["gender"]}",
                        age = (dr["age"] == DBNull.Value) ? string.Empty : $"{dr["age"]}",
                        last_apply_position = (dr["apply_position"] == DBNull.Value) ? "No Record" : $"{dr["apply_position"]}",
                        remarks = (dr["remarks"] == DBNull.Value) ? string.Empty : $"{dr["remarks"]}",
                        mobile = $"{dr["mobile_no_1"]}",
                        mail = (dr["mail"] == DBNull.Value) ? string.Empty : $"{dr["mail"]}",
                        create_date = string.IsNullOrEmpty($"{dr["create_time"]}") ? string.Empty : Convert.ToDateTime($"{dr["create_time"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        create_by = (dr["create_by"] == DBNull.Value) ? string.Empty : $"{dr["create_by"]}",
                        cv_id = (dr["cv_id"] == DBNull.Value) ? 0 : (int)dr["cv_id"],
                        //cv_file = (dr["cv_file"] == DBNull.Value) ? false : true,
                        //cv_path = cont.GetCvPath((int)dr["cand_id"]),
                    });
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
            return Json(res, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult UpdateCand(Candidate data, string type)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            string sql = string.Empty;
            SqlTransaction transaction;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            DateTime dt = DateTime.Now;

            string dtformat = dt.ToString("yyyy-MM-dd HH:mm:ss.fff");
            int cand_id = 0;
            try
            {
                switch (type)
                {
                    case "insert":
                        sql = $"INSERT INTO candidate" +
                    $"(first_name, last_name, birth_day, gender, mail, " +
                    $"mobile_no_1, mobile_no_2, toeic_score, start_working, employ_status, cand_remarks, id_card, create_by, create_time) " +
                    $"VALUES( @first_name, @last_name, @birth_day, @gender, @mail, " +
                    $"@mobile_no_1, @mobile_no_2, @toeic_score, @start_working, @employ_status, @cand_remarks, @id_card, @create_by, @create_time);SELECT scope_identity(); ";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@first_name", SqlDbType.NVarChar, 65).Value = data.first_name;
                        cmd.Parameters.Add("@last_name", SqlDbType.NVarChar, 65).Value = data.last_name;
                        if (string.IsNullOrEmpty(data.birth_day))
                            cmd.Parameters.AddWithValue("@birth_day", DBNull.Value);
                        else
                            cmd.Parameters.Add("@birth_day", SqlDbType.NVarChar, 15).Value = data.birth_day;
                        cmd.Parameters.Add("@gender", SqlDbType.NVarChar, 6).Value = data.gender;
                        if (string.IsNullOrEmpty(data.mail))
                            cmd.Parameters.AddWithValue("@mail", DBNull.Value);
                        else
                            cmd.Parameters.Add("@mail", SqlDbType.NVarChar, 50).Value = data.mail;
                        cmd.Parameters.Add("@mobile_no_1", SqlDbType.NVarChar, 13).Value = data.mobile_no_1;
                        if (string.IsNullOrEmpty(data.mobile_no_2))
                            cmd.Parameters.AddWithValue("@mobile_no_2", DBNull.Value);
                        else
                            cmd.Parameters.Add("@mobile_no_2", SqlDbType.NVarChar, 13).Value = data.mobile_no_2;
                        if (string.IsNullOrEmpty(data.toeic_score))
                            cmd.Parameters.AddWithValue("@toeic_score", DBNull.Value);
                        else
                            cmd.Parameters.Add("@toeic_score", SqlDbType.NVarChar, 6).Value = data.toeic_score;
                        if (string.IsNullOrEmpty(data.start_working))
                            cmd.Parameters.AddWithValue("@start_working", DBNull.Value);
                        else
                            cmd.Parameters.Add("@start_working", SqlDbType.NVarChar, 15).Value = data.start_working;
                        if (string.IsNullOrEmpty(data.employ_status))
                            cmd.Parameters.AddWithValue("@employ_status", DBNull.Value);
                        else
                            cmd.Parameters.Add("@employ_status", SqlDbType.NVarChar, 10).Value = data.employ_status;
                        if (string.IsNullOrEmpty(data.cand_remarks))
                            cmd.Parameters.AddWithValue("@cand_remarks", DBNull.Value);
                        else
                            cmd.Parameters.Add("@cand_remarks", SqlDbType.NVarChar, 512).Value = data.cand_remarks;
                        if (string.IsNullOrEmpty(data.id_card))
                            cmd.Parameters.AddWithValue("@id_card", DBNull.Value);
                        else
                            cmd.Parameters.Add("@id_card", SqlDbType.NVarChar, 13).Value = data.id_card;

                        cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                        cmd.Parameters.Add("@create_time", SqlDbType.DateTime).Value = dtformat;
                        cand_id = Convert.ToInt32(cmd.ExecuteScalar());
                        Management.InsertLog($"Candidate", $"Insert Candidate full_name={data.first_name} {data.last_name} cand_id={ cand_id}");
                        break;
                    case "update":
                        sql = $"update candidate set id_card = @id_card,first_name = @first_name,last_name = @last_name,birth_day = @birth_day,gender = @gender,mail = @mail,mobile_no_1 = @mobile_no_1," +
                        $" mobile_no_2 = @mobile_no_2,toeic_score = @toeic_score,start_working = @start_working,employ_status = @employ_status,cand_remarks = @cand_remarks,create_by = @create_by,create_time = @create_time where cand_id = @cand_id; ";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                        cmd.Parameters.Add("@first_name", SqlDbType.NVarChar, 65).Value = data.first_name;
                        cmd.Parameters.Add("@last_name", SqlDbType.NVarChar, 65).Value = data.last_name;
                        if (string.IsNullOrEmpty(data.birth_day))
                            cmd.Parameters.AddWithValue("@birth_day", DBNull.Value);
                        else
                            cmd.Parameters.Add("@birth_day", SqlDbType.NVarChar, 15).Value = data.birth_day;
                        cmd.Parameters.Add("@gender", SqlDbType.NVarChar, 6).Value = data.gender;
                        if (string.IsNullOrEmpty(data.mail))
                            cmd.Parameters.AddWithValue("@mail", DBNull.Value);
                        else
                            cmd.Parameters.Add("@mail", SqlDbType.NVarChar, 50).Value = data.mail;
                        cmd.Parameters.Add("@mobile_no_1", SqlDbType.NVarChar, 13).Value = data.mobile_no_1;
                        if (string.IsNullOrEmpty(data.mobile_no_2))
                            cmd.Parameters.AddWithValue("@mobile_no_2", DBNull.Value);
                        else
                            cmd.Parameters.Add("@mobile_no_2", SqlDbType.NVarChar, 13).Value = data.mobile_no_2;
                        if (string.IsNullOrEmpty(data.toeic_score))
                            cmd.Parameters.AddWithValue("@toeic_score", DBNull.Value);
                        else
                            cmd.Parameters.Add("@toeic_score", SqlDbType.NVarChar, 6).Value = data.toeic_score;
                        if (string.IsNullOrEmpty(data.start_working))
                            cmd.Parameters.AddWithValue("@start_working", DBNull.Value);
                        else
                            cmd.Parameters.Add("@start_working", SqlDbType.NVarChar, 15).Value = data.start_working;
                        if (string.IsNullOrEmpty(data.employ_status))
                            cmd.Parameters.AddWithValue("@employ_status", DBNull.Value);
                        else
                            cmd.Parameters.Add("@employ_status", SqlDbType.NVarChar, 10).Value = data.employ_status;
                        if (string.IsNullOrEmpty(data.cand_remarks))
                            cmd.Parameters.AddWithValue("@cand_remarks", DBNull.Value);
                        else
                            cmd.Parameters.Add("@cand_remarks", SqlDbType.NVarChar, 512).Value = data.cand_remarks;
                        if (string.IsNullOrEmpty(data.id_card))
                            cmd.Parameters.AddWithValue("@id_card", DBNull.Value);
                        else
                            cmd.Parameters.Add("@id_card", SqlDbType.NVarChar, 13).Value = data.id_card;
                        cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                        cmd.Parameters.Add("@create_time", SqlDbType.DateTime).Value = dtformat;
                        cmd.ExecuteNonQuery();

                        cand_id = data.cand_id;
                        List<Education> arr_edu = new List<Education>();
                        sql = $"select item_id from education where cand_id=@cand_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                        cmd.Transaction = transaction;
                        dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            arr_edu.Add(new Education() {
                                item_id = (int)dr["item_id"],
                            });
                        }
                        dr.Close();
                        if (arr_edu.Count > 0)
                        {
                            //delete all
                            if (data.educations == null)
                            {
                                sql = $"delete education where cand_id=@cand_id";
                                cmd = new SqlCommand(sql, db.GetConn());
                                cmd.Transaction = transaction;
                                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //delete is not new data
                                var res_edu = arr_edu.Where(d => !data.educations.Any(r => r.item_id == d.item_id));
                                var list_res = res_edu.ToList();
                                if (list_res.Count > 0)
                                {
                                    sql = "delete education where item_id in ({0})";
                                    string[] paramArr = list_res.Select((x, n) => "@list_in" + n).ToArray();
                                    cmd = new SqlCommand(sql, db.GetConn());
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = string.Format(sql, string.Join(",", paramArr));
                                    for (int i = 0; i < list_res.Count; i++)
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@list_in" + i, list_res[i].item_id));
                                    }
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        Management.InsertLog($"Candidate", $"Update Candidate cand_id={data.cand_id}");
                        break;
                }
                if (data.educations != null)
                {

                    for (var i = 0; i < data.educations.Count; i++)
                    {
                        if (data.educations[i].item_id == 0)
                        {
                            sql = $"insert into education (cand_id,degree,inst_name,faculty,major,grade,from_date,to_date,create_by,create_time)" +
                  $" values (@cand_id,@degree,@inst_name,@faculty,@major,@grade,@from_date,@to_date,@create_by,@create_time)";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                            cmd.Parameters.Add("@degree", SqlDbType.NVarChar, 55).Value = data.educations[i].degree;
                            cmd.Parameters.Add("@inst_name", SqlDbType.NVarChar, 175).Value = data.educations[i].inst_name;
                            cmd.Parameters.Add("@faculty", SqlDbType.NVarChar, 75).Value = data.educations[i].faculty;
                            cmd.Parameters.Add("@major", SqlDbType.NVarChar, 75).Value = data.educations[i].major;
                            if (string.IsNullOrEmpty(data.educations[i].grade))
                                cmd.Parameters.AddWithValue("@grade", DBNull.Value);
                            else
                                cmd.Parameters.Add("@grade", SqlDbType.VarChar, 10).Value = data.educations[i].grade;
                            if (string.IsNullOrEmpty(data.educations[i].from_date))
                                cmd.Parameters.AddWithValue("@from_date", DBNull.Value);
                            else
                                cmd.Parameters.Add("@from_date", SqlDbType.Date).Value = data.educations[i].from_date;
                            if (string.IsNullOrEmpty(data.educations[i].to_date))
                                cmd.Parameters.AddWithValue("@to_date", DBNull.Value);
                            else
                                cmd.Parameters.Add("@to_date", SqlDbType.Date).Value = data.educations[i].to_date;

                            cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                            cmd.Parameters.Add("@create_time", SqlDbType.Date).Value = dtformat;
                            cmd.ExecuteNonQuery();
                        }
                    }
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
            return Json(cand_id, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult InsertApplyRecordByCandID(ApplyRecord data)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            string sql = string.Empty;
            db.OpenConn();
            DateTime dt = DateTime.Now;
            string dtformat = dt.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string cv_no = string.Empty;
            try
            {
                sql = $"insert into apply_record (cand_id,cv_no,apply_position,exp_salary,industry,keyword,remarks,apply_date,create_by,create_time) " +
                  $"values(@cand_id,@cv_no,@apply_position,@exp_salary,@industry,@keyword,@remarks,@apply_date,@create_by,@create_time)";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = data.cand_id;
                cmd.Parameters.Add("@cv_no", SqlDbType.NVarChar, 20).Value = data.cv_no;
                cmd.Parameters.Add("@apply_position", SqlDbType.NVarChar, 100).Value = data.apply_position;
                if (string.IsNullOrEmpty(data.exp_salary))
                    cmd.Parameters.AddWithValue("@exp_salary", DBNull.Value);
                else
                    cmd.Parameters.Add("@exp_salary", SqlDbType.Decimal).Value = data.exp_salary;
                if (string.IsNullOrEmpty(data.industry))
                    cmd.Parameters.AddWithValue("@industry", DBNull.Value);
                else
                    cmd.Parameters.Add("@industry", SqlDbType.NVarChar, 75).Value = data.industry;
                if (string.IsNullOrEmpty(data.keyword))
                    cmd.Parameters.AddWithValue("@keyword", DBNull.Value);
                else
                    cmd.Parameters.Add("@keyword", SqlDbType.NVarChar, 512).Value = data.keyword;
                if (string.IsNullOrEmpty(data.remark))
                    cmd.Parameters.AddWithValue("@remarks", DBNull.Value);
                else
                    cmd.Parameters.Add("@remarks", SqlDbType.NVarChar, 256).Value = data.remark;

                cmd.Parameters.Add("@apply_date", SqlDbType.Date).Value = data.apply_date;
                cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                cmd.Parameters.Add("@create_time", SqlDbType.Date).Value = dtformat;
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Application Record", $"Create Application Record position={data.apply_position} cand_id={  data.cand_id}");
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(cv_no, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetCandDetail(int cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            string stp = string.Empty;
            List<Candidate> prs = new List<Candidate>();
            List<Education> edu = new List<Education>();
            List<ApplyRecord> apy = new List<ApplyRecord>();
            try
            {
                db.OpenConn();
                stp = "GetPersonalInformation";
                cmd = new SqlCommand(stp, db.GetConn());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    prs.Add(new Candidate() {
                        cand_id = (int)dr["cand_id"],
                        first_name = $"{dr["first_name"]}",
                        last_name = $"{dr["last_name"]}",
                        gender = $"{dr["gender"]}",
                        remarks = (dr["cand_remarks"] == DBNull.Value) ? string.Empty : $"{dr["cand_remarks"]}",
                        mobile = $"{dr["mobile_no_1"]}",
                        mail = (dr["mail"] == DBNull.Value) ? string.Empty : $"{dr["mail"]}",
                        birth_day = (dr["birth_day"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["birth_day"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        mobile_no_2 = (dr["mobile_no_2"] == DBNull.Value) ? string.Empty : $"{dr["mobile_no_2"]}",
                        toeic_score = (dr["toeic_score"] == DBNull.Value) ? string.Empty : $"{dr["toeic_score"]}",
                        start_working = (dr["start_working"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["start_working"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        employ_status = $"{dr["employ_status"]}",
                        id_card = $"{dr["id_card"]}"
                    });
                }
                dr.Close();
                stp = "GetCandidateEducation";
                cmd = new SqlCommand(stp, db.GetConn());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    edu.Add(new Education() {
                        item_id = (int)dr["item_id"],
                        degree = $"{dr["degree"]}",
                        inst_name = $"{dr["inst_name"]}",
                        faculty = $"{dr["faculty"]}",
                        major = $"{dr["major"]}",
                        grade = (dr["grade"] == DBNull.Value) ? string.Empty : $"{dr["grade"]}",
                        from_date = (dr["from_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["from_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        to_date = (dr["to_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["to_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                    });
                }
                dr.Close();
                stp = "GetCandidateCV";
                cmd = new SqlCommand(stp, db.GetConn());
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    apy.Add(new ApplyRecord()
                    {
                        item_id = (int)dr["item_id"],
                        cv_no = $"{dr["cv_no"]}",
                        cv_id = (int)dr["cv_id"],
                        //cv_file = (dr["cv_file"] == DBNull.Value) ? string.Empty : "not null",
                        apply_position = $"{dr["apply_position"]}",
                        exp_salary = $"{dr["exp_salary"]}",
                        industry = $"{dr["industry"]}",
                        keyword = $"{dr["keyword"]}",
                        remark = $"{dr["remarks"]}",
                        apply_date = Convert.ToDateTime($"{dr["apply_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                    });
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
            return Json(new { personalInfo = prs, education = edu, apply = apy }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult SaveFileSream()
        {
            string cv_no = string.Empty;
            //cv_no = Request["cv_no"];
            var full_name = Request["uploadername"];
            HttpFileCollectionBase files = Request.Files;
            var fullPath = string.Empty;

            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            int last_no = 0;
            string str_no = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select top 1 cv_no from cv_details order by  RIGHT(cv_no, 5) desc";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    string[] str = $"{dr["cv_no"]}".Split('-');
                    foreach (var n in str)
                    {
                        str_no = n;
                    }
                    last_no = Convert.ToInt32(str_no);
                }
                last_no++;
                dr.Close();
                DateTime dateTime = DateTime.UtcNow.Date;
                cv_no = $"CV-{DateTime.Now.Year.ToString()}-{DateTime.Now.ToString("MM")}-{String.Format("{0:D5}", last_no)}";


                if (files.Count != 0)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        //var file = files[i];
                        //Stream file_stream = file.InputStream;
                        //byte[] filestream = ConvertFileSream.ConvertFileStream(file_stream);
                        //ConvertFileSream.SaveCvFile(filestream, cv_no);
                        
                        //upload file
                        HttpPostedFileBase file = files[i];
                        string file_name = files[i].FileName;
                        float file_size = files[i].ContentLength;
                        string doc_extension = Path.GetExtension(file_name);
                        string doc_name = cv_no + "_"+full_name+ doc_extension;
                        fullPath = Path.Combine(Server.MapPath("~/Uploads/CV/"), doc_name);
                        file.SaveAs(fullPath);

                        //update database
                        ConvertFileSream.SaveCvFile(doc_name, doc_extension, cv_no);
                    }
                }
                else
                {
                    throw new Exception("");
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
            return Json(cv_no, JsonRequestBehavior.AllowGet);
        }
        public ActionResult OpenCvFile(string cv_id)
        {
            int int_cv_id = Management.EncodeToNumber(cv_id);
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            //byte[] cv_file = new byte[0];
            string path = string.Empty;
            string file_name = string.Empty;
            string file_extension = string.Empty;
            string position = string.Empty;
            string type = "application/pdf";
            byte[] res_file = new byte[0];
            try
            {
                db.OpenConn();
                string sql = "select doc_name,doc_extension,apply_position from cv_details c left join apply_record a on c.cv_no=a.cv_no where cv_id=@cv_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@cv_id", SqlDbType.Int).Value = int_cv_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    //if (dr["cv_file"] != DBNull.Value)
                    //{
                    //    cv_file = (byte[])dr["cv_file"];

                    //}
                    file_name = $"{dr["doc_name"]}";
                    file_extension = $"{dr["doc_extension"]}";
                    position = (dr["apply_position"] == DBNull.Value) ? string.Empty : $"{dr["apply_position"]}";
                }
                dr.Close();
                path = Path.Combine(Server.MapPath("~/Uploads/CV/"), file_name);
                res_file = System.IO.File.ReadAllBytes(path);
                type = ReadFileType.GetTypeFile(file_extension);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally {
                db.CloseConn();
            }
            if (type == "application/pdf")
            {
                Response.AddHeader("content-disposition", "inline;filename=" +  position + "_"+ file_name.Substring(16) + ".pdf");
                return  File(res_file, "application/pdf");
            }
            else
            {
                return File(res_file, type, $"{position}_{file_name.Substring(16)}");

            }
           
        }

        public JsonResult UpdateApplyByItemID(ApplyRecord data)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;

            try
            {
                db.OpenConn();
                string sql = $"update apply_record set apply_position=@apply_position, exp_salary=@exp_salary, apply_date=@apply_date," +
                    $" industry=@industry,remarks=@remarks,keyword=@keyword where item_id=@item_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@item_id", SqlDbType.Int).Value = data.item_id;
                cmd.Parameters.Add("@apply_position", SqlDbType.NVarChar, 100).Value = data.apply_position;
                if (string.IsNullOrEmpty(data.exp_salary))
                    cmd.Parameters.AddWithValue("@exp_salary", DBNull.Value);
                else
                    cmd.Parameters.Add("@exp_salary", SqlDbType.Decimal).Value = data.exp_salary;

                cmd.Parameters.Add("@apply_date", SqlDbType.Date).Value = data.apply_date;
                if (string.IsNullOrEmpty(data.industry))
                    cmd.Parameters.AddWithValue("@industry", DBNull.Value);
                else
                    cmd.Parameters.Add("@industry", SqlDbType.NVarChar, 75).Value = data.industry;
                if (string.IsNullOrEmpty(data.remark))
                    cmd.Parameters.AddWithValue("@remarks", DBNull.Value);
                else
                    cmd.Parameters.Add("@remarks", SqlDbType.NVarChar, 256).Value = data.remark;
                if (string.IsNullOrEmpty(data.keyword))
                    cmd.Parameters.AddWithValue("@keyword", DBNull.Value);
                else
                    cmd.Parameters.Add("@keyword", SqlDbType.NVarChar, 512).Value = data.keyword;
                cmd.ExecuteNonQuery();
                Management.InsertLog($"Application Record", $"Update Application Record position={data.apply_position} item_id={  data.item_id}");
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        [HttpPost]
        public JsonResult DelApplyByItemID(int item_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            SqlTransaction transaction;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            string sql = string.Empty;
            string cv_no = string.Empty;
            string doc_name = string.Empty;
            try
            {
                sql = $"select a.cv_no,c.doc_name from apply_record a inner join cv_details c on c.cv_no=a.cv_no where item_id = @item_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@item_id", SqlDbType.Int).Value = item_id;
                cmd.Transaction = transaction;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    cv_no = $"{dr["cv_no"]}";
                    doc_name = (dr["doc_name"] == DBNull.Value) ? string.Empty : $"{dr["doc_name"]}";
                    dr.Close();

                    if (!string.IsNullOrEmpty(doc_name))
                    {
                        string file_path = Path.Combine(Server.MapPath("~/Uploads/CV"), doc_name);

                        if (System.IO.File.Exists(file_path))
                        {
                            System.IO.File.Delete(file_path);
                        }
                    }
                    sql = $"delete apply_record where item_id=@item_id";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Transaction = transaction;
                    cmd.Parameters.Add("@item_id", SqlDbType.Int).Value = item_id;
                    cmd.ExecuteNonQuery();

                    sql = $"delete cv_details where cv_no=@cv_no";
                    cmd = new SqlCommand(sql, db.GetConn());
                    cmd.Parameters.Add("@cv_no", SqlDbType.NVarChar, 20).Value = cv_no;
                    cmd.Transaction = transaction;
                    cmd.ExecuteNonQuery();

                }

                transaction.Commit();
                Management.InsertLog($"Appliction Record", $"Delete Appliction Record cv_no={cv_no} item_id={ item_id}");
            }
            catch (Exception)
            {
                transaction.Rollback();
                throw;
            }
            finally
            {
                db.CloseConn();
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        [HttpPost]
        public JsonResult DelCandByCandID(int cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataAdapter da;
            SqlTransaction transaction;
            SqlDataReader dr;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            string sql = string.Empty;
            string cv_no = string.Empty;
            try
            {
                //select cv_no from apply 
                sql = $"select cv_no from apply_record where cand_id = @cand_id";
                da = new SqlDataAdapter(sql, db.GetConn());
                da.SelectCommand.Transaction = transaction;
                da.SelectCommand.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                DataTable table = new DataTable();
                da.Fill(table);
                foreach (DataRow row in table.Rows)
                {
                    cv_no = $"{ row["cv_no"]}";
                }
                //delete apply by cand_id
                sql = $"delete apply_record where cand_id=@cand_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                cmd.ExecuteNonQuery();

                sql = $"select doc_name from cv_details where cv_no=@cv_no";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@cv_no", SqlDbType.NVarChar, 20).Value = cv_no;
                cmd.Transaction = transaction;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    string doc_name = (dr["doc_name"] == DBNull.Value) ? string.Empty : $"{dr["doc_name"]}";
                    if (!string.IsNullOrEmpty(doc_name))
                    {
                        string file_path = Path.Combine(Server.MapPath("~/Uploads/CV"), doc_name);

                        if (System.IO.File.Exists(file_path))
                        {
                            System.IO.File.Delete(file_path);
                        }
                    }
                }
                dr.Close();

                //delete cv_det by cv_no
                sql = $"delete cv_details where cv_no=@cv_no";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@cv_no", SqlDbType.NVarChar, 20).Value = cv_no;
                cmd.ExecuteNonQuery();

                //delete education by cand_id
                sql = $"delete education where cand_id = @cand_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                cmd.ExecuteNonQuery();

                //delete cand by cand_id
                sql = $"delete candidate where cand_id = @cand_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                cmd.ExecuteNonQuery();

                transaction.Commit();
                Management.InsertLog($"Candidate", $"Delete Candidate cand_id={cand_id}");
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
        [HttpGet]
        public JsonResult CheckMobile(string mobile, int cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            bool checkmobile = false;
            try
            {
                db.OpenConn();
                string sql = $"select mobile_no_1 from candidate where mobile_no_1=@mobile";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@mobile", SqlDbType.NVarChar, 13).Value = mobile;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    checkmobile = true;
                    dr.Close();
                    if (cand_id != 0)
                    {
                        sql = $"select mobile_no_1 from candidate where cand_id=@cand_id and mobile_no_1=@mobile";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                        cmd.Parameters.Add("@mobile", SqlDbType.NVarChar, 13).Value = mobile;
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            checkmobile = false;
                        }
                        else {
                            checkmobile = true;
                        }
                    }
                }


            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return Json(checkmobile, JsonRequestBehavior.AllowGet);
        }
        //public bool ConvertCV()
        //{
        //    bool success = false;
        //    Connect_DB db = new Connect_DB();
        //    SqlDataReader dr;
        //    SqlCommand cmd;
        //    SqlDataAdapter da;
        //    SqlTransaction transaction;
        //    db.OpenConn();
        //    transaction = db.GetConn().BeginTransaction();
        //    string cvpath = string.Empty;
        //    var cont = new RequisitionController();
        //    try
        //    {
        //        string sql = $"select * from cv_details where srv_doc_name is not null and cv_file is null";
        //        da = new SqlDataAdapter(sql, db.GetConn());
        //        DataTable table = new DataTable();
        //        da.SelectCommand.Transaction = transaction;
        //        da.Fill(table);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            byte[] fileStream;
        //            cvpath = cont.DeCodeBase64($"{row["srv_doc_path"]}");
        //            var client = new WebClient();
        //            var content = client.DownloadData(cvpath);
        //            var stream = new MemoryStream(content);
        //            byte[] file;
        //            file = new byte[stream.Length];
        //            stream.Read(file, 0, file.Length);
        //            fileStream = file;

        //            sql = $"update cv_details set cv_file = @cv_file where cv_id = @cv_id";
        //            cmd = new SqlCommand(sql, db.GetConn());
        //            cmd.Transaction = transaction;
        //            cmd.Parameters.Add("@cv_id", SqlDbType.Int).Value = (int)row["cv_id"];
        //            cmd.Parameters.Add("@cv_file", SqlDbType.VarBinary).Value = fileStream;
        //            cmd.ExecuteNonQuery();
        //        }
        //        success = true;
        //        transaction.Commit();
        //    }
        //    catch (Exception ex)
        //    {
        //        transaction.Rollback();
        //        throw ex;
        //    }
        //    finally
        //    {
        //        db.CloseConn();
        //    }
        //    return success;
        //}
        public JsonResult CheckCandIdFK(string cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            bool res = false;
            try
            {
                db.OpenConn();
                string sql = $"select req_cand_id from req_cand r left join req_success s on s.cand_id = r.cand_id where r.cand_id =@cand_id ";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    res = true;
                }
                dr.Close();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return Json(res, JsonRequestBehavior.AllowGet);

        }
        public ActionResult OpenPPS(int cand_id)
        {
            byte[] pps = new byte[0];

            List<FilePPS> filePPs = new List<FilePPS>();
            MemoryStream stream = new MemoryStream();
            string position = string.Empty;
            string rev_no = string.Empty;
            try
            {
                string obj = new JavaScriptSerializer().Serialize(CheckPpsID(cand_id));
                dynamic data = JObject.Parse(obj);
                if (data["Data"] != 0)
                {
                    filePPs = CreateWordByOpenXML.GetDetailProposal(Convert.ToInt32(data["Data"]));
                    stream = filePPs[0].stream;
                    position = filePPs[0].position.ToLower();
                    rev_no = filePPs[0].rev_no;
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"proposal-{position}-rev{rev_no}.docx");

        }
        public JsonResult CheckPpsID(int cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            int proposal_id = 0;
            try
            {
                db.OpenConn();
                string sql = $"select top 1 proposal_id,rev_no from proposal_list p " +
                    $"inner join req_success s on s.succ_id = p.succ_id where cand_id = @cand_id and rev_no is not null order by rev_no desc";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    proposal_id = (int)dr["proposal_id"];
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
            return Json(proposal_id, JsonRequestBehavior.AllowGet);

        }
        public JsonResult AdvSearch(string emp_status, string month, string remark, string gender, string age, string toeic
            , string degree, string faculty, string major, string industry, string keyword)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            List<Cand> res = new List<Cand>();
            int no = 1;
            try
            {
                int m = Convert.ToInt32(DateTime.Now.ToString("MM"));
                int l_m_1 = Convert.ToInt32(DateTime.Now.AddMonths(-1).ToString("MM"));
                int l_m_2 = Convert.ToInt32(DateTime.Now.AddMonths(-2).ToString("MM"));

                int y = Convert.ToInt32(DateTime.Now.ToString("yyyy"));
                int l_y_1 = Convert.ToInt32(DateTime.Now.AddMonths(-1).ToString("yyyy"));
                int l_y_2 = Convert.ToInt32(DateTime.Now.AddMonths(-2).ToString("yyyy"));
                string year = $"like '%'";
                if (month.Equals("month"))
                {
                    month = $"like {m}";
                    year = $"like {y}";
                }
                else if (month.Equals("3m"))
                {
                    month = $"in ({m},{l_m_1},{l_m_2})";
                    year = $"in ({y},{l_y_1},{l_y_2})";
                }
                else if (month.Equals("year"))
                {
                    month = $"like %";
                    year = y.ToString();
                }
                else
                {
                    month = $"like '%'";
                }
                string[] a = age.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                if (a.Length == 1)
                {
                    age = $"DATEDIFF(yy, c.birth_day, getdate()) <= '{a[0]}'";
                }
                else if (a.Length == 2)
                {
                    age = $"DATEDIFF(yy, c.birth_day, getdate()) between {a[0]} and {a[1]}";
                }
                else
                {
                    age = $"(DATEDIFF(yy, c.birth_day, getdate()) like '%' or DATEDIFF(yy, c.birth_day, getdate()) is null )";
                }

                remark = (string.IsNullOrEmpty(remark)) ? $"(remarks like '%' or remarks is null )" : $"remarks like '%{remark}%'";
                toeic = (string.IsNullOrEmpty(toeic)) ? $"(toeic_score like '%' or toeic_score is null )" : $"toeic_score >= {toeic}";
                degree = (string.IsNullOrEmpty(degree)) ? $"(degree like '%' or degree is null )" : $"degree like '%{degree}%'";
                faculty = (string.IsNullOrEmpty(faculty)) ? $"(faculty like '%' or faculty is null )" : $"faculty like '%{faculty}%'";
                major = (string.IsNullOrEmpty(major)) ? $"(major like '%' or major is null )" : $"major like '%{major}%'";
                industry = (string.IsNullOrEmpty(industry)) ? $"(industry like '%' or industry is null )" : $"industry like '%{industry}%'";
                keyword = (string.IsNullOrEmpty(keyword)) ? $"(keyword like '%' or keyword is null )" : $"keyword like '%{keyword}%'";
                emp_status = (string.IsNullOrEmpty(emp_status)) ? $"%" : emp_status;
                gender = (string.IsNullOrEmpty(gender)) ? $"%" : gender;
                db.OpenConn();
                sql = $"select c.cand_id,c.first_name,c.last_name,c.gender, DATEDIFF(yy, c.birth_day, getdate()) as age,a.apply_position,a.apply_date, " +
                $"a.remarks,c.mobile_no_1,c.mail,c.create_by,c.create_time,cv.cv_id,toeic_score,e.degree,e.faculty,e.major,a.industry,a.keyword " +
                $"from candidate c " +
                $"outer apply(select top 1 apply_position, remarks, cv_no,create_by, create_time,apply_date,industry,keyword from apply_record as r where r.cand_id = c.cand_id order by create_time desc ) as a " +
                $"left join cv_details cv on cv.cv_no = a.cv_no " +
                $"outer apply(select top 1 degree, faculty, major from education as d where d.cand_id = c.cand_id order by create_time desc ) as e " +
                $"where c.employ_status like @emp_status  " +
                $"and DATEPART(Month, a.apply_date) {month} " +
                $"and DATEPART(YEAR, a.apply_date) {year} " +
                $"and {remark} " +
                $"and gender like @gender " +
                $"and {age} " +
                $"and {toeic} " +
                $"and {degree} " +
                $"and {faculty} " +
                $"and {major} " +
                $"and {industry} " +
                $"and {keyword} " +
                $"order by a.apply_date desc";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@emp_status", SqlDbType.NVarChar).Value = emp_status;
                cmd.Parameters.Add("@year", SqlDbType.NVarChar).Value = year;
                cmd.Parameters.Add("@gender", SqlDbType.NVarChar).Value = gender;
                //cmd.Parameters.Add("@age", SqlDbType.NVarChar).Value = age;
                //cmd.Parameters.Add("@month", SqlDbType.NVarChar).Value = month;
                //cmd.Parameters.Add("@remark", SqlDbType.NVarChar).Value = "%" + remark + "%";
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Cand()
                    {
                        no = no++,
                        cand_id = (int)dr["cand_id"],
                        full_name = $"{dr["first_name"]} {dr["last_name"]}",
                        gender = $"{dr["gender"]}",
                        age = (dr["age"] == DBNull.Value) ? string.Empty : $"{dr["age"]}",
                        last_apply_position = (dr["apply_position"] == DBNull.Value) ? "No Record" : $"{dr["apply_position"]}",
                        remarks = (dr["remarks"] == DBNull.Value) ? string.Empty : $"{dr["remarks"]}",
                        mobile = $"{dr["mobile_no_1"]}",
                        mail = (dr["mail"] == DBNull.Value) ? string.Empty : $"{dr["mail"]}",
                        create_date = string.IsNullOrEmpty($"{dr["create_time"]}") ? string.Empty : Convert.ToDateTime($"{dr["create_time"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        create_by = (dr["create_by"] == DBNull.Value) ? string.Empty : $"{dr["create_by"]}",
                        cv_id = (dr["cv_id"] == DBNull.Value) ? 0 : (int)dr["cv_id"],
                        //cv_file = (dr["cv_file"] == DBNull.Value) ? false : true,
                    });
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
            return Json(res, JsonRequestBehavior.AllowGet);
        }
        public ActionResult SavePropoSal()
        {
            var succ_id = Request["succ_id"];
            var position = Request["position"];
            HttpFileCollectionBase files = Request.Files;
            int rev = 0;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            var fullPath = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select top 1 rev_no from proposal_list where succ_id=@succ_id  order by rev_no desc";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = succ_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    rev = (int)dr["rev_no"];
                    rev++;
                }
                dr.Close();
                db.CloseConn();
                string fname = "proposal-official-" + position.ToLower() + "-rev" + rev;
                if (files.Count != 0)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        //var file = files[i];
                        //Stream file_stream = file.InputStream;
                        //byte[] filestream = ConvertFileSream.ConvertFileStream(file_stream);
                        //ConvertFileSream.SavePpsFile(filestream, Convert.ToInt32(succ_id), file_name, rev);
                        HttpPostedFileBase file = files[i];
                        string file_name = files[i].FileName;
                        float file_size = files[i].ContentLength;
                        string doc_extension = Path.GetExtension(file_name);
                        string doc_name = fname.Replace(@"/", "-") + doc_extension;
                        fullPath = Path.Combine(Server.MapPath("~/Uploads/Proposal/"), doc_name);
                        file.SaveAs(fullPath);

                        //update database
                        ConvertFileSream.SavePpsFile(doc_name, doc_extension, Convert.ToInt32(succ_id), rev);
                    }
                }
                else
                {
                    throw new Exception("");
                }
            }
            catch (Exception)
            {

                throw;
            }
            return Json(JsonRequestBehavior.AllowGet);
        }
        public ActionResult OpenPpsFile(string proposal_id)
        {
            int int_proposal_id = Management.EncodeToNumber(proposal_id);
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string path = string.Empty;
            string file_name = string.Empty;
            string file_extension = string.Empty;
            string type = "application/pdf";
            byte[] res_file = new byte[0];
            try
            {
                db.OpenConn();
                string sql = "select file_name,file_extension from proposal_list where proposal_id=@proposal_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@proposal_id", SqlDbType.Int).Value = int_proposal_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    file_name = $"{dr["file_name"]}";
                    file_extension = $"{dr["file_extension"]}";
                }
                dr.Close();
                path = Path.Combine(Server.MapPath("~/Uploads/Proposal/"), file_name);
                res_file = System.IO.File.ReadAllBytes(path);
                type = ReadFileType.GetTypeFile(file_extension);

            }
            catch (Exception)
            {

                throw;
            }
            finally {
                db.CloseConn();
            }
            return File(res_file, type);
        }
        [HttpPost]
        public JsonResult DeletePpsOfficial(int pps_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select file_name from proposal_list where proposal_id = @proposal_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@proposal_id", SqlDbType.Int).Value = pps_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                   string doc_name = $"{dr["file_name"]}";
                    dr.Close();
                    if (!string.IsNullOrEmpty(doc_name))
                    {
                        string file_path = Path.Combine(Server.MapPath("~/Uploads/Proposal"), doc_name);

                        if (System.IO.File.Exists(file_path))
                        {
                            System.IO.File.Delete(file_path);
                        }
                    }
                }
                sql = $"delete proposal_list where proposal_id=@proposal_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@proposal_id", SqlDbType.Int).Value = pps_id;
                cmd.ExecuteNonQuery();
                db.CloseConn();

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetCommByCandID(int cand_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string res = string.Empty;
            try
            {
                db.OpenConn();
                string sql = $"select comment from req_success where cand_id=@cand_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    res = $"{dr["comment"]}";
                }
                dr.Close();
                db.CloseConn();

            }
            catch (Exception ex)
            {

                throw ex; 
            }
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult InsertCommByCandID(int cand_id, string comment)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;

            try
            {
                db.OpenConn();
                string sql = $"update req_success set comment=@comment where cand_id=@cand_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@cand_id", SqlDbType.Int).Value = cand_id;
                cmd.Parameters.Add("@comment", SqlDbType.NVarChar,500).Value = comment;
                cmd.ExecuteNonQuery();
                db.CloseConn();

            }
            catch (Exception ex)
            {

                throw ex;
            }
            return Json(JsonRequestBehavior.DenyGet);
        }

        //public string ConvertByteArrToFileCV()
        //{
        //    Connect_DB db = new Connect_DB();
        //    SqlDataReader dr;
        //    SqlDataAdapter da;
        //    SqlCommand cmd;
        //    SqlTransaction transaction;
        //    string res = "failed";
        //    byte[] cv_file = new byte[0];
        //    db.OpenConn();
        //    transaction = db.GetConn().BeginTransaction();
        //    string cv_no_err = string.Empty;
        //    Regex reg = new Regex("[*'\",_&#^@:|'/]");
        //    try
        //    {
        //        string sql = "select cv_id,first_name,last_name,cv_file,c.cv_no from cv_details c inner join apply_record a on a.cv_no=c.cv_no inner join candidate d on d.cand_id=a.cand_id where  c.cv_file is not null";
        //        da = new SqlDataAdapter(sql, db.GetConn());
        //        da.SelectCommand.CommandTimeout = 300;
        //        //da.SelectCommand.Parameters.Add("@cv_id", SqlDbType.Int).Value = 185;
        //        da.SelectCommand.Transaction = transaction;
        //        DataTable table = new DataTable();
        //        da.Fill(table);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            cv_no_err = $"{row["cv_no"]}";
        //            if (row["cv_file"] != DBNull.Value)
        //            {
        //                string first_name = reg.Replace($"{row["first_name"]}", string.Empty);
        //                string last_name = reg.Replace($"{row["last_name"]}", string.Empty);
        //                string fname = $"{row["cv_no"]}_{last_name}_{last_name}" + ".pdf";
        //                var fullPath = Path.Combine(Server.MapPath("~/Uploads/CV/"), fname);
        //                cv_file = (byte[])row["cv_file"];
        //                System.IO.File.WriteAllBytes(fullPath, cv_file);

        //                sql = $"update cv_details set doc_name=@doc_name,doc_extension=@doc_extension where cv_id=@cv_id";
        //                cmd = new SqlCommand(sql, db.GetConn());
        //                cmd.Transaction = transaction;
        //                cmd.Parameters.Add("@cv_id", SqlDbType.Int).Value = (int)row["cv_id"];
        //                cmd.Parameters.Add("@doc_name", SqlDbType.NVarChar).Value = fname;
        //                cmd.Parameters.Add("@doc_extension", SqlDbType.NVarChar).Value = ".pdf";
        //                cmd.ExecuteNonQuery();

        //                res = "success";
        //            }
        //        }
        //        transaction.Commit();
        //    }
        //    catch (Exception ex)
        //    {
        //        System.IO.File.WriteAllText(@"D:\msvs_project2021\Janejira\Doc\MSPRO\ConvertByteArrToFileCV.txt", "message : " + ex.Message + ", cv_no : " + cv_no_err);
        //        transaction.Rollback();
        //        throw ex;
        //    }
        //    finally
        //    {
        //        db.CloseConn();
        //    }
        //    return res;
        //}
        //public string ConvertByteArrToFileReq()
        //{
        //    Connect_DB db = new Connect_DB();
        //    SqlDataReader dr;
        //    SqlDataAdapter da;
        //    SqlCommand cmd;
        //    SqlTransaction transaction;
        //    string res = "failed";
        //    byte[] cv_file = new byte[0];
        //    db.OpenConn();
        //    transaction = db.GetConn().BeginTransaction();
        //    try
        //    {

        //        string sql = "select req_id,req_no,req_file from requisition where req_file is not null";
        //        da = new SqlDataAdapter(sql, db.GetConn());
        //        da.SelectCommand.Transaction = transaction;
        //        da.SelectCommand.CommandTimeout = 300;
        //        //da.SelectCommand.Parameters.Add("@req_id", SqlDbType.Int).Value = 92;
        //        DataTable table = new DataTable();
        //        da.Fill(table);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            if (row["req_file"] != DBNull.Value)
        //            {
        //                string fname = "REQ" + "-" + $"{row["req_no"]}".Replace(@"/", "-") + ".pdf";
        //                var fullPath = Path.Combine(Server.MapPath("~/Uploads/REQ/"), fname);
        //                cv_file = (byte[])row["req_file"];
        //                System.IO.File.WriteAllBytes(fullPath, cv_file);
                       
        //                sql = $"update requisition set file_extension=@file_extension, file_name=@file_name where req_id=@req_id";
        //                cmd = new SqlCommand(sql, db.GetConn());
        //                cmd.Transaction = transaction;
        //                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = (int)row["req_id"];
        //                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = fname;
        //                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = ".pdf";
        //                cmd.ExecuteNonQuery();

        //                res = "success";
        //            }
        //        }
        //        transaction.Commit();
        //    }
        //    catch (Exception ex)
        //    {
        //        transaction.Rollback();
        //        System.IO.File.WriteAllText(@"D:\msvs_project2021\Janejira\Doc\MSPRO\ConvertByteArrToFileReq.txt", "message : " + ex.Message );
        //        throw ex;
        //    }
        //    finally
        //    {
        //        db.CloseConn();
        //    }
        //    return res;
        //}
        //public string ConvertByteArrToFileProp()
        //{
        //    Connect_DB db = new Connect_DB();
        //    SqlDataReader dr;
        //    SqlDataAdapter da;
        //    SqlCommand cmd;
        //    SqlTransaction transaction;
        //    string res = "failed";
        //    byte[] cv_file = new byte[0];
        //    db.OpenConn();
        //    transaction = db.GetConn().BeginTransaction();
        //    try
        //    {

        //        string sql = "select proposal_id,file_name,pps_official_file,rev_no from proposal_list where pps_official_file is not null";
        //        da = new SqlDataAdapter(sql, db.GetConn());
        //        da.SelectCommand.Transaction = transaction;
        //        da.SelectCommand.CommandTimeout = 300;
        //        //da.SelectCommand.Parameters.Add("@proposal_id", SqlDbType.Int).Value = 95;
        //        DataTable table = new DataTable();
        //        da.Fill(table);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            if (row["pps_official_file"] != DBNull.Value)
        //            {
        //                string fname = $"{row["file_name"]}".Replace(@"/", "-") + ".pdf";
        //                var fullPath = Path.Combine(Server.MapPath("~/Uploads/Proposal/"), fname);
        //                cv_file = (byte[])row["pps_official_file"];
        //                System.IO.File.WriteAllBytes(fullPath, cv_file);

        //                sql = $"update proposal_list set file_name=@file_name,file_extension=@file_extension where proposal_id=@proposal_id";
        //                cmd = new SqlCommand(sql, db.GetConn());
        //                cmd.Transaction = transaction;
        //                cmd.Parameters.Add("@proposal_id", SqlDbType.Int).Value = (int)row["proposal_id"];
        //                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = fname;
        //                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = ".pdf";
        //                cmd.ExecuteNonQuery();

        //                res = "success";
        //            }
        //        }
        //        transaction.Commit();
        //    }
        //    catch (Exception ex)
        //    {
        //        transaction.Rollback();
        //        throw ex;
        //    }
        //    finally
        //    {
        //        db.CloseConn();
        //    }
        //    return res;
        //}
        //public string ConvertByteArrToFileSrvAgr()
        //{
        //    Connect_DB db = new Connect_DB();
        //    SqlDataReader dr;
        //    SqlDataAdapter da;
        //    SqlCommand cmd;
        //    SqlTransaction transaction;
        //    string res = "failed";
        //    byte[] cv_file = new byte[0];
        //    db.OpenConn();
        //    transaction = db.GetConn().BeginTransaction();
        //    try
        //    {

        //        string sql = "select srv_id,srv_file,file_name from srv_agreement_list where srv_file is not null";
        //        da = new SqlDataAdapter(sql, db.GetConn());
        //        da.SelectCommand.Transaction = transaction;
        //        da.SelectCommand.CommandTimeout = 300;
        //        //da.SelectCommand.Parameters.Add("@srv_id", SqlDbType.Int).Value = 12;
        //        DataTable table = new DataTable();
        //        da.Fill(table);
        //        foreach (DataRow row in table.Rows)
        //        {
        //            if (row["srv_file"] != DBNull.Value)
        //            {
        //                string fname = $"{row["file_name"]}" + ".pdf";
        //                var fullPath = Path.Combine(Server.MapPath("~/Uploads/SrvAgr/"), fname);
        //                cv_file = (byte[])row["srv_file"];
        //                System.IO.File.WriteAllBytes(fullPath, cv_file);

        //                sql = $"update srv_agreement_list set file_name=@file_name,file_extension=@file_extension where srv_id=@srv_id";
        //                cmd = new SqlCommand(sql, db.GetConn());
        //                cmd.Transaction = transaction;
        //                cmd.Parameters.Add("@srv_id", SqlDbType.Int).Value = (int)row["srv_id"];
        //                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = fname;
        //                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = ".pdf";
        //                cmd.ExecuteNonQuery();

        //                res = "success";
        //            }
        //        }
        //        transaction.Commit();
        //    }
        //    catch (Exception ex)
        //    {
        //        transaction.Rollback();
        //        throw ex;
        //    }
        //    finally
        //    {
        //        db.CloseConn();
        //    }
        //    return res;
        //}
    }

}