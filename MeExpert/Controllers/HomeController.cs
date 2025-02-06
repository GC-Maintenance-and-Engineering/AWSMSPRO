using MeExpert.CommonProvide;
using MeExpert.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace MeExpert.Controllers
{
   
    public class HomeController : Controller
    {
        //[AuthenFilter]
        public ActionResult Index(string emp_code, string token)
        {
            return View();
        }
        public ActionResult Permission(string message)
        {
            ViewBag.Message = message;
            return View();
        }
        public JsonResult Get() 
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Dashboard> res = new List<Dashboard>();

            try
            {
                db.OpenConn();
                string sql = $" select 'requisition' as name ," +
                    $"count(*) as count " +
                    $"from requisition union all select 'candidate' ," +
                    $"count(*) from req_cand  where status='pass'  union all select 'client' ," +
                    $"count(*) from client";
                cmd = new SqlCommand(sql,db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Dashboard()
                    {
                        name = $"{dr["name"]}",
                        number = (int)dr["count"],
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
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetPieChart()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Dashboard> res = new List<Dashboard>();
            try
            {
                db.OpenConn();
                string sql = $"select req_status,COUNT(*) as count from requisition group by req_status order by req_status asc";
                cmd = new SqlCommand(sql,db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Dashboard() { 
                        name = $"{dr["req_status"]}",
                        number = (int)dr["count"],
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
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetBarChart()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Dur> res = new List<Dur>();
            try
            {
                db.OpenConn();
                string sql = $"select top 4 s_date,rs.create_date,req_no" +
                $" from requisition r" +
                $" outer apply (select top 1 create_date from req_success s where s.req_id = r.req_id order by s.create_date asc ) as rs" +
                $" where req_status = 'close' and s_date IS NOT NULL and rs.create_date IS NOT NULL" +
                $" order by close_date desc";
                cmd = new SqlCommand(sql,db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Dur()
                    {
                        req_no = $"{dr["req_no"]}",
                        s_date = Convert.ToDateTime($"{dr["s_date"]}", CultureInfo.CurrentCulture).ToString("yyyy'-'MM'-'dd"),
                        e_date = Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("yyyy'-'MM'-'dd"),
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
            return Json(res,JsonRequestBehavior.AllowGet);
        }
        public JsonResult UserAuthen()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            bool res = false;
            try
            {
                db.OpenConn();
                string sql = $"select * from role_matrix where emp_id = @emp_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@emp_id", System.Data.SqlDbType.Int).Value = "26010298";
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
            finally
            {
                db.CloseConn();
            }
            return Json(res, JsonRequestBehavior.AllowGet);
        }
    }
}