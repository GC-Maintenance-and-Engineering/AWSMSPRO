using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MeExpert.CommonProvide
{
    public class AuthenFilter: AuthorizeAttribute
    {
        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            try
            {
                bool isauthen = filterContext.HttpContext.User.Identity.IsAuthenticated;
                string strname = filterContext.HttpContext.User.Identity.Name;
                string[] user_id = strname.Split('\\');
                if (isauthen)
                {
                    Connect_DB db = new Connect_DB();
                    SqlDataReader dr;
                    SqlCommand cmd;
                    db.OpenConn();
                    try
                    {
                        string sql = $"select * from role_matrix where emp_id = @emp_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Parameters.Add("@emp_id", System.Data.SqlDbType.Int).Value = user_id[1];
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            if ($"{dr["role_name"]}".ToLower().Equals("admin"))
                            {
                                base.OnAuthorization(filterContext);
                            }
                            else
                            {
                                filterContext.HttpContext.Response.RedirectToRoute(new { controller = "Home", action = "Permission", message = "incorrectuser" });
                            }
                        }
                        else
                        {
                            filterContext.HttpContext.Response.RedirectToRoute(new { controller = "Home", action = "Permission", message = "althenfailed" });
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
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
           
        }
    }
}