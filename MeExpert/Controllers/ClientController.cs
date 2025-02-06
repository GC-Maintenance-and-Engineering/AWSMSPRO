using MeExpert.CommonProvide;
using MeExpert.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace MeExpert.Controllers
{
    //[AuthenFilter]
    public class ClientController : Controller
    {
        // GET: Client
        public ActionResult Index()
        {
            return View();
        } 

        public async Task<dynamic> Test(string emp_code)
        {
            return null;
        }
        private string UserAuthen()
        {
            var usr_str = HttpContext.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            return usr_id[1];
        }
        public JsonResult Get(int client_id) 
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlDataAdapter da;
            SqlCommand cmd;
            List<Client> res = new List<Client>();
            List<ContactPerson> res_sub = new List<ContactPerson>();
            List<Remark> res_rmk = new List<Remark>();
            string sql = string.Empty;
            int number = 1;
            try
            {
                db.OpenConn();
                switch (client_id) {
                    case 0:
                        sql = $"select client_id,client_name,client_address,client_tel,cm.remark,cm.create_by,cm.create_date from client c"+
                        $" outer apply(select top 1 remark,create_by,create_date from client_remark m where c.client_id = m.client_id order by m.create_date desc)as cm" +
                        $" order by cm.create_date desc";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new Client()
                            {
                                no = number++,
                                client_id = (int)dr["client_id"],
                                client_name = $"{dr["client_name"]}",
                                client_address = $"{dr["client_address"]}",
                                client_tel = $"{dr["client_tel"]}",
                                remark = (dr["remark"] == DBNull.Value) ? string.Empty : $"{dr["remark"]}",
                                create_by = (dr["create_by"] == DBNull.Value) ? string.Empty : $"({dr["create_by"]}, {Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")})",
                            }) ;
                        }
                        break;
                    default:
                        sql = $"select c.client_id,c.client_name,c.client_address,c.client_tel,c.remark,r.client_id as isitem " +
                            $"from client as c outer apply(select top 1 client_id,create_date " +
                            $"from requisition where client_id = c.client_id) r " +
                            $"where c.client_id=@client_id;";
                        da = new SqlDataAdapter(sql, db.GetConn());
                        da.SelectCommand.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                        DataTable table = new DataTable();
                        da.Fill(table);
                       foreach(DataRow row in table.Rows)
                        {
                            sql = $"select * from contact_person where client_id=@client_id";
                            cmd = new SqlCommand(sql,db.GetConn());
                            cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = (int)row["client_id"];
                            dr = cmd.ExecuteReader();
                            while (dr.Read())
                            {
                                res_sub.Add(new ContactPerson() {
                                    cont_id = (int)dr["cont_id"],
                                    cont_person_mobile = $"{dr["cont_person_mobile"]}",
                                    cont_person_name = $"{dr["cont_person_name"]}",
                                    cont_pos = $"{dr["position"]}",
                                });
                            }
                            dr.Close();

                            sql = $"select * from client_remark where client_id=@client_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = (int)row["client_id"];
                            dr = cmd.ExecuteReader();
                            while (dr.Read())
                            {
                                res_rmk.Add(new Remark()
                                {
                                    item_id = (int)dr["item_id"],
                                    remark = $"{dr["remark"]}",
                                    create_by = $"{dr["create_by"]} ,<br/>{Convert.ToDateTime($"{dr["create_date"]}",CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                                });
                            }
                            dr.Close();

                            res.Add(new Client()
                            {
                                client_id = (int)row["client_id"],
                                client_name = $"{row["client_name"]}",
                                client_address = $"{row["client_address"]}",
                                client_tel = $"{row["client_tel"]}",
                                remark = (row["remark"] == DBNull.Value) ? string.Empty : $"{row["remark"]}",
                                isitem = (row["isitem"] == DBNull.Value) ? false : true, 
                                arr_data = res_sub,
                                arr_rmk = res_rmk,
                            }); 
                        }
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
        [HttpPost]
        public JsonResult Update(Client data,string type)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            SqlTransaction transaction;
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            string sql = string.Empty;
            int client_id = 0;
            try
            {
               
                switch (type)
                {
                    case "insert":
                        sql = $"insert into client (client_name,client_address,client_tel,create_by) " +
                            $"values(@client_name,@client_address,@client_tel,@create_by);SELECT scope_identity();";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_name", SqlDbType.NVarChar, 200).Value = data.client_name;
                        cmd.Parameters.Add("@client_address", SqlDbType.NVarChar, 200).Value = (string.IsNullOrEmpty(data.client_address))? string.Empty : data.client_address;
                        cmd.Parameters.Add("@client_tel", SqlDbType.NVarChar, 20).Value = (string.IsNullOrEmpty(data.client_tel)) ? string.Empty : data.client_tel;
                        cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
                        client_id = Convert.ToInt32(cmd.ExecuteScalar());
                        Management.InsertLog($"Client List", $"Create Client client_name={data.client_name}");
                        break;
                    case "update":
                        sql = $"update client set client_name=@client_name,client_address=@client_address,client_tel=@client_tel" +
                            $" where client_id=@client_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_name", SqlDbType.NVarChar, 200).Value = data.client_name;
                        cmd.Parameters.Add("@client_address", SqlDbType.NVarChar, 200).Value = (string.IsNullOrEmpty(data.client_address)) ? string.Empty : data.client_address;
                        cmd.Parameters.Add("@client_tel", SqlDbType.NVarChar, 20).Value = (string.IsNullOrEmpty(data.client_tel)) ? string.Empty : data.client_tel;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.ExecuteNonQuery();

                        client_id = data.client_id;
                        List <ContactPerson> arr_cont = new List<ContactPerson>();
                        sql = $"select cont_id from contact_person where client_id =@client_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            arr_cont.Add(new ContactPerson() { 
                                cont_id = (int)dr["cont_id"],
                            });
                        }
                        dr.Close();
                        if (arr_cont.Count > 0)
                        {
                            //delete all
                            if (data.arr_data == null)
                            {
                                sql = $"delete contact_person where client_id = @client_id";
                                cmd = new SqlCommand(sql, db.GetConn());
                                cmd.Transaction = transaction; 
                                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //delete if not in new data
                                var res_cont = arr_cont.Where( d => !data.arr_data.Any(r => r.cont_id == d.cont_id));
                                var list_res = res_cont.ToList();
                                if (list_res.Count > 0)
                                {
                                    sql = "delete contact_person where cont_id in ({0})";
                                    string[] paramArr = list_res.Select((x, n) => "@list_in" + n).ToArray();
                                    cmd = new SqlCommand(sql,db.GetConn());
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = string.Format(sql,string.Join(",",paramArr));
                                    for (int i = 0; i < list_res.Count;i++)
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@list_in" + i, list_res[i].cont_id));
                                    }
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        dr.Close();
                        List<Remark> arr_rmk = new List<Remark>();
                        sql = $"select item_id from client_remark where client_id=@client_id";
                        cmd = new SqlCommand(sql,db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            arr_rmk.Add(new Remark() { 
                                item_id = (int)dr["item_id"],
                            });
                        }
                        dr.Close();
                        if (arr_rmk.Count > 0)
                        {
                            if (data.arr_rmk == null)
                            {
                                sql = $"delete client_remark where client_id = @client_id";
                                cmd = new SqlCommand(sql, db.GetConn());
                                cmd.Transaction = transaction;
                                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                var res_rmk = arr_rmk.Where(d => !data.arr_rmk.Any(r => r.item_id == d.item_id));
                                var list_res = res_rmk.ToList();
                                if (list_res.Count > 0)
                                {
                                    sql = "delete client_remark where item_id in ({0})";
                                    string[] parmArr = list_res.Select((x, n) => "@list_in" + n).ToArray();
                                    cmd = new SqlCommand(sql,db.GetConn());
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = string.Format(sql,string.Join(",",parmArr));
                                    for (int i = 0; i < list_res.Count;i++)
                                    {
                                        cmd.Parameters.Add(new SqlParameter("@list_in" + i,list_res[i].item_id)) ;
                                    }
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        Management.InsertLog($"Client List", $"Update Client client_name={data.client_name}");

                        break;
                    case "delete":
                        sql = $"delete contact_person where client_id = @client_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.ExecuteNonQuery();

                        sql = $"delete client_remark where client_id = @client_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.ExecuteNonQuery();

                        sql = $"delete srv_agreement_list where client_id = @client_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.ExecuteNonQuery();

                        sql = $"delete client where client_id=@client_id";
                        cmd = new SqlCommand(sql, db.GetConn());
                        cmd.Transaction = transaction;
                        cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = data.client_id;
                        cmd.ExecuteNonQuery();
                        Management.InsertLog($"Client List", $"Delete Client client_id={data.client_id}");

                        data.arr_data = null;
                        break;
                }
                //insert and update contact person
                if (data.arr_data != null)
                {
                    for (var i = 0; i < data.arr_data.Count; i++)
                    {
                        if (data.arr_data[i].cont_id == 0)
                        {
                            sql = $"insert into contact_person (client_id,cont_person_name,cont_person_mobile,position) " +
                            $"values(@client_id,@cont_person_name,@cont_person_mobile,@position)";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                            cmd.Parameters.Add("@cont_person_name", SqlDbType.NVarChar, 100).Value = data.arr_data[i].cont_person_name;
                            cmd.Parameters.Add("@cont_person_mobile", SqlDbType.NVarChar, 20).Value = string.IsNullOrEmpty(data.arr_data[i].cont_person_mobile) ? string.Empty : data.arr_data[i].cont_person_mobile;
                            cmd.Parameters.Add("@position", SqlDbType.NVarChar, 50).Value = string.IsNullOrEmpty(data.arr_data[i].cont_pos) ? string.Empty : data.arr_data[i].cont_pos;
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            sql = $"update contact_person set cont_person_name=@cont_person_name,cont_person_mobile=@cont_person_mobile,position=@position where cont_id =@cont_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@cont_id", SqlDbType.Int).Value = data.arr_data[i].cont_id;
                            cmd.Parameters.Add("@cont_person_name", SqlDbType.NVarChar, 100).Value = data.arr_data[i].cont_person_name;
                            cmd.Parameters.Add("@cont_person_mobile", SqlDbType.NVarChar, 20).Value = string.IsNullOrEmpty(data.arr_data[i].cont_person_mobile) ? string.Empty : data.arr_data[i].cont_person_mobile;
                            cmd.Parameters.Add("@position", SqlDbType.NVarChar, 50).Value = string.IsNullOrEmpty(data.arr_data[i].cont_pos) ? string.Empty : data.arr_data[i].cont_pos;
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                if (data.arr_rmk != null)
                {
                    for (var i =0; i < data.arr_rmk.Count;i++)
                    {
                        if (data.arr_rmk[i].item_id == 0)
                        {
                            sql = $"insert into client_remark (client_id,remark,create_by) values (@client_id,@remark,@create_by)";
                            cmd = new SqlCommand(sql,db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                            cmd.Parameters.Add("@remark", SqlDbType.NVarChar,250).Value = data.arr_rmk[i].remark;
                            cmd.Parameters.Add("@create_by",SqlDbType.NVarChar,13).Value = UserAuthen();
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            sql = $"update client_remark set remark=@remark,create_by=@create_by where item_id=@item_id";
                            cmd = new SqlCommand(sql, db.GetConn());
                            cmd.Transaction = transaction;
                            cmd.Parameters.Add("@item_id", SqlDbType.Int).Value = data.arr_rmk[i].item_id;
                            cmd.Parameters.Add("@remark", SqlDbType.NVarChar, 250).Value = data.arr_rmk[i].remark;
                            cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = UserAuthen();
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
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetReqByClentID(int id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Req> res = new List<Req>();
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select req_id,req_no,title,p.position_name,r.req_status from requisition as r " +
                      $"inner join position as  p on r.position_id = p.position_id where r.client_id = @id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Req()
                    {
                        req_id = (int)dr["req_id"],
                        req_no = $"{dr["req_no"]}",
                        title = $"{dr["title"]}",
                        position = $"{dr["position_name"]}",
                        req_status = $"{dr["req_status"]}",
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
        public ActionResult SaveClientFile()
        {
            var client_id = Request["client_id"];
            var srv_file = Request["srv_file_name"];
            HttpFileCollectionBase files = Request.Files;
            int rev = 0;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            var fullPath = string.Empty;


            try
            {
                //db.OpenConn();
                //sql = $"select top 1 rev_no from srv_agreement_list where client_id=@client_id  order by rev_no desc";
                //cmd = new SqlCommand(sql, db.GetConn());
                //cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                //dr = cmd.ExecuteReader();
                //if (dr.Read())
                //{
                //    rev = (int)dr["rev_no"];
                //    rev++;
                //}
                //dr.Close();
                //db.CloseConn();
                string fname = srv_file;
                if (files.Count != 0)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        //var file = files[i];
                        //Stream file_stream = file.InputStream;
                        //byte[] filestream = ConvertFileSream.ConvertFileStream(file_stream);
                        HttpPostedFileBase file = files[i];
                        string file_name = files[i].FileName;
                        float file_size = files[i].ContentLength;
                        string doc_extension = Path.GetExtension(file_name);
                        string doc_name = fname + doc_extension;
                        fullPath = Path.Combine(Server.MapPath("~/Uploads/SrvAgr/"), doc_name);
                        file.SaveAs(fullPath);
                        ConvertFileSream.SaveSrvFile(doc_name, doc_extension, Convert.ToInt32(client_id),  rev);
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
            return Json(JsonRequestBehavior.DenyGet);
        }
        public JsonResult GetSrvByClientID(int client_id)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<SrvFile> res = new List<SrvFile>();
            try
            {
                db.OpenConn();
                string sql = $"select * from srv_agreement_list where client_id=@client_id order by rev_no asc";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new SrvFile()
                    {
                        srv_id = (int)dr["srv_id"],
                        file_name = $"{dr["file_name"]}",
                        rev_no = (int)dr["rev_no"],
                        create_by = $"{dr["create_by"]}, {Convert.ToDateTime($"{dr["create_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy")}",
                    });
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
        public ActionResult OpenClientFile(string srv_id)
        {
            int int_srv_id = Management.EncodeToNumber(srv_id);
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
                string sql = $"select file_name,file_extension from srv_agreement_list where srv_id=@srv_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@srv_id", SqlDbType.Int).Value = int_srv_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    file_name = $"{dr["file_name"]}";
                    file_extension = $"{dr["file_extension"]}";
                }
                dr.Close();
                path = Path.Combine(Server.MapPath("~/Uploads/SrvAgr/"), file_name);
                res_file = System.IO.File.ReadAllBytes(path);
                type = ReadFileType.GetTypeFile(file_extension);
                db.CloseConn();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return File(res_file, type);
        }
        [HttpPost]
        public JsonResult DeleteSrv(int srv_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select file_name from srv_agreement_list where srv_id = @srv_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@srv_id", SqlDbType.Int).Value = srv_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    string file_name = $"{dr["file_name"]}";
                    dr.Close();

                    string file_path = Path.Combine(Server.MapPath("~/Uploads/SrvAgr"), file_name);

                    if (System.IO.File.Exists(file_path))
                    {
                        System.IO.File.Delete(file_path);
                    }
                }
                sql = $"delete srv_agreement_list where srv_id=@srv_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@srv_id", SqlDbType.Int).Value = srv_id;
                cmd.ExecuteNonQuery();
                db.CloseConn();

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Json(JsonRequestBehavior.DenyGet);
        }
        public bool DeleteClient(int client_id)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            SqlDataReader dr;
            SqlTransaction transaction;
            bool res = false;
            string sql = string.Empty;
            List<int> cl_list = new List<int>();
            db.OpenConn();
            transaction = db.GetConn().BeginTransaction();
            try
            {
                sql = $"delete srv_agreement_list where client_id=@client_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();

                sql = $"delete client_remark where client_id=@client_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();

                sql = $"delete contact_person where client_id=@client_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();

                sql = $"delete client where client_id=@client_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();

                transaction.Commit();
                res = true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                throw ex;
            }
            return res;
        }
    }
}