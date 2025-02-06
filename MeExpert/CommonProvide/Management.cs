
using MeExpert.Controllers;
using MeExpert.Models;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace MeExpert.CommonProvide
{
    public class Management
    {
        public List<Client> GetClients() {
			Connect_DB db = new Connect_DB();
			SqlDataReader dr;
			SqlCommand cmd;
			List<Client> res = new List<Client>();
			try
			{
				db.OpenConn();
				string sql = $"select * from client";
				cmd = new SqlCommand(sql, db.GetConn());
				dr = cmd.ExecuteReader();
				while (dr.Read())
				{
					res.Add(new Client()
					{
						client_id = (int)dr["client_id"],
						client_name = $"{dr["client_name"]}",
						client_address = $"{dr["client_address"]}",
						client_tel = $"{dr["client_tel"]}",
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
			return res;
        }
        public List<Position> GetPosition()
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
            return res;
        }
        public List<Cand> GetCands()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<Cand> res = new List<Cand>();
            var cont = new RequisitionController(); 
            try
            {
                db.OpenConn();
                string sql = $"select cand_id,first_name,last_name,gender,mail,mobile_no_1,ap.apply_position,apply_date,exp_salary,cv_id " +
                    $"from candidate as c outer apply(select top 1 apply_position,apply_date,exp_salary,cv_no from apply_record as a " +
                    $"where  c.cand_id = a.cand_id) as ap "+
                    $"inner join cv_details cv on cv.cv_no=ap.cv_no" ;
                cmd = new SqlCommand(sql,db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Cand()
                    {
                        cv_id = (int)dr["cv_id"],
                        cand_id = (int)dr["cand_id"],
                        first_name = $"{dr["first_name"]}",
                        last_name = $"{dr["last_name"]}",
                        gender = $"{dr["gender"]}",
                        mail = string.IsNullOrEmpty($"{dr["mail"]}") ? string.Empty : $"{dr["mail"]}",
                        mobile_no_1 = $"{dr["mobile_no_1"]}",
                        apply_position = $"{dr["apply_position"]}",
                        //cv_path = cont.GetCvPath((int)dr["cand_id"]),
                        apply_date = string.IsNullOrEmpty($"{dr["apply_date"]}") ? string.Empty : Convert.ToDateTime($"{dr["apply_date"]}", CultureInfo.CurrentCulture).ToString("dd'-'MMM'-'yyyy"),
                        exp_salary = string.IsNullOrEmpty($"{dr["exp_salary"]}") ? string.Empty : $"{dr["exp_salary"]}"
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
            return res;
        }
        public async Task<UserApi> GetUserDetail(string usr_id) 
        {
            UserApi res = new UserApi();
            clsApi cls = new clsApi();
            try
            {
                var user_det = await cls.wsGet($"CDEC/GetStaffWithCode?strEmpCode={usr_id}");
                if (user_det.IsSuccessStatusCode)
                {
                    string res_str = await user_det.Content.ReadAsStringAsync();
                    res = JsonConvert.DeserializeObject<UserApi>(res_str);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return res;
        }
        public async Task<List<UserApi>> GetUserList()
        {
            List<UserApi> res = new List<UserApi>();
            clsApi clsApi = new clsApi();
            string res_str = string.Empty;
            try
            {
                var user_list = await clsApi.wsGet("CDEC/GetAllStaff");
                if (user_list.IsSuccessStatusCode)
                {
                    res_str = await user_list.Content.ReadAsStringAsync();
                    res = JsonConvert.DeserializeObject<List<UserApi>>(res_str);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return res;
        }
        public List<Cand> GetCandByReq(int req_id)
        {
            List<Cand> res = new List<Cand>();
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"select r.cand_id,c.first_name,c.last_name,req_cand_id from req_cand as r inner join candidate as c on c.cand_id = r.cand_id where r.req_id=@req_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_id",SqlDbType.Int).Value = req_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new Cand() { 
                        cand_id = (int)dr["cand_id"],
                        full_name = $"{dr["first_name"]} {dr["last_name"]}",
                        req_cand_id = (int)dr["req_cand_id"],
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
       
        public static bool InsertLog(string module_name,string log_desc)
        {
            bool res = false;
            SqlCommand cmd;
            string sql = string.Empty;
            Connect_DB db = new Connect_DB();
            var usr_str = HttpContext.Current.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            try
            {
                db.OpenConn();
                sql = $"insert into log_list (usr_id,module_name,log_desc) values(@usr_id,@module_name,@log_desc)";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@usr_id", SqlDbType.Int).Value = usr_id[1];
                cmd.Parameters.Add("@module_name", SqlDbType.NVarChar, 60).Value = module_name;
                cmd.Parameters.Add("@log_desc", SqlDbType.NVarChar, 200).Value = log_desc;
                cmd.ExecuteNonQuery();
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
        public List<Education> GetInst(string type)
        {
            List<Education> res = new List<Education>();
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;

            try
            {
                db.OpenConn();
                switch (type)
                {
                    case "edu":
                        sql = $"select inst_name from education group by inst_name";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new Education()
                            {
                                inst_name = $"{dr["inst_name"]}",
                            });
                        }
                        dr.Close();
                        break;
                    case "major":
                        sql = $"select major from education group by major";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new Education()
                            {
                                major = $"{dr["major"]}",
                            });
                        }
                        dr.Close();
                        break;
                    case "faculty":
                        sql = $"select faculty from education group by faculty";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new Education()
                            {
                                faculty = $"{dr["faculty"]}",
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
            finally {
                db.CloseConn();
            }
            return res;
        }
        public List<ContactPerson> GetPrsByClientID(int client_id)
        {
            List<ContactPerson> res = new List<ContactPerson>();
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;

            try
            {
                db.OpenConn();
                string sql = $"select cont_id, cont_person_name from contact_person where client_id=@client_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new ContactPerson() { 
                        cont_id = (int)dr["cont_id"],
                        cont_person_name = $"{dr["cont_person_name"]}",
                    });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally {
                db.CloseConn();
            }
            return res;
        }
        public List<ApplyRecord> GetExperience(string type)
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<ApplyRecord> res = new List<ApplyRecord>();
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                switch (type)
                {
                    case "industry":
                        sql = $"select industry  from apply_record where industry is not null and industry != '' group by industry";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new ApplyRecord()
                            {
                                industry = $"{dr["industry"]}",
                            });
                        }
                        dr.Close();
                        break;
                    case "keyword":
                        sql = $"select keyword  from apply_record where keyword is not null and keyword != '' group by keyword";
                        cmd = new SqlCommand(sql, db.GetConn());
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            res.Add(new ApplyRecord()
                            {
                                keyword = $"{dr["keyword"]}",
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
            return res;
        }
        public List<FilterReq> GetYesrCreateReq()
        {
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            List<FilterReq> res = new List<FilterReq>();

            try
            {
                db.OpenConn();
                string sql = $"select  datepart(YEAR, create_date) as year from requisition group by datepart(YEAR, create_date) order by YEAR desc";
                cmd = new SqlCommand(sql, db.GetConn());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    res.Add(new FilterReq()
                    {
                        year = (int)dr["year"],
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
        public static int EncodeToNumber(string encodestr)
        {
            int res = 0;
            try
            {
                var decode = Uri.UnescapeDataString(encodestr.ToString());
                byte[] data = Convert.FromBase64String(decode);
                string decodedString = Encoding.UTF8.GetString(data);
                res = Convert.ToInt32(decodedString);
            }
            catch (Exception ex)
            {

                throw ex;

            }
            return res;
        }
    }
}