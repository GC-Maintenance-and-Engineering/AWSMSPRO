using MeExpert.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;

namespace MeExpert.CommonProvide
{
    public class ConvertFileSream
    {
        public static byte[] ConvertFileStream(Stream file_stream)
        {
            byte[] file;
            file = new byte[file_stream.Length];
            file_stream.Read(file, 0, file.Length);
            byte[] fileStream = file;
            return fileStream;
            //SaveCvFile(fileStream, cv_no);
        }
        public static bool SaveCvFile(string doc_name,string doc_extension, string cv_no)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            bool res = false;
            string sql = string.Empty;
            DateTime dt = DateTime.Now;
            string dtformat = dt.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var usr_str = HttpContext.Current.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            try
            {
                db.OpenConn();
                sql = $"insert into cv_details (doc_name,doc_extension,cv_no,create_by,create_time) values(@doc_name,@doc_extension,@cv_no,@create_by,@create_time);";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@doc_name", SqlDbType.NVarChar).Value = doc_name;
                cmd.Parameters.Add("@doc_extension", SqlDbType.NVarChar).Value = doc_extension;
                cmd.Parameters.Add("@cv_no", SqlDbType.NVarChar,20).Value = cv_no;
                cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = usr_id[1];
                cmd.Parameters.Add("@create_time", SqlDbType.DateTime).Value = dtformat;
                cmd.ExecuteNonQuery();
                res = true;
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
        public static bool SaveReqFile(string doc_name, string doc_extension, int req_id)
        {
            bool res = false;
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            string sql = string.Empty;
            try
            {
                db.OpenConn();
                sql = $"update requisition set file_extension=@file_extension, file_name=@file_name where req_id=@req_id";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@req_id", SqlDbType.Int).Value = req_id;
                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = doc_name;
                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = doc_extension;
                cmd.ExecuteNonQuery();
                res = true;
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
        public static bool SavePpsFile(string file_name,string file_extension, int succ_id,int rev)
        {
            bool res = false;
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string sql = string.Empty;
            var usr_str = HttpContext.Current.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            try
            {
                db.OpenConn();
                sql = $"insert proposal_list (succ_id,file_name,file_extension,create_by,rev_no) values(@succ_id,@file_name,@file_extension,@create_by,@rev_no)";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@succ_id", SqlDbType.Int).Value = succ_id;
                cmd.Parameters.Add("@rev_no", SqlDbType.Int).Value = rev;
                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = file_name;
                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = file_extension;
                cmd.Parameters.Add("@create_by", SqlDbType.NVarChar, 13).Value = usr_id[1];
                cmd.ExecuteNonQuery();
                res = true;
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
        public static bool SaveSrvFile(string file_name,string file_extension, int client_id, int rev)
        {
            Connect_DB db = new Connect_DB();
            SqlCommand cmd;
            bool res = false;
            var usr_str = HttpContext.Current.User.Identity.Name;
            var usr_id = usr_str.Split('\\');
            try
            {
                db.OpenConn();
                string sql = $"insert into srv_agreement_list (client_id,rev_no,file_name,file_extension,create_by) values(@client_id,@rev_no,@file_name,@file_extension,@create_by)";
                cmd = new SqlCommand(sql,db.GetConn());
                cmd.Parameters.Add("@client_id", SqlDbType.Int).Value = client_id;
                cmd.Parameters.Add("@file_extension", SqlDbType.NVarChar).Value = file_extension;
                cmd.Parameters.Add("@file_name", SqlDbType.NVarChar).Value = file_name;
                cmd.Parameters.Add("@rev_no", SqlDbType.Int).Value = rev;
                cmd.Parameters.Add("@create_by", SqlDbType.NVarChar,10).Value = usr_id[1];
                cmd.ExecuteNonQuery();
                db.CloseConn();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return res;
        }
    }
}