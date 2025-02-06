using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace MeExpert.CommonProvide
{
    public class Connect_DB
    {
        private SqlConnection con;
        string connstr;
        public Connect_DB()
        {
            connstr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
            con = new SqlConnection(connstr);
            try
            {
                con.Open();
                con.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public bool OpenConn()
        {
            bool isopen = false;
            try
            {
                con.Open();
                isopen = true;

            }
            catch (Exception ex)
            {

                throw ex;
            }
            return isopen;
        }
        public bool CloseConn()
        {
            bool isclose = false;
            try
            {
                con.Close();
                isclose = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isclose;
        }
        public SqlConnection GetConn()
        {
            return con;
        }
    }
}