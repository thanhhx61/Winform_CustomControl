using Npgsql;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;


namespace AITCallProcedure
{
    class AITConnect
    {
        /// <summary>
        /// Connect SqlServer
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="procName"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        #region Sqlserver
        public List<T> ConnectSqlServer<T>(string procName, object param = null) where T : new()
        {
            {
                List<string> paramString = new List<string>();
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        paramString.Add("@" + i.Name);
                }
                var stringQuery = procName + " " + (string.Join(", ", paramString));
                List<T> lisstData = new List<T>();
                string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
                SqlConnection MyConnection = new SqlConnection(connetionString);
                MyConnection.Open();
                SqlCommand cmd = new SqlCommand(stringQuery, MyConnection);
                cmd.CommandType = CommandType.Text;
                if (param != null)
                {
                    foreach (var i in param.GetType().GetProperties())
                    {
                        if (i.GetValue(param) != null)
                            cmd.Parameters.AddWithValue("@" + i.Name, i.GetValue(param));
                    }
                }
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    T objCodeName = new T();
                    PropertyInfo[] properties = objCodeName.GetType().GetProperties();
                    foreach (PropertyInfo prop in properties)
                    {
                        if (dr[prop.Name] != null && dr[prop.Name] != DBNull.Value)
                            prop.SetValue(objCodeName, dr[prop.Name]);
                    }
                    lisstData.Add(objCodeName);
                }
                dr.Close();

                return lisstData;
            }
        }

        public T ConnectSqlServer<T>(string procName, object param, string name)
        {
            {
                List<string> paramString = new List<string>();
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        paramString.Add("@" + i.Name);
                }
                var stringQuery = procName + " " + (string.Join(", ", paramString));
                string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
                SqlConnection MyConnection = new SqlConnection(connetionString);
                MyConnection.Open();
                SqlCommand cmd = new SqlCommand(stringQuery, MyConnection);
                cmd.CommandType = CommandType.Text;
                if (param != null)
                {
                    foreach (var i in param.GetType().GetProperties())
                    {
                        if (i.GetValue(param) != null)
                            cmd.Parameters.AddWithValue("@" + i.Name, i.GetValue(param));
                    }
                }
                SqlDataReader dr = cmd.ExecuteReader();
                object value = null;
                while (dr.Read())
                {
                    value = dr[name];
                }
                dr.Close();

                return (T)Convert.ChangeType(value, typeof(T));
            }
        }
        #endregion


        /// <summary>
        /// Connect PostgreSql
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="procName"></param>
        /// <param name="param"></param>
        /// <returns></returns>

        #region PosgresSql
        public List<T> ConnectSqlPostgres<T>(string procName, object param = null) where T : new()
        {
            List<string> paramString = new List<string>();
            foreach (var i in param.GetType().GetProperties())
            {
                if (i.GetValue(param) != null)
                    paramString.Add("@" + i.Name);
            }
            var stringQuery = procName + " " + (string.Join(", ", paramString));
            List<T> lisstData = new List<T>();
            string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
            NpgsqlConnection MyConnection = new NpgsqlConnection(connetionString);
            MyConnection.Open();
            NpgsqlCommand cmd = new NpgsqlCommand(stringQuery, MyConnection);
            cmd.CommandType = CommandType.Text;
            if (param != null)
            {
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        cmd.Parameters.AddWithValue("@" + i.Name, i.GetValue(param));
                }
            }
            NpgsqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                T objCodeName = new T();
                PropertyInfo[] properties = objCodeName.GetType().GetProperties();
                foreach (PropertyInfo prop in properties)
                {
                    if (dr[prop.Name] != null && dr[prop.Name] != DBNull.Value)
                        prop.SetValue(objCodeName, dr[prop.Name]);
                }
                lisstData.Add(objCodeName);
            }
            dr.Close();

            return lisstData;
        }
        public T ConnectSqlPostgres<T>(string procName, object param, string name)
        {

            List<string> paramString = new List<string>();
            foreach (var i in param.GetType().GetProperties())
            {
                if (i.GetValue(param) != null)
                    paramString.Add("@" + i.Name);
            }
            var stringQuery = procName + " " + (string.Join(", ", paramString));
            string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
            NpgsqlConnection MyConnection = new NpgsqlConnection(connetionString);
            MyConnection.Open();
            NpgsqlCommand cmd = new NpgsqlCommand(stringQuery, MyConnection);
            cmd.CommandType = CommandType.Text;
            foreach (var i in param.GetType().GetProperties())
            {
                if (i.GetValue(param) != null)
                    cmd.Parameters.AddWithValue("@" + i.Name, i.GetValue(param));
            }
            NpgsqlDataReader dr = cmd.ExecuteReader();
            object value = null;
            while (dr.Read())
            {
                value = dr[name];
            }
            dr.Close();
            return (T)Convert.ChangeType(value, typeof(T));
        }
        #endregion


        /// <summary>
        /// Connect OracleSql
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="procName"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        #region OracleSql
        public List<T> ConnectOracleSql<T>(string procName, object param = null) where T : new()
        {
            List<string> paramString = new List<string>();
            foreach (var i in param.GetType().GetProperties())
            {
                if (i.GetValue(param) != null)
                    paramString.Add("@" + i.Name);
            }
            var stringQuery = procName + " " + (string.Join(", ", paramString));
            List<T> lisstData = new List<T>();
            string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
            OracleConnection MyConnection = new OracleConnection(connetionString);
            MyConnection.Open();
            OracleCommand cmd = new OracleCommand(stringQuery, MyConnection);
            cmd.CommandType = CommandType.Text;
            cmd.Connection = MyConnection;
            if (param != null)
            {
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        cmd.Parameters.Add("@" + i.Name, i.GetValue(param));
                }
            }
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                T objCodeName = new T();
                PropertyInfo[] properties = objCodeName.GetType().GetProperties();
                foreach (PropertyInfo prop in properties)
                {
                    if (dr[prop.Name] != null && dr[prop.Name] != DBNull.Value)
                        prop.SetValue(objCodeName, dr[prop.Name]);
                }
                lisstData.Add(objCodeName);
            }
            dr.Close();

            return lisstData;
        }
        public T ConnectOracleSql<T>(string procName, object param, string name)
        {
            {
                List<string> paramString = new List<string>();
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        paramString.Add("@" + i.Name);
                }
                var stringQuery = procName + " " + (string.Join(", ", paramString));
                string connetionString = ConfigurationManager.ConnectionStrings["connetionString"].ConnectionString;
                OracleConnection MyConnection = new OracleConnection(connetionString);
                MyConnection.Open();
                OracleCommand cmd = new OracleCommand(stringQuery, MyConnection);
                cmd.CommandType = CommandType.Text;
                cmd.Connection = MyConnection;
                foreach (var i in param.GetType().GetProperties())
                {
                    if (i.GetValue(param) != null)
                        cmd.Parameters.Add("@" + i.Name, i.GetValue(param));
                }
                OracleDataReader dr = cmd.ExecuteReader();
                object value = null;
                while (dr.Read())
                {
                    value = dr[name];
                }
                dr.Close();

                return (T)Convert.ChangeType(value, typeof(T));
            }
        }
        #endregion
    }
    class Message
    {
        public IEnumerable data { get; set; }
        public string message { get; set; }
    }
}