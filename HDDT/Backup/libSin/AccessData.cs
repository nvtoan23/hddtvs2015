using System;
using System.Data;
using System.Xml;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Npgsql;
using NpgsqlTypes;

namespace libSin
{
    public class AccessData
    {
        NpgsqlDataAdapter dest;
        NpgsqlConnection con;
        NpgsqlCommand cmd;
        string sConn = "Server=192.168.1.14;Port=5432;User Id=medisoft;Password=links1920;Database=hddt;Encoding=UNICODE;Pooling=true;CommandTimeout=30000;";
        string sComputer = null;
        public string s_ngonngu = "";
        string dblink = "";
        string sHost = "", sPort = "";
        const string key = "9189";        
        string sql = "", schema = "hddt", owner = "medisoft", password = "links1920", userid = "hddt", database = "hddt";
        DataSet ds = new DataSet();        
        public static AccessData GetImplement()
        {
            if (init == null)
                init = new AccessData();
            return init;
        }
        static AccessData init = null;
        public AccessData()
        {

            string s_Key = Maincode("Key");
            if (s_Key == "1")
            {
                if (Maincode("Ip") != "") sHost = DeCode(Maincode("Ip"), key);
                if (Maincode("Post") != "") sPort = DeCode(Maincode("Post"), key);
                if (Maincode("UserID") != "") owner = DeCode(Maincode("UserID"), key);
                if (Maincode("Password") != "") password = DeCode(Maincode("Password"), key);
                if (Maincode("Database") != "") database = DeCode(Maincode("Database"), key);
            }
            else
            {
                if (Maincode("Ip") != "") sHost = Maincode("Ip");
                if (Maincode("Post") != "") sPort = Maincode("Post");
                if (Maincode("UserID") != "") owner = Maincode("UserID");
                if (Maincode("Password") != "") password = Maincode("Password");
                if (Maincode("Database") != "") database = Maincode("Database");
            }
            if (Maincode("User") != "") userid = Maincode("User");
            if (Maincode("ngonngu") != "") 
            userid = userid + s_ngonngu;
            sComputer = System.Environment.MachineName.Trim().ToUpper();
            sConn = "Server=" + sHost + ";Port=" + sPort + ";User Id=" + owner + ";Password=" + password + ";Database=" + database + ";Encoding=UNICODE;Pooling=true;";
           
        }
        public string Maincode(string sql)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("..//..//..//xml//maincode.xml");
                XmlNodeList nodeLst = doc.GetElementsByTagName(sql);
                return nodeLst.Item(0).InnerText;
            }
            catch
            {
                ds = new DataSet();
                ds.ReadXml("..//..//..//xml//maincode.xml");
                DataColumn dc = new DataColumn();
                dc.ColumnName = sql;
                dc.DataType = Type.GetType("System.String");
                ds.Tables[0].Columns.Add(dc);
                ds.Tables[0].Rows[0][sql] = "";
                ds.WriteXml("..//..//..//xml//maincode.xml");
                return "";
            }
        }
        internal string DeCode(string values, string s_key)
        {
            string val = "";
            val = DeCrypt(values, s_key);
            return val;
        }

        string DeCrypt(string values, string s_key)
        {
            byte[] keyArray;
            byte[] toEncryptArray = Convert.FromBase64String(values);
            MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
            keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(s_key));
            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
        public string mmyy(string ngay)
        {
            if (ngay.Length == 4) return ngay;
            else return ngay.Substring(3, 2) + ngay.Substring(8, 2);
        }
        public DataSet get_data(string sql)
        {
            //string ammyy = "";
            //ammyy = mmyy(DateTime.Today.ToString("dd/MM/yyyy"));
            //sql = sql.Replace("medibvmmyy.", user + ammyy + ".");
            //sql = sql.Replace("medibv.", user + ".");
            //sql = sql.Replace("xxx.", user + ammyy + ".");
            //string sqlLower = sql.ToLower();
            return get_data0(sql);

        }
        public DataSet get_data0(string sql)
        {
            DataSet dstmp = new DataSet();
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            con = new NpgsqlConnection(sConn);
            try
            {
                con.Open();

                //sql = LibMedi.Medisoft_S.GetInstanle().SQL_MedisoftS(sql);
                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                dest = new NpgsqlDataAdapter(cmd);
                dest.Fill(dstmp);
                cmd.Dispose();
                con.Close();
                con.Dispose();
            }
            catch (NpgsqlException ex)
            {
                upd_error(sql + "-" + ex.Message.ToString().Trim(), "?");
                //Log.getInstanlize().WriteLog("Libmedi.getData0", sql + " \n\t\tFAIL:" + ex.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            dstmp.AcceptChanges();
            return dstmp;
        }
        public string user { get { return userid; } }
        public void upd_error(string m_message, string m_table)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            //if (!bMmyy(mmyy(m_ngay))) return;
            sql = "insert into " + user + ".error(message,tables) values (:m_message,:m_table)";
            con = new NpgsqlConnection(sConn);
            try
            {
                con.Open();

                //sql = LibMedi.Medisoft_S.GetInstanle().SQL_MedisoftS(sql);
                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_message", NpgsqlDbType.Text).Value = m_message;
                cmd.Parameters.Add("m_table", NpgsqlDbType.Varchar, 20).Value = m_table;
                cmd.ExecuteNonQuery();
            }
            catch { }
            finally
            {
                cmd.Dispose();
                con.Close(); con.Dispose();
            }
        }
        public bool execute_data(string sql)
        {

            sql = sql.Replace("medibv.", user + ".");
            try
            {
                if (con != null)
                {
                    con.Close(); con.Dispose();
                }
                con = new NpgsqlConnection(sConn);
                con.Open();

                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                DateTime n = DateTime.Now;
                cmd.ExecuteNonQuery();
                TimeSpan s = DateTime.Now.Subtract(n);
                cmd.Dispose();
                con.Close(); con.Dispose();
                return true;
            }
            catch (NpgsqlException ex)
            {
                upd_error(sql + "; " + ex.Message.ToString().Trim(), "?");
                return false;
            }
        }
    }
}
