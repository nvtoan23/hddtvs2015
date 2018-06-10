using System;
using Microsoft.Win32;
using System.Data;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using dichso;
using ThongSoData.Medisoft;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace LibMedi
{
    public class AccessData
    {
        public const string Msg = "Medisoft THIS", Normal = "Bình thường", Msg_di = "Dinh dưỡng";
        public const string links_userid = "links", links_pass = "link7155019s20", uxxx = "tuneig@mp", pxxx = "tuneig@mp", xxxxx = "¯Ò¡Ì©Î«³²°Ô£";
        public const int Tiepdon = 0, Khambenh = 1, Ngoaitru = 2, Phongluu = 3, Nhanbenh = 4, Khoa = 5, Phauthuthuat = 6, Vienphi = 7, Duoc = 8, Xetnghiem = 9, Sieuam = 10, Noisoi = 11, Xquang = 12, Le = 13, Khucxa = 14, Cls = 15, Taikham = 16, Pttt = 17;
        public const int giamdoc = 1, phogiamdoc = 2, truongkhoa = 3, nhanvien = 8, nghiviec = 9, ybs_cls = 10;
        public int iHaophi = 5;
        string sConn = "Data Source=MEDISOFT;user id=MEDIBV;password=MEDIBV";
        string sConn_mysql = "SERVER=127.0.0.1;PORT=3306;DATABASE=dbtgmedisoft;UID=tgmedisoft;PASSWORD=huyethocxHd!23";
        [DllImport("winmm.dll")]
        private static extern bool PlaySound(string lpszName, int hModule, int dwFlags);
        private int iRownum = 1;
        private decimal _e = 0, _p = 0, _l = 0, _g = 0;
        private string ma16_ngtru_quyenloimoi = "";
        private string ma18_ngtru_khongchitra = "";

        OracleDataAdapter dest;
        OracleConnection con;
        OracleCommand cmd;

        MySqlConnection con_mysql;
        MySqlDataAdapter dest_mysql;
        MySqlCommand cmd_mysql;

        string sComputer = null;
        string m_hotenkdau, sql = "", userid = "medibv", service_name = "medisoft", b_sobienlaitamung = "", b_tatca_sobienlaitamung = "";
        decimal tc_tamung = 0;
        bool b_sovaovien, b_soluutru;
        DataSet ds = null;
        DoiTuong _ctsdoituong;

        public AccessData()
        {
            if (Maincode("Con") != "") sConn = Maincode("Con");
            sConn_mysql = mySqlConnection();
            sComputer = System.Environment.MachineName.Trim().ToUpper();            
            userid = sConn.Substring(sConn.LastIndexOf("=") + 1).Trim();
            service_name = sConn.Substring(sConn.IndexOf("=") + 1, sConn.IndexOf(";") - 1 - sConn.IndexOf("=")).Trim();
            ds = get_data("select rownum,computer from dmcomputer");
            DataRow r = getrowbyid(ds.Tables[0], "computer='" + sComputer + "'");
            if (r != null) iRownum = int.Parse(r["rownum"].ToString());
            _ctsdoituong = new DoiTuong();
        }

        public string user { get { return userid; } }
        public string service { get { return service_name; } }
        private string get_conn_d(string mmyy)
        {
            return "Data Source=" + service_name + ";user id=" + userid + "d" + mmyy + ";password=" + userid + "d" + mmyy;
        }
        private string get_conn_v(string mmyy)
        {
            return "Data Source=" + service_name + ";user id=" + userid + mmyy + ";password=" + userid + mmyy;
        }

        public string Maincode(string sql)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("..\\..\\..\\xml\\maincode.xml");
                XmlNodeList nodeLst = doc.GetElementsByTagName(sql);
                return nodeLst.Item(0).InnerText;
            }
            catch { }
            return "";
        }
        public static string mySqlConnection()
        {
            string host = "127.0.0.1";
            int port = 3306;
            string database = "toannv_qlpk";
            string username = "tgmedisoft";
            string password = "huyethocxHd!23";
            string s_Conn = "Server=" + host + ";Database=" + database
                + ";port=" + port + ";User Id=" + username + ";password=" + password;
            return s_Conn;
            //return GetDBConnection(host, port, database, username, password);
        }
        public static MySqlConnection GetDBConnection(string host, int port, string database, string username, string password)
        {
            // Connection String.
            String connString = "Server=" + host + ";Database=" + database
                + ";port=" + port + ";User Id=" + username + ";password=" + password;

            MySqlConnection conn = new MySqlConnection(connString);

            return conn;
        }

        public string ConStr { get { return sConn; } }

        public string Madau
        {
            get
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("..\\..\\..\\xml\\maincode.xml");
                XmlNodeList nodeLst = doc.GetElementsByTagName("Madau");
                return nodeLst.Item(0).InnerText;
            }
        }

        #region xxx
        public string Host
        {
            get
            {
                return "203.162.56.241";
            }
        }

        public string User
        {
            get
            {
                return Mabv.Substring(0, 3);
            }
        }
        public string Mabv
        {
            get
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("..\\..\\..\\xml\\maincode.xml");
                XmlNodeList nodeLst = doc.GetElementsByTagName("Mabv");
                return (nodeLst.Item(0).InnerText == "") ? "701.1.01" : nodeLst.Item(0).InnerText;
            }
        }

        public string Pass
        {
            get
            {
                return User + "09t039i3e7066n8" + User;
            }
        }

        public string Dir
        {
            get
            {
                return "";
            }
        }
        #endregion

        public void writeXml(string tenfile, string cot, string s)
        {
            ds = new DataSet();
            try
            {
                ds.ReadXml("..\\..\\..\\xml\\" + tenfile.Replace(".xml", "") + ".xml");
                ds.Tables[0].Rows[0][cot] = s;
            }
            catch
            {
                DataColumn dc = new DataColumn();
                dc.ColumnName = cot;
                dc.DataType = Type.GetType("System.String");
                ds.Tables[0].Columns.Add(dc);
                ds.Tables[0].Rows[0][cot] = s;
            }
            ds.WriteXml("..\\..\\..\\xml\\" + tenfile + ".xml");
        }
        public DataSet get_data(string asql)
        {
            DataSet ds1 = new DataSet();
            try
            {
                if (con != null)
                {
                    con.Close(); con.Dispose();
                }
                ds1 = new DataSet();
                con = new OracleConnection(sConn);
                con.Open();
                cmd = new OracleCommand(asql, con);
                cmd.CommandType = CommandType.Text;
                dest = new OracleDataAdapter(cmd);
                dest.Fill(ds1);
                cmd.Dispose();
                con.Close(); con.Dispose();
                if (ds1.Tables.Count == 0) { ds1 = null; }
            }
            catch (OracleException ex)
            {
                upd_error(ex.Message.ToString().Trim() + " - asql: " + asql,sComputer, "?");
                ds1 = null;
            }
            return ds1;
        }
        public void upd_error(string m_message, string m_computer, string m_table)
        {
            con.Close(); con.Dispose();
            string sql1 = "insert into error(message,computer,tables,ngayud) values (:m_message,:m_computer,:m_table,sysdate)";
            con = new OracleConnection(sConn);
            try
            {
                con.Open();
                cmd = new OracleCommand(sql1, con);
                cmd.CommandType = CommandType.Text;

                cmd.Parameters.Add("m_message", OracleType.VarChar).Value = m_message;
                cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                cmd.Parameters.Add("m_table", OracleType.VarChar, 20).Value = m_table;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            catch { }
            finally
            {
                con.Close(); con.Dispose();
            }
        }

        public DataRow getrowbyid(DataTable dt, string exp)
        {
            try
            {
                DataRow[] r = dt.Select(exp);
                return r[0];
            }
            catch { return null; }
        }
        public DataRow getrowbyid(DataTable dt, string exp, string sort)
        {
            try
            {
                DataRow[] r = dt.Select(exp, sort);
                return r[0];
            }
            catch { return null; }
        }
        public string mmyy(string ngay)
        {
            try
            {
                return (ngay.Length == 4) ? ngay : ngay.Substring(3, 2) + ngay.Substring(8, 2);
            }
            catch { return ""; }
        }
        public bool execute_data(string sql)
        {
            try
            {
                con = new OracleConnection(sConn);
                con.Open();
                cmd = new OracleCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close(); con.Dispose();
                return true;
            }
            catch (OracleException ex)
            {
                upd_error(ex.Message.ToString().Trim() + "-SQL: " + sql, sComputer, "?");
                return false;
            }
        }
        public DataSet get_data_mmyy(string str, string tu, string den)
        {
            DataSet tmp = new DataSet();

            DateTime dt1 = StringToDate(tu);
            DateTime dt2 = StringToDate(den);
            int y1 = dt1.Year, m1 = dt1.Month;
            int y2 = dt2.Year, m2 = dt2.Month;
            int itu, iden;
            string mmyy = "";
            bool be = true;
            for (int i = y1; i <= y2; i++)
            {
                itu = (i == y1) ? m1 : 1;
                iden = (i == y2) ? m2 : 12;
                for (int j = itu; j <= iden; j++)
                {
                    mmyy = j.ToString().PadLeft(2, '0') + i.ToString().Substring(2, 2);
                    if (bMmyy(mmyy))
                    {
                        sql = str.Replace("xxx", user + mmyy);
                        if (be || tmp == null)
                        {
                            tmp = get_data(sql);
                            be = false;
                        }
                        else tmp.Merge(get_data(sql));
                    }
                }
            }
            return tmp;
        }
        public System.DateTime StringToDate(string s)
        {
            string[] format ={ "dd/MM/yyyy" };
            return System.DateTime.ParseExact(s.Substring(0, 10), format, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }
        public string ngayhienhanh_server
        {
            get
            {
                return get_data("select to_char(sysdate,'dd/mm/yyyy hh24:mi') as ngay from dual").Tables[0].Rows[0]["ngay"].ToString();
            }
        }
        public bool bMmyy(string m_mmyy)
        {
            return get_data("select * from m_table where mmyy='" + ((m_mmyy.Trim().Length == 4) ? m_mmyy : mmyy(m_mmyy)) + "'").Tables[0].Rows.Count > 0;
        }
        public string Ngaygio_hienhanh
        {
            get { return DateTime.Now.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Year.ToString() + " " + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0'); }
        }
        public System.DateTime StringToDateTime(string s)
        {
            string[] format1 ={ "dd/MM/yyyy" }, format2 ={ "dd/MM/yyyy HH:mm" };
            return System.DateTime.ParseExact(s.ToString(), (s.Length == 10) ? format1 : format2, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }
        public System.DateTime StringToDateTime(string s, string f)
        {
            string[] format ={ f };
            return System.DateTime.ParseExact(s.ToString(), format, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }


        public bool upd_hddt_tonghop(decimal m_id, string m_sohoso, decimal m_sotien, int m_loaivp, int m_quyenso, int m_sobienlai,
            int m_idthungan, string m_hotenthungan, string m_makhachhang, string m_tenkhachhang, string m_namsinh,
            int m_phai, string m_diachi, decimal m_miengiam, string m_ngayhd, string m_mmyy)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            sql = "update hddt.tonghop set sohoso=:m_sohoso,quyenso=:m_quyenso,sobienlai=:m_sobienlai,sotien=:m_sotien,idthungan=:m_idthungan,hotenthungan=:m_hotenthungan,makhachhang=:m_makhachhang,tenkhachhang=:m_tenkhachhang,namsinh=:m_namsinh,phai=:m_phai,diachi=:m_diachi,miengiam=:m_miengiam,ngayhd=to_date(:m_ngayhd,'dd/mm/yyyy hh24:mi'),mmyy=:m_mmyy";
            sql += " where id=:m_id and loaivp=:m_loaivp";
            con = new OracleConnection(sConn);
            try
            {
                con.Open();
                cmd = new OracleCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_sohoso", OracleType.VarChar, 500).Value = m_sohoso;
                cmd.Parameters.Add("m_loaivp", OracleType.Number).Value = m_loaivp;
                cmd.Parameters.Add("m_quyenso", OracleType.Number).Value = m_quyenso;
                cmd.Parameters.Add("m_sobienlai", OracleType.Number).Value = m_sobienlai;
                cmd.Parameters.Add("m_sotien", OracleType.Number).Value = m_sotien;
                cmd.Parameters.Add("m_idthungan", OracleType.Number).Value = m_idthungan;
                cmd.Parameters.Add("m_hotenthungan", OracleType.VarChar, 500).Value = m_hotenthungan;
                cmd.Parameters.Add("m_makhachhang", OracleType.VarChar, 500).Value = m_makhachhang;
                cmd.Parameters.Add("m_tenkhachhang", OracleType.VarChar, 500).Value = m_tenkhachhang;
                cmd.Parameters.Add("m_namsinh", OracleType.VarChar, 500).Value = m_namsinh;
                cmd.Parameters.Add("m_phai", OracleType.Number).Value = m_phai;
                cmd.Parameters.Add("m_diachi", OracleType.VarChar, 500).Value = m_diachi;
                cmd.Parameters.Add("m_miengiam", OracleType.Number).Value = m_miengiam;
                cmd.Parameters.Add("m_ngayhd", OracleType.VarChar, 16).Value = m_ngayhd;
                cmd.Parameters.Add("m_mmyy", OracleType.VarChar, 500).Value = m_mmyy;
                cmd.Parameters.Add("m_id", OracleType.Number).Value = m_id;
                int irec = cmd.ExecuteNonQuery();
                cmd.Dispose();
                if (irec == 0)
                {
                    sql = "INSERT INTO hddt.tonghop(id,sohoso,loaivp,quyenso,sobienlai,sotien,idthungan,hotenthungan,makhachhang,tenkhachhang,namsinh,phai,diachi,miengiam,ngayhd,mmyy)";
                    sql += " values (:m_id,:m_sohoso,:m_loaivp,:m_quyenso,:m_sobienlai,:m_sotien,:m_idthungan,:m_hotenthungan,:m_makhachhang,:m_tenkhachhang,:m_namsinh,:m_phai,:m_diachi,:m_miengiam,to_date(:m_ngayhd,'dd/mm/yyyy hh24:mi'),:m_mmyy)";

                    cmd = new OracleCommand(sql, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add("m_id", OracleType.Number).Value = m_id;
                    cmd.Parameters.Add("m_sohoso", OracleType.VarChar, 500).Value = m_sohoso;
                    cmd.Parameters.Add("m_loaivp", OracleType.Number).Value = m_loaivp;
                    cmd.Parameters.Add("m_quyenso", OracleType.Number).Value = m_quyenso;
                    cmd.Parameters.Add("m_sobienlai", OracleType.Number).Value = m_sobienlai;
                    cmd.Parameters.Add("m_sotien", OracleType.Number).Value = m_sotien;
                    cmd.Parameters.Add("m_idthungan", OracleType.Number).Value = m_idthungan;
                    cmd.Parameters.Add("m_hotenthungan", OracleType.VarChar, 500).Value = m_hotenthungan;
                    cmd.Parameters.Add("m_makhachhang", OracleType.VarChar, 500).Value = m_makhachhang;
                    cmd.Parameters.Add("m_tenkhachhang", OracleType.VarChar, 500).Value = m_tenkhachhang;
                    cmd.Parameters.Add("m_namsinh", OracleType.VarChar, 500).Value = m_namsinh;
                    cmd.Parameters.Add("m_phai", OracleType.Number).Value = m_phai;
                    cmd.Parameters.Add("m_diachi", OracleType.VarChar, 500).Value = m_diachi;
                    cmd.Parameters.Add("m_miengiam", OracleType.Number).Value = m_miengiam;
                    cmd.Parameters.Add("m_ngayhd", OracleType.VarChar, 16).Value = m_ngayhd;
                    cmd.Parameters.Add("m_mmyy", OracleType.VarChar, 500).Value = m_mmyy;
                    irec = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
            }
            catch (OracleException ex)
            {
                upd_error(ex.Message,sComputer, "tonghop");
                return false;
            }
            finally
            {
                con.Close(); con.Dispose();
            }
            return true;
        }

        public bool upd_hddt_chitiet(decimal m_id, int m_stt, string m_ma, string m_tenhang,
            string m_donvitinh, decimal m_soluong, decimal m_dongia, decimal m_thanhtien, decimal m_TYLEBH,
                decimal m_MUCBHTRA, decimal m_BHXHTRA, string m_nhom, int m_loaivp, string m_mmyy)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            sql = "update hddt.chitiet set ma=:m_ma,tenhang=:m_tenhang,donvitinh=:m_donvitinh,";
            sql += " soluong=:m_soluong,dongia=:m_dongia,thanhtien=:m_thanhtien,TYLEBH=:m_TYLEBH,MUCBHTRA=:m_MUCBHTRA,BHXHTRA=:m_BHXHTRA,";
            sql += " nhom=:m_nhom";
            sql += " where id=:m_id and stt=:m_stt and loaivp=:m_loaivp and mmyy=:m_mmyy";
            con = new OracleConnection(sConn);
            try
            {
                con.Open();
                //cmd = new NpgsqlCommand(sql, con);
                cmd = new OracleCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_stt", OracleType.Number).Value = m_stt;
                cmd.Parameters.Add("m_ma", OracleType.VarChar, 500).Value = m_ma;
                cmd.Parameters.Add("m_tenhang", OracleType.VarChar, 500).Value = m_tenhang;
                cmd.Parameters.Add("m_donvitinh", OracleType.VarChar, 500).Value = m_donvitinh;
                cmd.Parameters.Add("m_soluong", OracleType.Number).Value = m_soluong;
                cmd.Parameters.Add("m_dongia", OracleType.Number).Value = m_dongia;
                cmd.Parameters.Add("m_thanhtien", OracleType.Number).Value = m_thanhtien;
                cmd.Parameters.Add("m_TYLEBH", OracleType.Number).Value = m_TYLEBH;
                cmd.Parameters.Add("m_MUCBHTRA", OracleType.Number).Value = m_MUCBHTRA;
                cmd.Parameters.Add("m_BHXHTRA", OracleType.Number).Value = m_BHXHTRA;
                cmd.Parameters.Add("m_nhom", OracleType.VarChar, 500).Value = m_nhom;
                cmd.Parameters.Add("m_mmyy", OracleType.VarChar, 500).Value = m_mmyy;
                cmd.Parameters.Add("m_loaivp", OracleType.Number).Value = m_loaivp;
                cmd.Parameters.Add("m_id", OracleType.Number).Value = m_id;

                int irec = cmd.ExecuteNonQuery();
                cmd.Dispose();
                if (irec == 0)
                {
                    sql = "insert into hddt.chitiet (id,stt,ma,tenhang,donvitinh,soluong,dongia,thanhtien,TYLEBH,MUCBHTRA,BHXHTRA,nhom,loaivp,mmyy)";
                    sql += " values (:m_id,:m_stt,:m_ma,:m_tenhang,:m_donvitinh,:m_soluong,:m_dongia,:m_thanhtien,:m_TYLEBH,:m_MUCBHTRA,:m_BHXHTRA,:m_nhom,:m_loaivp,:m_mmyy)";

                    cmd = new OracleCommand(sql, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add("m_id", OracleType.Number).Value = m_id;
                    cmd.Parameters.Add("m_stt", OracleType.Number).Value = m_stt;
                    cmd.Parameters.Add("m_ma", OracleType.VarChar, 500).Value = m_ma;
                    cmd.Parameters.Add("m_tenhang", OracleType.VarChar, 500).Value = m_tenhang;
                    cmd.Parameters.Add("m_donvitinh", OracleType.VarChar, 500).Value = m_donvitinh;
                    cmd.Parameters.Add("m_soluong", OracleType.Number).Value = m_soluong;
                    cmd.Parameters.Add("m_dongia", OracleType.Number).Value = m_dongia;
                    cmd.Parameters.Add("m_thanhtien", OracleType.Number).Value = m_thanhtien;
                    cmd.Parameters.Add("m_TYLEBH", OracleType.Number).Value = m_TYLEBH;
                    cmd.Parameters.Add("m_MUCBHTRA", OracleType.Number).Value = m_MUCBHTRA;
                    cmd.Parameters.Add("m_BHXHTRA", OracleType.Number).Value = m_BHXHTRA;
                    cmd.Parameters.Add("m_nhom", OracleType.VarChar, 500).Value = m_nhom;
                    cmd.Parameters.Add("m_mmyy", OracleType.VarChar, 500).Value = m_mmyy;
                    cmd.Parameters.Add("m_loaivp", OracleType.Number).Value = m_loaivp;

                    irec = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
            }
            catch (OracleException ex)
            {
                upd_error(ex.Message,sComputer, "chitiet");
                return false;
            }
            finally
            {
                con.Close(); con.Dispose();
            }
            return true;
        }

        public bool upd_hddt_tonghop_mysql(decimal m_id, string m_sohoso, decimal m_sotien, int m_loaivp, int m_quyenso, int m_sobienlai,
           int m_idthungan, string m_hotenthungan, string m_makhachhang, string m_tenkhachhang, string m_namsinh,
           int m_phai, string m_diachi, decimal m_miengiam, string m_ngayhd, string m_mmyy)
        {
            if (con_mysql != null)
            {
                con_mysql.Close(); 
                con_mysql.Dispose();
            }
            sql = "update tb_master set tenkhachhang=:m_tenkhachhang,diachi=:m_diachi,tongtientt=:m_sotien,nguoithu=:m_idthungan,hotennguoithu=:m_hotenthungan,ngaylap=STR_TO_DATE(:m_ngayhd, '%d/%m/%Y %H:%i')";
            sql += " where sohoso=:m_sohoso";
            con_mysql = new MySqlConnection(sConn_mysql);
            try
            {
                con.Open();
                cmd_mysql = new MySqlCommand(sql, con_mysql);
                cmd_mysql.CommandType = CommandType.Text;
                cmd_mysql.Parameters.Add("m_sotien", MySqlDbType.Double).Value = m_sotien;
                cmd_mysql.Parameters.Add("m_idthungan", MySqlDbType.VarChar).Value = m_sotien;
                cmd_mysql.Parameters.Add("m_hotenthungan", MySqlDbType.VarChar,255).Value = m_hotenthungan;
                cmd_mysql.Parameters.Add("m_tenkhachhang", MySqlDbType.VarChar, 255).Value = m_tenkhachhang;
                cmd_mysql.Parameters.Add("m_diachi", MySqlDbType.Text, 500).Value = m_diachi;
                cmd_mysql.Parameters.Add("m_ngayhd", MySqlDbType.String, 16).Value = m_ngayhd;
                cmd_mysql.Parameters.Add("m_sohoso", MySqlDbType.VarChar, 255).Value = m_sohoso;
                //cmd_mysql.Parameters.Add("m_id", OracleType.Number).Value = m_id;
                int irec = cmd_mysql.ExecuteNonQuery();
                cmd_mysql.Dispose();
                if (irec == 0)
                {
                    sql = "INSERT INTO tb_master(sohoso,tenkhachhang,diachi,tongtientt,ngaylap)";
                    sql += " values (:m_sohoso,:m_tenkhachhang,:m_diachi,:m_sotien,STR_TO_DATE(:m_ngayhd, '%d/%m/%Y %H:%i'))";

                    cmd_mysql = new MySqlCommand(sql, con_mysql);
                    cmd_mysql.CommandType = CommandType.Text;
                    //cmd.Parameters.Add("m_id", MySqlDbType.Number).Value = m_id;
                    cmd_mysql.Parameters.Add("m_sohoso", MySqlDbType.VarChar, 255).Value = m_sohoso;
                    cmd_mysql.Parameters.Add("m_sotien", MySqlDbType.Double).Value = m_sotien;
                    cmd_mysql.Parameters.Add("m_idthungan", MySqlDbType.VarChar).Value = m_sotien;
                    cmd_mysql.Parameters.Add("m_hotenthungan", MySqlDbType.VarChar, 255).Value = m_hotenthungan;
                    cmd_mysql.Parameters.Add("m_tenkhachhang", MySqlDbType.VarChar, 255).Value = m_tenkhachhang;
                    cmd_mysql.Parameters.Add("m_diachi", MySqlDbType.Text, 500).Value = m_diachi;
                    cmd_mysql.Parameters.Add("m_ngayhd", MySqlDbType.String, 16).Value = m_ngayhd;
                    irec = cmd.ExecuteNonQuery();
                    cmd_mysql.Dispose();
                }
            }
            catch 
            {
                //upd_error(ex.Message, sComputer, "tonghop");
                return false;
            }
            finally
            {
                con_mysql.Close(); con_mysql.Dispose();
            }
            return true;
        }

        public bool upd_hddt_chitiet_mysql(string m_sohoso, string m_ma, string m_loaidv, string m_tenhang, double m_dongia, string m_donvitinh,
            double m_soluong, double m_thanhtien, double m_TYLEBH,double m_MUCBHTRA,double m_BHXHTRA)
        {
            if (con_mysql != null)
            {
                con_mysql.Close(); con_mysql.Dispose();
            }
            sql = "update tb_detail set MAHANG=:m_ma,LOAIDV=:m_loaidv,tenhang=:m_tenhang,dongia=:m_dongia,donvitinh=:m_donvitinh,";
            sql += " soluong=:m_soluong,thanhtien=:m_thanhtien,TONGTIEN=:m_thanhtien,TYLEBH=:m_TYLEBH,MUCBHTRA=:m_MUCBHTRA,BHXHTRA=:m_BHXHTRA";
            sql += " where SOHOSO=:m_sohoso";
            con_mysql = new MySqlConnection(sConn_mysql);
            try
            {
                con_mysql.Open();
                cmd_mysql = new MySqlCommand(sql, con_mysql);
                //cmd_mysql = new OracleCommand(sql, con);
                cmd_mysql.CommandType = CommandType.Text;
                cmd_mysql.Parameters.Add("m_ma", MySqlDbType.VarChar, 50).Value = m_ma;
                cmd_mysql.Parameters.Add("m_loaidv", MySqlDbType.VarChar, 50).Value = m_loaidv;
                cmd_mysql.Parameters.Add("m_tenhang", MySqlDbType.VarChar, 155).Value = m_tenhang;
                cmd_mysql.Parameters.Add("m_donvitinh", MySqlDbType.VarChar, 50).Value = m_donvitinh;
                cmd_mysql.Parameters.Add("m_soluong", MySqlDbType.Double).Value = m_soluong;
                cmd_mysql.Parameters.Add("m_dongia", MySqlDbType.Double).Value = m_dongia;
                cmd_mysql.Parameters.Add("m_thanhtien", MySqlDbType.Double).Value = m_thanhtien;
                cmd_mysql.Parameters.Add("m_TYLEBH", MySqlDbType.Double).Value = m_TYLEBH;
                cmd_mysql.Parameters.Add("m_MUCBHTRA", MySqlDbType.Double).Value = m_MUCBHTRA;
                cmd_mysql.Parameters.Add("m_BHXHTRA", MySqlDbType.Double).Value = m_BHXHTRA;
                cmd_mysql.Parameters.Add("m_sohoso", MySqlDbType.VarChar, 255).Value = m_sohoso;

                int irec = cmd_mysql.ExecuteNonQuery();
                cmd_mysql.Dispose();
                if (irec == 0)
                {
                    sql = "insert into tb_detail (SOHOSO,MAHANG,LOAIDV,tenhang,donvitinh,soluong,dongia,thanhtien,TYLEBH,MUCBHTRA,BHXHTRA)";
                    sql += " values (:m_sohoso,:m_ma,:m_loaidv,:m_tenhang,:m_donvitinh,:m_soluong,:m_dongia,:m_thanhtien,:m_TYLEBH,:m_MUCBHTRA,:m_BHXHTRA)";

                    cmd_mysql = new MySqlCommand(sql, con_mysql);
                    cmd_mysql.CommandType = CommandType.Text;
                    cmd_mysql.Parameters.Add("m_sohoso", MySqlDbType.VarChar, 255).Value = m_sohoso;
                    cmd_mysql.Parameters.Add("m_ma", MySqlDbType.VarChar, 50).Value = m_ma;
                    cmd_mysql.Parameters.Add("m_loaidv", MySqlDbType.VarChar, 50).Value = m_loaidv;
                    cmd_mysql.Parameters.Add("m_tenhang", MySqlDbType.VarChar, 155).Value = m_tenhang;
                    cmd_mysql.Parameters.Add("m_donvitinh", MySqlDbType.VarChar, 50).Value = m_donvitinh;
                    cmd_mysql.Parameters.Add("m_soluong", MySqlDbType.Double).Value = m_soluong;
                    cmd_mysql.Parameters.Add("m_dongia", MySqlDbType.Double).Value = m_dongia;
                    cmd_mysql.Parameters.Add("m_thanhtien", MySqlDbType.Double).Value = m_thanhtien;
                    cmd_mysql.Parameters.Add("m_TYLEBH", MySqlDbType.Double).Value = m_TYLEBH;
                    cmd_mysql.Parameters.Add("m_MUCBHTRA", MySqlDbType.Double).Value = m_MUCBHTRA;
                    cmd_mysql.Parameters.Add("m_BHXHTRA", MySqlDbType.Double).Value = m_BHXHTRA;

                    irec = cmd_mysql.ExecuteNonQuery();
                    cmd_mysql.Dispose();
                }
            }
            catch (MySqlException ex)
            {
                return false;
            }
            finally
            {
                con_mysql.Close(); con_mysql.Dispose();
            }
            return true;
        }
    }
}