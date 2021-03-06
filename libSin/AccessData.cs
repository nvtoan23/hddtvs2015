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

namespace libSin_Postgres
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
                //doc.Load("..//..//..//xml//maincode.xml");
                doc.Load("..//..//xml//maincode.xml");
                XmlNodeList nodeLst = doc.GetElementsByTagName(sql);
                return nodeLst.Item(0).InnerText;
            }
            catch
            {
                ds = new DataSet();
                ds.ReadXml("..//..//xml//maincode.xml");
                DataColumn dc = new DataColumn();
                dc.ColumnName = sql;
                dc.DataType = Type.GetType("System.String");
                ds.Tables[0].Columns.Add(dc);
                ds.Tables[0].Rows[0][sql] = "";
                ds.WriteXml("..//..//xml//maincode.xml");
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
        public void upd_error(string m_ngay, string m_message, string m_computer, string m_table)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            if (!bMmyy(mmyy(m_ngay))) return;
            sql = "insert into " + user + mmyy(m_ngay) + ".error(message,computer,tables) values (:m_message,:m_computer,:m_table)";
            con = new NpgsqlConnection(sConn);
            try
            {
                con.Open();

                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_message", NpgsqlDbType.Text).Value = m_message;
                cmd.Parameters.Add("m_computer", NpgsqlDbType.Varchar, 20).Value = m_computer;
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

            //sql = sql.Replace("medibv.", user + ".");
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
                //MessageBox.Show("Lỗi: " + sql + " - " + ex.Message.ToString());
                return false;
            }
        }
        public DataSet get_data_mmyy(string str, string tu, string den)
        {
            DataSet tmp = null;
            DateTime dt1 = StringToDate(tu);
            DateTime dt2 = StringToDate(den);
            int y1 = dt1.Year, m1 = dt1.Month;
            int y2 = dt2.Year, m2 = dt2.Month;
            int itu, iden;
            string mmyy = "";
            bool be = true;
            Npgsql.NpgsqlConnection connct = new NpgsqlConnection(ConStr);
            connct.Open();
            for (int i = y1; i <= y2; i++)
            {
                itu = (i == y1) ? m1 : 1;
                iden = (i == y2) ? m2 : 12;
                for (int j = itu; j <= iden; j++)
                {
                    mmyy = j.ToString().PadLeft(2, '0') + i.ToString().Substring(2, 2);
                    if (bMmyy(mmyy))
                    {
                        sql = str.Replace("xxx", "medibv" + mmyy);
                        sql = str.Replace("medibvmmyy", "medibv" + mmyy);
                        using (Npgsql.NpgsqlCommand cmm = new NpgsqlCommand(sql, connct))
                        {
                            //   cmm.Connection.Open();
                            Npgsql.NpgsqlDataReader drd = null;
                            try
                            {
                                drd = cmm.ExecuteReader();
                                if (tmp == null)
                                    tmp = new DataSet();
                                if (tmp.Tables.Count == 0 && drd.FieldCount > 0)
                                    tmp.Tables.Add("Table");
                                if (tmp.Tables.Count > 0)
                                {
                                    for (int ia = 0; ia < drd.FieldCount; ia++)
                                    {
                                        if (!tmp.Tables[0].Columns.Contains(drd.GetName(ia)))
                                            tmp.Tables[0].Columns.Add(drd.GetName(ia), drd.GetFieldType(ia));
                                    }
                                    while (drd.Read())
                                    {
                                        DataRow ndtr = tmp.Tables[0].NewRow();
                                        for (int ie = 0; ie < drd.FieldCount; ie++)
                                            ndtr[drd.GetName(ie)] = drd[ie];
                                        tmp.Tables[0].Rows.Add(ndtr);
                                    }
                                }
                            }
                            catch
                            {
                            }
                            finally
                            {
                                if (drd != null)
                                {
                                    drd.Close();
                                    drd.Dispose();
                                }
                            }
                        }
                    }
                }
            }
            connct.Close();
            return tmp;
        }
        public System.DateTime StringToDate(string s)
        {
            s = (s == "" || s == null) ? ngayhienhanh_server.Substring(0, 10) : s;
            string[] aa = s.Split('/');
            s = aa[0].ToString().PadLeft(2, '0') + "/" + aa[1].ToString().PadLeft(2, '0') + "/" + aa[2].ToString();
            string[] format = { "dd/MM/yyyy" };
            return System.DateTime.ParseExact(s.Substring(0, 10), format, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }
        public string ngayhienhanh_server
        {
            get
            {
                return get_data("select to_char(now(),'dd/mm/yyyy hh24:mi')").Tables[0].Rows[0][0].ToString();
            }
        }
        public string ConStr
        {
            get
            {
                return sConn;
            }
        }
        public bool bMmyy(string m_mmyy)
        {
            return get_data("select * from information_schema.schemata where schema_name = 'medibv" + ((m_mmyy.Trim().Length == 4) ? m_mmyy : mmyy(m_mmyy)) + "'").Tables[0].Rows.Count > 0;
        }
        public string Ngaygio_hienhanh
        {
            get { return ngayhienhanh_server; }//DateTime.Now.Day.ToString().PadLeft(2,'0')+"/"+DateTime.Now.Month.ToString().PadLeft(2,'0')+"/"+DateTime.Now.Year.ToString()+" "+DateTime.Now.Hour.ToString().PadLeft(2,'0')+":"+DateTime.Now.Minute.ToString().PadLeft(2,'0');}
        }
        public System.DateTime StringToDateTime(string s)
        {
            if (s.Length >= 16) s = s.Substring(0, 16);
            else if (s.Length > 10) s = s.Substring(0, 10);
            string[] format1 ={ "dd/MM/yyyy" }, format2 ={ "dd/MM/yyyy HH:mm" };
            return System.DateTime.ParseExact(s.ToString(), (s.Length == 10) ? format1 : format2, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }

        public System.DateTime StringToDateTime(string s, string f)
        {
            string[] format ={ f };
            return System.DateTime.ParseExact(s.ToString(), format, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
        }

        public bool upd_hddt_tonghop(decimal m_id,string m_sohoso,decimal m_sotien, int m_loaivp,int m_quyenso,int m_sobienlai,
            int m_idthungan,string m_hotenthungan,string m_makhachhang,string m_tenkhachhang,string m_namsinh,
            int m_phai,string m_diachi,decimal m_miengiam,string m_ngayhd,string m_mmyy)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            sql = "update hddt.tonghop set sohoso=:m_sohoso,quyenso=:m_quyenso,sobienlai=:m_sobienlai,sotien=:m_sotien,idthungan=:m_idthungan,hotenthungan=:m_hotenthungan,makhachhang=:m_makhachhang,tenkhachhang=:m_tenkhachhang,namsinh=:m_namsinh,phai=:m_phai,diachi=:m_diachi,miengiam=:m_miengiam,ngayhd=to_date(:m_ngayhd,'dd/mm/yyyy hh24:mi'),mmyy=:m_mmyy";
            sql += " where id=:m_id and loaivp=:m_loaivp";
            con = new NpgsqlConnection(sConn);
            try
            {
                con.Open();
                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_sohoso", NpgsqlDbType.Varchar).Value = m_sohoso;
                cmd.Parameters.Add("m_loaivp", NpgsqlDbType.Numeric).Value = m_loaivp;
                cmd.Parameters.Add("m_quyenso", NpgsqlDbType.Numeric).Value = m_quyenso;
                cmd.Parameters.Add("m_sobienlai", NpgsqlDbType.Numeric).Value = m_sobienlai;
                cmd.Parameters.Add("m_sotien", NpgsqlDbType.Numeric).Value = m_sotien;
                cmd.Parameters.Add("m_idthungan", NpgsqlDbType.Numeric).Value = m_idthungan;
                cmd.Parameters.Add("m_hotenthungan", NpgsqlDbType.Text).Value = m_hotenthungan;
                cmd.Parameters.Add("m_makhachhang", NpgsqlDbType.Varchar).Value = m_makhachhang;
                cmd.Parameters.Add("m_tenkhachhang", NpgsqlDbType.Text).Value = m_tenkhachhang;
                cmd.Parameters.Add("m_namsinh", NpgsqlDbType.Varchar).Value = m_namsinh;
                cmd.Parameters.Add("m_phai", NpgsqlDbType.Numeric).Value = m_phai;
                cmd.Parameters.Add("m_diachi", NpgsqlDbType.Text).Value = m_diachi;
                cmd.Parameters.Add("m_miengiam", NpgsqlDbType.Numeric).Value = m_miengiam;
                cmd.Parameters.Add("m_ngayhd", NpgsqlDbType.Varchar, 16).Value = m_ngayhd;
                cmd.Parameters.Add("m_mmyy", NpgsqlDbType.Varchar).Value = m_mmyy;
                cmd.Parameters.Add("m_id", NpgsqlDbType.Numeric).Value = m_id;
                int irec = cmd.ExecuteNonQuery();
                cmd.Dispose();
                if (irec == 0)
                {
                    sql = "INSERT INTO hddt.tonghop(id,sohoso,loaivp,quyenso,sobienlai,sotien,idthungan,hotenthungan,makhachhang,tenkhachhang,namsinh,phai,diachi,miengiam,ngayhd,mmyy)";
                    sql += " values (:m_id,:m_sohoso,:m_loaivp,:m_quyenso,:m_sobienlai,:m_sotien,:m_idthungan,:m_hotenthungan,:m_makhachhang,:m_tenkhachhang,:m_namsinh,:m_phai,:m_diachi,:m_miengiam,to_date(:m_ngayhd,'dd/mm/yyyy hh24:mi'),:m_mmyy)";

                    cmd = new NpgsqlCommand(sql, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add("m_id", NpgsqlDbType.Numeric).Value = m_id;
                    cmd.Parameters.Add("m_sohoso", NpgsqlDbType.Varchar).Value = m_sohoso;
                    cmd.Parameters.Add("m_loaivp", NpgsqlDbType.Numeric).Value = m_loaivp;
                    cmd.Parameters.Add("m_quyenso", NpgsqlDbType.Numeric).Value = m_quyenso;
                    cmd.Parameters.Add("m_sobienlai", NpgsqlDbType.Numeric).Value = m_sobienlai;
                    cmd.Parameters.Add("m_sotien", NpgsqlDbType.Numeric).Value = m_sotien;
                    cmd.Parameters.Add("m_idthungan", NpgsqlDbType.Numeric).Value = m_idthungan;
                    cmd.Parameters.Add("m_hotenthungan", NpgsqlDbType.Text).Value = m_hotenthungan;
                    cmd.Parameters.Add("m_makhachhang", NpgsqlDbType.Varchar).Value = m_makhachhang;
                    cmd.Parameters.Add("m_tenkhachhang", NpgsqlDbType.Text).Value = m_tenkhachhang;
                    cmd.Parameters.Add("m_namsinh", NpgsqlDbType.Varchar).Value = m_namsinh;
                    cmd.Parameters.Add("m_phai", NpgsqlDbType.Numeric).Value = m_phai;
                    cmd.Parameters.Add("m_diachi", NpgsqlDbType.Text).Value = m_diachi;
                    cmd.Parameters.Add("m_miengiam", NpgsqlDbType.Numeric).Value = m_miengiam;
                    //cmd.Parameters.Add("m_ngayhd", NpgsqlDbType.Timestamp).Value = StringToDateTime(m_ngayhd);
                    cmd.Parameters.Add("m_ngayhd", NpgsqlDbType.Varchar, 16).Value = m_ngayhd;
                    cmd.Parameters.Add("m_mmyy", NpgsqlDbType.Varchar).Value = m_mmyy;
                    irec = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
            }
            catch (NpgsqlException ex)
            {
                upd_error(ex.Message, "tonghop");
                return false;
            }
            finally
            {
                con.Close(); con.Dispose();
            }
            return true;
        }

        public bool upd_hddt_chitiet(decimal m_id,int m_stt,string m_ma,string m_tenhang,
            string m_donvitinh,decimal m_soluong,decimal m_dongia,decimal m_thanhtien,decimal m_TYLEBH,
                decimal m_MUCBHTRA,decimal m_BHXHTRA,string m_nhom,int m_loaivp,string m_mmyy)
        {
            if (con != null)
            {
                con.Close(); con.Dispose();
            }
            sql = "update hddt.chitiet set ma=:m_ma,tenhang=:m_tenhang,donvitinh=:m_donvitinh,";
            sql +=" soluong=:m_soluong,dongia=:m_dongia,thanhtien=:m_thanhtien,TYLEBH=:m_TYLEBH,MUCBHTRA=:m_MUCBHTRA,BHXHTRA=:m_BHXHTRA,";
            sql +=" nhom=:m_nhom";
            sql += " where id=:m_id and stt=:m_stt and loaivp=:m_loaivp and mmyy=:m_mmyy";
            con = new NpgsqlConnection(sConn);
            try
            {
                con.Open();
                cmd = new NpgsqlCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add("m_stt", NpgsqlDbType.Numeric).Value = m_stt;
                cmd.Parameters.Add("m_ma", NpgsqlDbType.Varchar).Value = m_ma;
                cmd.Parameters.Add("m_tenhang", NpgsqlDbType.Text).Value = m_tenhang;
                cmd.Parameters.Add("m_donvitinh", NpgsqlDbType.Text).Value = m_donvitinh;
                cmd.Parameters.Add("m_soluong", NpgsqlDbType.Numeric).Value = m_soluong;
                cmd.Parameters.Add("m_dongia", NpgsqlDbType.Numeric).Value = m_dongia;
                cmd.Parameters.Add("m_thanhtien", NpgsqlDbType.Numeric).Value = m_thanhtien;
                cmd.Parameters.Add("m_TYLEBH", NpgsqlDbType.Numeric).Value = m_TYLEBH;
                cmd.Parameters.Add("m_MUCBHTRA", NpgsqlDbType.Numeric).Value = m_MUCBHTRA;
                cmd.Parameters.Add("m_BHXHTRA", NpgsqlDbType.Numeric).Value = m_BHXHTRA;
                cmd.Parameters.Add("m_nhom", NpgsqlDbType.Text).Value = m_nhom;
                cmd.Parameters.Add("m_mmyy", NpgsqlDbType.Varchar).Value = m_mmyy;
                cmd.Parameters.Add("m_loaivp", NpgsqlDbType.Numeric).Value = m_loaivp;
                cmd.Parameters.Add("m_id", NpgsqlDbType.Numeric).Value = m_id;

                int irec = cmd.ExecuteNonQuery();
                cmd.Dispose();
                if (irec == 0)
                {
                    sql = "insert into hddt.chitiet (id,stt,ma,tenhang,donvitinh,soluong,dongia,thanhtien,TYLEBH,MUCBHTRA,BHXHTRA,nhom,loaivp,mmyy)";
                    sql +=" values (:m_id,:m_stt,:m_ma,:m_tenhang,:m_donvitinh,:m_soluong,:m_dongia,:m_thanhtien,:m_TYLEBH,:m_MUCBHTRA,:m_BHXHTRA,:m_nhom,:m_loaivp,:m_mmyy)";

                    cmd = new NpgsqlCommand(sql, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add("m_id", NpgsqlDbType.Numeric).Value = m_id;
                    cmd.Parameters.Add("m_stt", NpgsqlDbType.Numeric).Value = m_stt;
                    cmd.Parameters.Add("m_ma", NpgsqlDbType.Varchar).Value = m_ma;
                    cmd.Parameters.Add("m_tenhang", NpgsqlDbType.Text).Value = m_tenhang;
                    cmd.Parameters.Add("m_donvitinh", NpgsqlDbType.Text).Value = m_donvitinh;
                    cmd.Parameters.Add("m_soluong", NpgsqlDbType.Numeric).Value = m_soluong;
                    cmd.Parameters.Add("m_dongia", NpgsqlDbType.Numeric).Value = m_dongia;
                    cmd.Parameters.Add("m_thanhtien", NpgsqlDbType.Numeric).Value = m_thanhtien;
                    cmd.Parameters.Add("m_TYLEBH", NpgsqlDbType.Numeric).Value = m_TYLEBH;
                    cmd.Parameters.Add("m_MUCBHTRA", NpgsqlDbType.Numeric).Value = m_MUCBHTRA;
                    cmd.Parameters.Add("m_BHXHTRA", NpgsqlDbType.Numeric).Value = m_BHXHTRA;
                    cmd.Parameters.Add("m_nhom", NpgsqlDbType.Text).Value = m_nhom;
                    cmd.Parameters.Add("m_mmyy", NpgsqlDbType.Varchar).Value = m_mmyy;
                    cmd.Parameters.Add("m_loaivp", NpgsqlDbType.Numeric).Value = m_loaivp;
                    
                    irec = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
            }
            catch (NpgsqlException ex)
            {
                upd_error(ex.Message, "chitiet");
                return false;
            }
            finally
            {
                con.Close(); con.Dispose();
            }
            return true;
        }
        
    }
}
