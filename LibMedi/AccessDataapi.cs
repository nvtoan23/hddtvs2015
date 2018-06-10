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
using System.Net.Http;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text;
using libHDDT;

namespace LibMedi
{
    public class AccessDataApi
    {
        private SCreenReader.ConvertFont convert_font = new SCreenReader.ConvertFont();
        public const string Msg = "Medisoft THIS", Normal = "Bình thường", Msg_di = "Dinh dưỡng";
        public const string links_userid = "links", links_pass = "link7155019s20", uxxx = "tuneig@mp", pxxx = "tuneig@mp", xxxxx = "¯Ò¡Ì©Î«³²°Ô£";
        public const int Tiepdon = 0, Khambenh = 1, Ngoaitru = 2, Phongluu = 3, Nhanbenh = 4, Khoa = 5, Phauthuthuat = 6, Vienphi = 7, Duoc = 8, Xetnghiem = 9, Sieuam = 10, Noisoi = 11, Xquang = 12, Le = 13, Khucxa = 14, Cls = 15, Taikham = 16, Pttt = 17;
        public const int giamdoc = 1, phogiamdoc = 2, truongkhoa = 3, nhanvien = 8, nghiviec = 9, ybs_cls = 10;
        public int iHaophi = 5;
        string sConn = "Data Source=MEDISOFT;user id=MEDIBV;password=MEDIBV";
        string sConn_mysql = "SERVER=127.0.0.1;PORT=3306;DATABASE=dbtgmedisoft;UID=tgmedisoft;PASSWORD=huyethocxHd!23";
        AccessDataAPI apidata = new AccessDataAPI();
        private string apiGetOracle = "http://127.0.0.1:81/api/orc/getDatabySql";
        private string apiGetMySQL = "http://127.0.0.1:81/api/mysql/getDatabySql";
        private string apiExcuteMySQL = "http://127.0.0.1:81/api/mysql/executeQuery";
        private string apiGetMsSql = "http://127.0.0.1:81/api/mssql/getDatabySql";
        private string apiExcuteMsSql = "http://127.0.0.1:81/api/mssql/executeQuery";
        public string ipApi = "127.0.0.1:81";
        [DllImport("winmm.dll")]
        private static extern bool PlaySound(string lpszName, int hModule, int dwFlags);
        private int iRownum = 1;
        private decimal _e = 0, _p = 0, _l = 0, _g = 0;

        private DataSet dsvp = null;
        private DataSet dsfield = null;
        private DataSet ds = null;

        OracleDataAdapter dest;
        OracleConnection con;
        OracleCommand cmd;

        MySqlConnection con_mysql;
        MySqlDataAdapter dest_mysql;
        MySqlCommand cmd_mysql;

        string sComputer = null;
        string m_hotenkdau, sql = "", userid = "medibvtest", service_name = "medisoft", b_sobienlaitamung = "", b_tatca_sobienlaitamung = "";
        decimal tc_tamung = 0;
        bool b_sovaovien, b_soluutru;
        
        DoiTuong _ctsdoituong;
        HttpClient client = new HttpClient();
        public const string ANA_TIT1 = "Thuốc";
        public const string ANA_TIT2 = "Máu";
        public const string ANA_TIT3 = "Xét nghiệm";
        public const string ANA_TIT4 = "Chẩn đoán hình ảnh";
        public const string ANA_TIT5 = "Thăm dò chức năng";
        public const string ANA_TIT6 = "Thủ thuật, phẩu thuật";
        public const string ANA_TIT7 = "Khám bệnh";
        public const string ANA_TIT8 = "Vật tư y tế";
        public const string ANA_TIT9 = "Giường";
        public const string ANA_TIT10 = "Giường dịch vụ";
        public const string ANA_TIT11 = "chi phí vận chuyển";
        public const string ANA_TIT12 = "Thẻ chăm sóc bệnh nhân";
        public const string ANA_TIT13 = "Dịch truyền,đạm";
        public const string ANA_TIT14 = "Dinh dưỡng";
        public const string ANA_TIT15 = "Thuốc khoa dược";

        private string tenfile;
        public AccessDataApi()
        {
            client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            if (Maincode("Con") != "")
            {
                sConn = Maincode("Con");
                ipApi = Maincode("Api");
            }
            sConn_mysql = mySqlConnection();
            sComputer = System.Environment.MachineName.Trim().ToUpper();
            userid = sConn.Substring(sConn.LastIndexOf("=") + 1).Trim();
            service_name = sConn.Substring(sConn.IndexOf("=") + 1, sConn.IndexOf(";") - 1 - sConn.IndexOf("=")).Trim();
            //ds = get_data("select rownum,computer from dmcomputer");
            //DataRow r = getrowbyid(ds.Tables[0], "computer='" + sComputer + "'");
            //if (r != null) iRownum = int.Parse(r["rownum"].ToString());
            dsvp = read_tables_vp(s_user_vp(DateTime.Now.Year.ToString().PadLeft(4, '0').Substring(2, 2)));
            dsfield = read_field_name();
            apiGetOracle = "http://" + ipApi + "/api/orc/getDatabySql";
            apiGetMySQL = "http://" + ipApi + "/api/mysql/getDatabySql";
            apiExcuteMySQL = "http://" + ipApi + "/api/mysql/executeQuery";
            apiGetMsSql = "http://" + ipApi + "/api/mssql/getDatabySql";
            apiExcuteMsSql = "http://" + ipApi + "/api/mssql/executeQuery";
            // _ctsdoituong = new DoiTuong();


        }
        public AccessDataApi(string _ipApi)
        {
            client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            if (Maincode("Con") != "")
            {
                sConn = Maincode("Con");
                ipApi = Maincode("Api");
            }
            sConn_mysql = mySqlConnection();
            sComputer = System.Environment.MachineName.Trim().ToUpper();            
            userid = sConn.Substring(sConn.LastIndexOf("=") + 1).Trim();
            service_name = sConn.Substring(sConn.IndexOf("=") + 1, sConn.IndexOf(";") - 1 - sConn.IndexOf("=")).Trim();
            //ds = get_data("select rownum,computer from dmcomputer");
            //DataRow r = getrowbyid(ds.Tables[0], "computer='" + sComputer + "'");
            //if (r != null) iRownum = int.Parse(r["rownum"].ToString());
            dsvp = read_tables_vp(s_user_vp(DateTime.Now.Year.ToString().PadLeft(4, '0').Substring(2, 2)));
            dsfield = read_field_name();
            ipApi = _ipApi;
            apiGetOracle = "http://" + ipApi + "/api/orc/getDatabySql";
            apiGetMySQL = "http://" + ipApi + "/api/mysql/getDatabySql";
            apiExcuteMySQL = "http://" + ipApi + "/api/mysql/executeQuery";
            apiGetMsSql = "http://" + ipApi + "/api/mssql/getDatabySql";
            apiExcuteMsSql = "http://" + ipApi + "/api/mssql/executeQuery";
            // _ctsdoituong = new DoiTuong();


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
        public JObject truyvansql(string sql)
        {
            var httprs = client.PostAsync("http://" + ipApi + "/api/orc/querydata", new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                return JObject.Parse(tjson);
            }
            else
            {
                //ghiloi(sql);
                return null;
            }

        }
        public JObject truyvanMySQL(string sql)
        {
            var httprs = client.PostAsync("http://" + ipApi + "/api/mysql/getDatabySql", new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                return JObject.Parse(tjson);
            }
            else
            {
                //ghiloi(sql);
                return null;
            }

        }
        public bool thucThiSql(string sql)
        {
            var httprs = client.PostAsync("http://" + ipApi + "/api/mysql/executeQuery", new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                var tobjsion = JObject.Parse(tjson);
                if ((bool)tobjsion["ok"] &&(int) tobjsion["data"] >0 )
                    return true;
                else
                    return false;
            }
            return false;

        }
        public bool thucThiSql(string sql,string api)
        {
            var httprs = client.PostAsync(api, new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                var tobjsion = JObject.Parse(tjson);
                if ((bool)tobjsion["ok"] && (int)tobjsion["data"] > 0)
                    return true;
                else
                {
                    //ghiloi(sql + "api:" + api);
                    return false;
                }
            }
            else
            {
                //ghiloi(sql + "api:" + api);
                return false;
            }

        }
        public JObject tryVanServerAna(string sql)
        {
            var httprs = client.PostAsync("http://" + ipApi + "/api/mssql/getDatabySql", new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                return JObject.Parse(tjson);
            }
            return null;

        }
        public bool thucThiServerAna(string sql)
        {
            var httprs = client.PostAsync("http://" + ipApi + "/api/mssql/executeQuery", new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                var tobjsion = JObject.Parse(tjson);
                if ((bool)tobjsion["ok"] && (int)tobjsion["data"] > 0)
                    return true;
                else
                {
                    //ghiloi(sql);
                    return false;
                }
            }
            else
            {
                //ghiloi(sql);
                return false;
            }

        }
        public DataSet get_data(string asql)
        {
           
            var dsjson = truyvansql(asql);

            
            if(dsjson != null && (bool)dsjson["ok"])
            {
                var arjson = dsjson["data"];
                var objson = JsonConvert.SerializeObject(new { data = arjson });
              return  JsonConvert.DeserializeObject<DataSet>(objson);
            }
            else
            {
                return null;
            }
        }
        public DataSet get_data(string asql,string api)
        {

            var dsjson = truyvansql(asql);

            if (dsjson != null && (bool)dsjson["ok"])
            {
                var arjson = dsjson["data"];
                var objson = JsonConvert.SerializeObject(new { data = arjson });
                return JsonConvert.DeserializeObject<DataSet>(objson);
            }
            else
            {
                return null;
            }
        }
        public DataSet get_data_mySQL(string asql)
        {

            var dsjson = truyvanMySQL(asql);

            if (dsjson != null && (bool)dsjson["ok"])
            {
                var arjson = dsjson["data"];
                var objson = JsonConvert.SerializeObject(new { data = arjson });
                return JsonConvert.DeserializeObject<DataSet>(objson);
            }
            else
            {
                //ghiloi(asql);
                return null;
            }
        }
        public DataSet f_get_chiphict(string sDateFormat,string sTungay, string sDenngay, string tungay, string denngay, string mabn, string loaibn, bool laytheongayxuatkhoa, string nhomvpbhyt, bool laybaocaothongke, bool laythuocglivec)
        {
            bool flag = laythuocglivec;
            loaibn = loaibn.Trim(new char[] { ',' });
            DataSet set = new DataSet();
            string str = "select a.id,a.ma,a.ten,a.dang as dvt,a.donvi,a.kythuat,a.bhyt,e.idnhombhyt as nhombhyt,0 as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,1 as thuoc,a.hamluong,a.sodk,a.maduongdung duongdung,a.tenhc,'' as lieuluong,'' as mavattu,a.masobyt,a.tenbyt,a.gia_bh_toida from " + user + ".d_dmbd a inner join " + user + ".d_dmnhom b on a.manhom=b.id inner join " + user + ".v_nhomvp e on b.nhomvp=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id  union all select c.id,cast(c.ma as varchar2(50)) as ma,cast(c.ten as nvarchar2(100)) as ten,c.dvt,null donvi,c.kythuat,c.bhyt,e.idnhombhyt as nhombhyt,d.id as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,0 as thuoc,null as hamluong,null as sodk,null as duongdung,null tenhc,null as lieuluong,c.mavattubyt as mavattu,c.masobyt,c.tenbyt,c.gia_bh_toida from " + user + ".v_giavp c inner join " + user + ".v_loaivp d on c.id_loai=d.id inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id ";
            string str2 = "";
            string str3 = user + str2;
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str2 = time3.ToString("MMyy");
                if (this.bMmyy(str2))
                {
                    str3 = user + str2;
                    string str4 = "select e.id from " + str3 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    string sql = "select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,to_char(c2.maql) maql,c.loaibn,to_char(a.id) id,to_char(b.id) as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(round(a.soluong,2)) as soluong,round(a.dongia,2) dongia,sum(round(round(a.soluong,2)*round(dongia,2),2)) as sotien1,b.loaivp,sum(round(a.bhyttra,2)) as bhyttra,sum(round(round(a.soluong,2)*round(dongia,2),2)-round(a.bhyttra,2)) as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi,case when c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%'  and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%'  then b.bhyt else case when c3.traituyen=1 then b.bhyt else 100 end end as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt,case when b.bhyt<>0 and b.bhyt<>100 and c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%' and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%' then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  case when c3.traituyen=1 then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  round(round(a.dongia,2)*sum(round(a.soluong,2)),2) end end as sotien,0 bntt,0 muchuong,0 nguonkhac from " + str3 + ".v_ttrvct a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + user + ".btdkp_bv a2 on a2.makp=a.makp inner join (" + str + ") b on a.mavp=b.id where a.madoituong=1  and a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id ";
                    sql += " where to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'" + sDateFormat + "'),'" + sDateFormat + "')  between to_date('" + sTungay + "','" + sDateFormat + "') and to_date('" + sDenngay + "','" + sDateFormat + "')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ")";
                    sql += " and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,c3.traituyen,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.bhyt,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    if (flag)
                    {
                        string str6 = sql;
                        sql = str6 + " union all select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,to_char(c2.maql) maql,c.loaibn,to_char(a.id) id,to_char(b.id)  as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(a.soluong) as soluong,a.dongia,sum(a.soluong*a.dongia) as sotien1,b.loaivp,sum(a.dongia*a.soluong) as bhyttra,0 as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi, 100  as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt, sum(a.dongia*a.soluong)  as sotien,0 bntt,0 muchuong, 0  as nguonkhac from d_xuatsd_glivec a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + user + ".btdkp_bv a2 on a2.makp=c.makp inner join (" + str + ") b on a.mabd=b.id where  a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id";
                        sql += " where to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'" + sDateFormat + "'),'" + sDateFormat + "')  between to_date('" + sTungay + "','" + sDateFormat + "') and to_date('" + sDenngay + "','" + sDateFormat + "')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ") and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    }
                    if (laybaocaothongke)
                    {
                        sql = "select loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,sum(soluong) soluong,sum(soluong)*dongia*bhyt/100 sotien,sum(sotien) sotien1 from (" + sql + ") group by loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,bhyt";
                    }
                    try
                    {
                        set.Merge(this.get_data(sql));
                    }
                    catch
                    {
                    }
                }
            }
            return set;
        }
        public DataSet f_get_chiphict(string tungay, string denngay, string mabn, string loaibn, bool laytheongayxuatkhoa, string nhomvpbhyt, bool laybaocaothongke, bool laythuocglivec)
        {
            bool flag = laythuocglivec;
            loaibn = loaibn.Trim(new char[] { ',' });
            DataSet set = new DataSet();
            string str = "select a.id,a.ma,a.ten,a.dang as dvt,a.donvi,a.kythuat,a.bhyt,e.idnhombhyt as nhombhyt,0 as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,1 as thuoc,a.hamluong,a.sodk,a.maduongdung duongdung,a.tenhc,'' as lieuluong,'' as mavattu,a.masobyt,a.tenbyt,a.gia_bh_toida from " + user + ".d_dmbd a inner join " + user + ".d_dmnhom b on a.manhom=b.id inner join " + user + ".v_nhomvp e on b.nhomvp=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id  union all select c.id,cast(c.ma as varchar2(50)) as ma,cast(c.ten as nvarchar2(100)) as ten,c.dvt,null donvi,c.kythuat,c.bhyt,e.idnhombhyt as nhombhyt,d.id as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,0 as thuoc,null as hamluong,null as sodk,null as duongdung,null tenhc,null as lieuluong,c.mavattubyt as mavattu,c.masobyt,c.tenbyt,c.gia_bh_toida from " + user + ".v_giavp c inner join " + user + ".v_loaivp d on c.id_loai=d.id inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id ";
            string str2 = "";
            string str3 = user + str2;
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str2 = time3.ToString("MMyy");
                if (this.bMmyy(str2))
                {
                    str3 = user + str2;
                    string str4 = "select e.id from " + str3 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    string sql = "select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,to_char(c2.maql) maql,c.loaibn,to_char(a.id) id,to_char(b.id) as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(round(a.soluong,2)) as soluong,round(a.dongia,2) dongia,sum(round(round(a.soluong,2)*round(dongia,2),2)) as sotien1,b.loaivp,sum(round(a.bhyttra,2)) as bhyttra,sum(round(round(a.soluong,2)*round(dongia,2),2)-round(a.bhyttra,2)) as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi,case when c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%'  and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%'  then b.bhyt else case when c3.traituyen=1 then b.bhyt else 100 end end as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt,case when b.bhyt<>0 and b.bhyt<>100 and c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%' and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%' then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  case when c3.traituyen=1 then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  round(round(a.dongia,2)*sum(round(a.soluong,2)),2) end end as sotien,0 bntt,0 muchuong,0 nguonkhac from " + str3 + ".v_ttrvct a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + user + ".btdkp_bv a2 on a2.makp=a.makp inner join (" + str + ") b on a.mavp=b.id where a.madoituong=1  and a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id where    to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy')  between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ") and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,c3.traituyen,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.bhyt,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    if (flag)
                    {
                        string str6 = sql;
                        sql = str6 + " union all select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,to_char(c2.maql) maql,c.loaibn,to_char(a.id) id,to_char(b.id)  as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(a.soluong) as soluong,a.dongia,sum(a.soluong*a.dongia) as sotien1,b.loaivp,sum(a.dongia*a.soluong) as bhyttra,0 as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi, 100  as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt, sum(a.dongia*a.soluong)  as sotien,0 bntt,0 muchuong, 0  as nguonkhac from d_xuatsd_glivec a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + user + ".btdkp_bv a2 on a2.makp=c.makp inner join (" + str + ") b on a.mabd=b.id where  a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id where    to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy')  between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ") and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    }
                    if (laybaocaothongke)
                    {
                        sql = "select loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,sum(soluong) soluong,sum(soluong)*dongia*bhyt/100 sotien,sum(sotien) sotien1 from (" + sql + ") group by loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,bhyt";
                    }
                    try
                    {
                        set.Merge(this.get_data(sql));
                    }
                    catch
                    {
                    }
                }
            }
            return set;
        }
        
        public DataSet f_get_chiphi(string sDateFormat, string sTungay, string sDenngay, string tungay, string denngay, string mabn, bool vbngoai, bool vbnoi, bool laytheongayxuatkhoa)
        {
            string str = "";
            DataSet set = new DataSet();
            string str2 = user + str;
            string sql = "";
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str = time3.ToString("MMyy");
                if (bMmyy(str))
                {
                    str2 = user + str;
                    sql = "";
                    string str4 = "select to_char(e.id) id from " + str2 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.AddMonths(-2).ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    if (vbngoai)
                    {
                        sql = "select to_char(b.id) id,to_char(a.maql) maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'dd/mm/yyyy') end as namsinh_ana,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao_ana,to_char(a.ngayra,'dd/mm/yyyy') as ngayra_ana,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,to_char(b.ngay,'dd/mm/yyyy') as ngaythu_ana,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,case when b1.tungay is null then to_char(a1.tungay,'yyyymmdd') else to_char(b1.tungay,'yyyymmdd') end as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,case when b1.tungay is null then to_char(a1.tungay,'dd/mm/yyyy') else to_char(b1.tungay,'dd/mm/yyyy') end as tungay_ana,to_char(b1.ngay,'dd/mm/yyyy') as denngay_ana,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,1 as songaydt,1 as tinhtrangrv,2 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp ,vlogin.hoten as hotenthungan from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id  inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id left join " + str2 + ".bhyt a1 on a1.maql=a.maql inner join " + user + ".btdbn a2 on a.mabn=a2.mabn left join " + user + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + user + ".btdquan a22 on a2.maqu=a22.maqu left join " + user + ".btdtt a23 on a2.matt=a23.matt left join " + user + ".btdkp_bv b2 on b2.makp=b.makp left join v_dlogin vlogin on vlogin.id=b.userid where  b.loaibn in(3,2,4) and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'" + sDateFormat + "'),'" + sDateFormat + "') between to_date('" + sTungay + "','" + sDateFormat + "') and to_date('" + sDenngay + "','" + sDateFormat + "') and b.bhyt>0" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    if (vbnoi)
                    {
                        string str5 = sql;
                        sql = str5 + ((sql == "") ? "" : " union all ") + " select to_char(b.id) id,to_char(a.maql) maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'dd/mm/yyyy') end as namsinh_ana,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao_ana,to_char(a.ngayra,'dd/mm/yyyy') as ngayra_ana,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,to_char(b.ngay,'dd/mm/yyyy') as ngaythu_ana,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,to_char(b1.tungay,'yyyymmdd') as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,to_char(b1.tungay,'dd/mm/yyyy') as tungay_ana,to_char(b1.ngay,'dd/mm/yyyy') as denngay_ana,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,round(to_date(to_char(ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy'))+1 as songaydt,a43.mabhyt2348 as tinhtrangrv,a42.mabhyt2348 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp ,vlogin.hoten as hotenthungan from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id inner join " + user + ".btdbn a2 on a.mabn=a2.mabn left join " + user + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + user + ".btdquan a22 on a2.maqu=a22.maqu left join " + user + ".btdtt a23 on a2.matt=a23.matt left join " + user + ".xuatvien a4 on a.maql=a4.maql left join " + user + ".btdkp_bv b2 on b2.makp=b.makp left join " + user + ".ketqua a42 on a42.ma=a4.ketqua left join " + user + ".ttxk a43 on a43.ma=a4.ttlucrv left join v_dlogin vlogin on vlogin.id=b.userid where  b.loaibn in(1) and b.bhyt>0 and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'" + sDateFormat + "'),'" + sDateFormat + "') between to_date('" + sTungay + "','" + sDateFormat + "') and to_date('" + sDenngay + "','" + sDateFormat + "')" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    try
                    {
                        if (set.Tables[0].Rows.Count > 0)
                        {
                            set.Merge(get_data(sql));
                        }
                        else
                        {
                            set = get_data(sql);
                        }
                    }
                    catch
                    {
                        set = get_data(sql);
                    }
                }
            }
            return set;
        }
        public DataSet f_get_chiphi(string tungay, string denngay, string mabn, bool vbngoai, bool vbnoi, bool laytheongayxuatkhoa)
        {
            string str = "";
            DataSet set = new DataSet();
            string str2 = user + str;
            string sql = "";
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str = time3.ToString("MMyy");
                if (bMmyy(str))
                {
                    str2 = user + str;
                    sql = "";
                    string str4 = "select to_char(e.id) id from " + str2 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.AddMonths(-2).ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    if (vbngoai)
                    {
                        sql = "select to_char(b.id) id,to_char(a.maql) maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'dd/mm/yyyy') end as namsinh_ana,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao_ana,to_char(a.ngayra,'dd/mm/yyyy') as ngayra_ana,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,to_char(b.ngay,'dd/mm/yyyy') as ngaythu_ana,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,case when b1.tungay is null then to_char(a1.tungay,'yyyymmdd') else to_char(b1.tungay,'yyyymmdd') end as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,case when b1.tungay is null then to_char(a1.tungay,'dd/mm/yyyy') else to_char(b1.tungay,'dd/mm/yyyy') end as tungay_ana,to_char(b1.ngay,'dd/mm/yyyy') as denngay_ana,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,1 as songaydt,1 as tinhtrangrv,2 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp ,vlogin.hoten as hotenthungan from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id  inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id left join " + str2 + ".bhyt a1 on a1.maql=a.maql inner join " + user + ".btdbn a2 on a.mabn=a2.mabn left join " + user + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + user + ".btdquan a22 on a2.maqu=a22.maqu left join " + user + ".btdtt a23 on a2.matt=a23.matt left join " + user + ".btdkp_bv b2 on b2.makp=b.makp left join v_dlogin vlogin on vlogin.id=b.userid where  b.loaibn in(3,2,4) and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy') and b.bhyt>0" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    if (vbnoi)
                    {
                        string str5 = sql;
                        sql = str5 + ((sql == "") ? "" : " union all ") + " select to_char(b.id) id,to_char(a.maql) maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'dd/mm/yyyy') end as namsinh_ana,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao_ana,to_char(a.ngayra,'dd/mm/yyyy') as ngayra_ana,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,to_char(b.ngay,'dd/mm/yyyy') as ngaythu_ana,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,to_char(b1.tungay,'yyyymmdd') as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,to_char(b1.tungay,'dd/mm/yyyy') as tungay_ana,to_char(b1.ngay,'dd/mm/yyyy') as denngay_ana,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,round(to_date(to_char(ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy'))+1 as songaydt,a43.mabhyt2348 as tinhtrangrv,a42.mabhyt2348 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp ,vlogin.hoten as hotenthungan from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id inner join " + user + ".btdbn a2 on a.mabn=a2.mabn left join " + user + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + user + ".btdquan a22 on a2.maqu=a22.maqu left join " + user + ".btdtt a23 on a2.matt=a23.matt left join " + user + ".xuatvien a4 on a.maql=a4.maql left join " + user + ".btdkp_bv b2 on b2.makp=b.makp left join " + user + ".ketqua a42 on a42.ma=a4.ketqua left join " + user + ".ttxk a43 on a43.ma=a4.ttlucrv left join v_dlogin vlogin on vlogin.id=b.userid where  b.loaibn in(1) and b.bhyt>0 and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    try
                    {
                        if (set.Tables[0].Rows.Count > 0)
                        {
                            set.Merge(get_data(sql));
                        }
                        else
                        {
                            set = get_data(sql);
                        }
                    }
                    catch
                    {
                        set = get_data(sql);
                    }
                }
            }
            return set;
        }
        public string f_get_ngayra(string maql, string id)
        {
            string str = "";
            string sql = "";
            string str4 = sql;
            sql = str4 + "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + user + "d" + maql.Substring(2, 2) + maql.Substring(0, 2) + ".bhytkb WHERE maql in(" + maql + ")";
            try
            {
                foreach (DataRow row in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row["ngay"].ToString();
                    }
                }
                if (!(str == ""))
                {
                    return str;
                }
                sql = "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + user + "d" + id.Substring(2, 2) + id.Substring(0, 2) + ".bhytkb WHERE maql =" + maql + " or idttrv=" + id;
                foreach (DataRow row2 in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row2["ngay"].ToString();
                    }
                }
            }
            catch
            {
            }
            return str;
        }
        public string f_get_ngayra_ana(string maql, string id)
        {
            string str = "";
            string sql = "";
            string str4 = sql;
            sql = str4 + "select distinct to_char(ngayud,'dd/mm/yyyy') ngay from " + user + "d" + maql.Substring(2, 2) + maql.Substring(0, 2) + ".bhytkb WHERE maql in(" + maql + ")";
            try
            {
                foreach (DataRow row in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row["ngay"].ToString();
                    }
                }
                if (!(str == ""))
                {
                    return str;
                }
                sql = "select distinct to_char(ngayud,'dd/mm/yyyy') ngay from " + user + "d" + id.Substring(2, 2) + id.Substring(0, 2) + ".bhytkb WHERE maql =" + maql + " or idttrv=" + id;
                foreach (DataRow row2 in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row2["ngay"].ToString();
                    }
                }
            }
            catch
            {
            }
            return str;
        }
        public string f_get_ngayvao(string maql, string mabn)
        {
            string str = "";
            string sql = "";
            string str4 = sql;
            string str5 = str4 + "#select distinct to_char(ngay,'yyyymmddhh24mi') ngay from " + user + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE maql in(" + maql + ")";
            sql = (str5 + "#select distinct to_char(ngay,'yyyymmddhh24mi') ngay from " + user + ".benhandt WHERE maql =" + maql + " ").Trim(new char[] { '#' }).Replace("#", " union all ");
            try
            {
                foreach (DataRow row in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row["ngay"].ToString();
                    }
                }
                if (!(str == ""))
                {
                    return str;
                }
                sql = "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + user + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE to_char(maql) like '" + maql.Substring(0, 6) + "%' and mabn='" + mabn + "'";
                foreach (DataRow row2 in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row2["ngay"].ToString();
                    }
                }
            }
            catch
            {
            }
            return str;
        }
        public string f_get_ngayvao_ana(string maql, string mabn)
        {
            string str = "";
            string sql = "";
            string str4 = sql;
            string str5 = str4 + "#select distinct to_char(ngay,'dd/mm/yyyy') ngay from " + user + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE maql in(" + maql + ")";
            sql = (str5 + "#select distinct to_char(ngay,'dd/mm/yyyy') ngay from " + user + ".benhandt WHERE maql =" + maql + " ").Trim(new char[] { '#' }).Replace("#", " union all ");
            try
            {
                foreach (DataRow row in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row["ngay"].ToString();
                    }
                }
                if (!(str == ""))
                {
                    return str;
                }
                sql = "select distinct to_char(ngayud,'dd/mm/yyyy') ngay from " + user + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE to_char(maql) like '" + maql.Substring(0, 6) + "%' and mabn='" + mabn + "'";
                foreach (DataRow row2 in get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row2["ngay"].ToString();
                    }
                }
            }
            catch
            {
            }
            return str;
        }
        public string f_get_ngayvao_ra(string ngayvao, double sogio)
        {
            string str = ngayvao;
            try
            {
                DateTime time = new DateTime(int.Parse(ngayvao.Substring(0, 4)), int.Parse(ngayvao.Substring(4, 2)), int.Parse(ngayvao.Substring(6, 2)), int.Parse(ngayvao.Substring(8, 2)), int.Parse(ngayvao.Substring(10, 2)), 0);
                return time.AddHours(sogio).ToString("yyyyMMddHHmm");
            }
            catch
            {
            }
            return str;
        }

        public string f_get_sqlfull(string vsql, string tungay, string denngay)
        {
            string str = "";
            string str2 = "";
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            for (DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1); time2 <= time; time2 = time2.AddMonths(1))
            {
                str = time2.ToString("MMyy");
                if (bMmyy(str))
                {
                    str2 = str2 + vsql.Replace("xxxd.", user + "d" + str + ".").Replace("xxx.", user + str + ".") + "#";
                }
            }
            return str2.Trim(new char[] { '#' }).Replace("#", " union all ");
        }
        void ghiloi(string loi)
        {
            if (!Directory.Exists("data"))
            {
                Directory.CreateDirectory("data");
            }
            System.IO.File.AppendAllText("data\\error.txt",DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + " " + loi + "\n"); 
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
        public int iTunguyen
        {
            get
            {
                try
                {
                    DataSet ds = get_data("select ten from thongso where id=149");
                    return int.Parse(ds.Tables[0].Rows[0][0].ToString());
                }
                catch { return 2; }
            }
        }
        public bool bmMmyy(string mmyy)
        {
            return get_data("select * from m_table where mmyy='" + mmyy + "'").Tables[0].Rows.Count > 0;
        }
        public bool s_giavbangdongiacongvattu
        {
            get
            {
                bool r = false;
                try
                {
                    r = (get_data("select giatri from v_option where ma='chkGiavpbangdongiacongvattu'").Tables[0].Rows[0][0].ToString().Trim() == "1");
                }
                catch
                {
                    r = false;
                }
                return r;
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
        public DataSet get_data_mmyy(string str, string tu, string den,string apiType)
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

                            tmp = apidata.dataSetFromSql(apiType, sql);
                            be = false;
                        }
                        else tmp.Merge(get_data(sql));
                    }
                }
            }
            return tmp;
        }
        public DataSet get_data_all_vp(DateTime v_tu, DateTime v_den, string v_sql)
        {
            string asql = build_sqlvp(v_tu, v_den, v_sql);
            return get_data(asql);
        }
        public string build_sqlvp(DateTime v_tu, DateTime v_den, string v_sql)
        {
            string auser = "", sql = "", asqlo = "", ausert = "", atmp = "";
            int n = 0;
            DateTime atu, aden;
            v_sql = v_sql.Trim().ToLower().Replace(")", " )");
            v_sql = v_sql.Trim().ToLower().Replace(",", " ,");
            while (v_sql.IndexOf("  ,") >= 0)
            {
                v_sql = v_sql.Replace("  ,", " ,");
            }
            while (v_sql.IndexOf("  )") >= 0)
            {
                v_sql = v_sql.Replace("  )", " )");
            }
            v_sql = v_sql.Trim() + " ";
            asqlo = v_sql;
            try
            {
                foreach (DataRow r in dsvp.Tables[0].Rows)
                {
                    sql = "";
                    ausert = "";
                    atmp = get_field_vp(r["table_name"].ToString().Trim().ToUpper());
                    n = 0;
                    atu = new DateTime(v_tu.Year, v_tu.Month, 1, 1, 1, 1, 1);
                   // atu = atu.AddMonths(-1);
                    aden = new DateTime(v_den.Year, v_den.Month, 1, 1, 1, 1, 1);
                   // aden = aden.AddMonths(1);
                    if (atu.Day > 1) atu = atu.AddDays(-(atu.Day - 1));


                    while (DateTime.Compare(aden, atu) >= 0)//(atu<=aden)
                    {
                        auser = s_user_vp(atu.Month.ToString().PadLeft(2, '0') + atu.Year.ToString().Substring(2));
                        if (s_iscreated(auser))
                        {
                            if (r["table_name"].ToString() == "V_TTRVCT_chenhlech")//ThanhCuong-17062011
                            {
                                if (v_sql.IndexOf(r["table_name"].ToString().ToLower().Trim() + " ") >= 0)
                                {
                                    n++;
                                    ausert = auser + "." + r["table_name"].ToString().Substring(0, 8).ToLower().Trim();
                                    if (sql.Length > 0)
                                    {
                                        sql = sql + " union all select " + atmp + " from " + auser + "." + r["table_name"].ToString().Substring(0, 8).ToLower().Trim() + " where madoituong=" + iTunguyen + " group by id,makp,ngay ";
                                    }
                                    else
                                    {
                                        sql = "select " + atmp + " from " + auser + "." + r["table_name"].ToString().Substring(0, 8).ToLower().Trim() + " where madoituong=" + iTunguyen + " group by id,makp,ngay ";
                                    }
                                }
                            }
                            else
                            {
                                if (v_sql.IndexOf(r["table_name"].ToString().ToLower().Trim() + " ") >= 0)
                                {
                                    n++;
                                    ausert = auser + "." + r["table_name"].ToString().ToLower().Trim();
                                    if (sql.Length > 0)
                                    {
                                        sql = sql + " union all select " + atmp + " from " + auser + "." + r["table_name"].ToString().ToLower().Trim();
                                    }
                                    else
                                    {
                                        sql = "select " + atmp + " from " + auser + "." + r["table_name"].ToString().ToLower().Trim();
                                    }
                                }
                            }
                        }
                        atu = atu.AddMonths(1);
                    }
                    sql = sql.Trim();

                    if (sql.Length > 0)
                    {
                        if (n > 1)
                        {
                            v_sql = v_sql.Replace(r["table_name"].ToString().ToLower().Trim() + " ", "(" + sql + ") ");
                        }
                        else
                        {
                            v_sql = v_sql.Replace(r["table_name"].ToString().ToLower().Trim() + " ", ausert + " ");
                        }
                    }
                }
            }
            catch
            {
                v_sql = asqlo;
            }
            return v_sql;
        }
        public void f_Ngoaitru_xuatExcel_mau79_Mau2(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("stt", typeof(decimal));
            dset.Tables[0].Columns.Add("ma_bn", typeof(string));
            dset.Tables[0].Columns.Add("ho_ten", typeof(string));
            dset.Tables[0].Columns.Add("ngay_sinh", typeof(string));
            dset.Tables[0].Columns.Add("gioi_tinh", typeof(int));
            dset.Tables[0].Columns.Add("dia_chi", typeof(string));
            dset.Tables[0].Columns.Add("ma_the", typeof(string));
            dset.Tables[0].Columns.Add("ma_dkbd", typeof(string));
            dset.Tables[0].Columns.Add("gt_the_tu", typeof(string));
            dset.Tables[0].Columns.Add("gt_the_den", typeof(string));
            dset.Tables[0].Columns.Add("ma_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benhkhac", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("tenkp", typeof(string));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_xn", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_cdha", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_thuoc", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_mau", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_pttt", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_vtyt", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_dvkt_tyle", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_thuoc_tyle", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_vtyt_tyle", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_kham", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_giuong", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_vchuyen", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("ma_khoa", typeof(string));
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ngaythanhtoan", typeof(string));
            dset.Tables[0].Columns.Add("maql", typeof(decimal));
            dsdulieu.WriteXml("tam.xml", XmlWriteMode.WriteSchema);
            string mabv = Mabv;
            decimal d = 0M;
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    DataRow row2 = dset.Tables[0].NewRow();
                    row2["STT"] = d++;
                    row2["ma_bn"] = row["mabn"].ToString();
                    row2["ho_ten"] = Replace_ErrorFont(row["HOTEN"].ToString());
                    row2["ngay_sinh"] = row["ngaysinh"].ToString();
                    row2["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row2["dia_chi"] = Replace_ErrorFont(row["diachi"].ToString());
                    try
                    {
                        row2["ma_the"] = row["sothe"].ToString();
                    }
                    catch
                    {
                        row2["ma_the"] = row["sothe"].ToString();
                    }
                    row2["ma_dkbd"] = row["MANOIDK"].ToString();
                    if (row["tungay"].ToString().IndexOf('/') > -1)
                    {
                        try
                        {
                            row2["gt_the_tu"] = row["tungay"].ToString().Substring(6, 4) + row["tungay"].ToString().Substring(3, 2) + row["tungay"].ToString().Substring(0, 2);
                        }
                        catch
                        {
                            row2["gt_the_tu"] = "";
                        }
                    }
                    else
                    {
                        try
                        {
                            row2["gt_the_tu"] = row["tungay"].ToString();
                        }
                        catch
                        {
                            row2["gt_the_tu"] = "";
                        }
                    }
                    if (row["denngay"].ToString().IndexOf('/') > -1)
                    {
                        try
                        {
                            row2["gt_the_den"] = row["denngay"].ToString().Substring(6, 4) + row["denngay"].ToString().Substring(3, 2) + row["denngay"].ToString().Substring(0, 2);
                        }
                        catch
                        {
                            row2["gt_the_den"] = "";
                        }
                    }
                    else
                    {
                        try
                        {
                            row2["gt_the_den"] = row["denngay"].ToString();
                        }
                        catch
                        {
                            row2["gt_the_den"] = "";
                        }
                    }
                    row2["ma_benh"] = this.f_get_fix_maicd(row["MAICD"].ToString());
                    try
                    {
                        row2["ma_benh"] = row2["ma_benh"].ToString().Split(new char[] { ';' })[0];
                    }
                    catch
                    {
                    }
                    row2["ma_benhkhac"] = row["maicdkt"].ToString();
                    row2["ma_lydo_vvien"] = row["lydo"].ToString();
                    if (row["traituyen"].ToString() != "0")
                    {
                        row2["ma_lydo_vvien"] = 3;
                    }
                    else if (row["traituyen"].ToString() == "0")
                    {
                        row2["ma_lydo_vvien"] = 1;
                    }
                    else
                    {
                        row2["ma_lydo_vvien"] = 2;
                    }
                    row2["ma_noi_chuyen"] = "";
                    if (row["NGAYVAO"].ToString().IndexOf('/') > -1)
                    {
                        try
                        {
                            row2["ngay_vao"] = row["NGAYVAO"].ToString().Substring(6, 4) + row["NGAYVAO"].ToString().Substring(3, 2) + row["NGAYVAO"].ToString().Substring(0, 2);
                        }
                        catch
                        {
                            row2["ngay_vao"] = "";
                        }
                    }
                    else
                    {
                        try
                        {
                            row2["ngay_vao"] = row["NGAYVAO"].ToString();
                        }
                        catch
                        {
                            row2["ngay_vao"] = "";
                        }
                    }
                    if (row["NGAYRA"].ToString().IndexOf('/') > -1)
                    {
                        try
                        {
                            row2["ngay_ra"] = row["NGAYRA"].ToString().Substring(6, 4) + row["NGAYRA"].ToString().Substring(3, 2) + row["NGAYRA"].ToString().Substring(0, 2);
                        }
                        catch
                        {
                            row2["ngay_ra"] = "";
                        }
                    }
                    else
                    {
                        try
                        {
                            row2["ngay_ra"] = row["NGAYRA"].ToString();
                        }
                        catch
                        {
                            row2["ngay_ra"] = "";
                        }
                    }
                    if (row["NGAYTHU"].ToString().IndexOf('/') > -1)
                    {
                        try
                        {
                            row2["ngaythanhtoan"] = row["NGAYTHU"].ToString().Substring(6, 4) + row["NGAYTHU"].ToString().Substring(3, 2) + row["NGAYTHU"].ToString().Substring(0, 2);
                        }
                        catch
                        {
                            row2["ngaythanhtoan"] = "";
                        }
                    }
                    else
                    {
                        try
                        {
                            row2["ngaythanhtoan"] = row["NGAYTHU"].ToString();
                        }
                        catch
                        {
                            row2["ngaythanhtoan"] = "";
                        }
                    }
                    try
                    {
                        row2["so_ngay_dtri"] = row["SONGAY"].ToString();
                    }
                    catch
                    {
                        row2["so_ngay_dtri"] = 1;
                    }
                    row2["ket_qua_dtri"] = 1;
                    row2["tinh_trang_rv"] = 1;
                    row2["tenkp"] = Replace_ErrorFont(row["tenkp"].ToString());
                    row2["t_tongchi"] = row["tongcong"].ToString();
                    try
                    {
                        row2["t_xn"] = row["st_1"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_cdha"] = row["st_2"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_thuoc"] = row["st_3"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_mau"] = row["st_4"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_pttt"] = row["st_5"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_vtyt"] = row["st_6"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_dvkt_tyle"] = row["st_7"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_kham"] = row["st_9"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_giuong"] = row["st_11"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["t_vchuyen"] = row["st_10"].ToString();
                    }
                    catch
                    {
                    }
                    row2["t_bhtt"] = row["bhyttra"].ToString();
                    try
                    {
                        row2["t_bntt"] = decimal.Parse(row2["t_tongchi"].ToString()) - decimal.Parse(row2["t_bhtt"].ToString());
                    }
                    catch
                    {
                    }
                    row2["t_ngoaids"] = 0;
                    row2["nam_qt"] = denngay.Substring(6, 4);
                    row2["thang_qt"] = denngay.Substring(3, 2);
                    row2["ma_loaikcb"] = 1;
                    if (row2["t_giuong"].ToString() != "0")
                    {
                        row2["ma_loaikcb"] = 3;
                    }
                    row2["ma_cskcb"] = Mabv;
                    try
                    {
                        row2["ma_khuvuc"] = row["makhuvuc"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row2["ma_khoa"] = row["makp_byt"].ToString();
                    }
                    catch
                    {
                    }
                    dset.Tables[0].Rows.Add(row2);
                }
                catch 
                {
                    //this._lib.f_write_log(exception.ToString());
                }
            }
            dset.Tables[0].Columns.Remove("maql");
            for (int i = 0; i < dset.Tables[0].Columns.Count; i++)
            {
                dset.Tables[0].Columns[i].ColumnName = dset.Tables[0].Columns[i].ColumnName.ToUpper();
            }
            dset.AcceptChanges();
            dset.WriteXml("bang7980.xml", XmlWriteMode.WriteSchema);
            this.tenfile = Export_Excel(dset, "bccpkcb02");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }
        public string f_hoten(string hoten)
        {
            string str = "";
            foreach (string str2 in hoten.Trim().Split(new char[] { ' ' }))
            {
                try
                {
                    string str3 = str2.Replace(" ", "");
                    str = str + str3.Substring(0, 1).ToUpper() + str3.Substring(1).ToLower() + " ";
                }
                catch
                {
                }
            }
            return str.Trim(new char[] { ' ' });
        }
        public string Replace_ErrorFont(string errFontString)
        {
            convert_font.Convert(ref errFontString, SCreenReader.FontIndex.iNotKnown, SCreenReader.FontIndex.iUNI);
            errFontString = errFontString.Replace("áº ", "Ạ");
            errFontString = errFontString.Replace("Æ ", "Ơ");
            errFontString = errFontString.Replace("Ãư", "í");

            return errFontString;
        }
        public string f_get_fix_maicd(string maicd)
        {
            string str = "";
            try
            {
                string[] strArray = maicd.Replace(" ^", "").Replace("; ", ";").Replace(",", ";").Trim().Trim(new char[] { ';' }).Split(new char[] { ';' });
                str = ";";
                for (int i = 0; i < strArray.Length; i++)
                {
                    if (!(strArray[i] == "") && (str.IndexOf(";" + strArray[i] + ";") <= -1))
                    {
                        str = str + strArray[i] + ";";
                    }
                }
                str = str.Trim(new char[] { ';' });
            }
            catch
            {
            }
            return str;
        }
        public string Export_Excel(DataSet dset, string tenfile)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "Excel";
                string str2 = path + @"\" + tenfile + ".xls";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                StreamWriter writer = new StreamWriter(str2, false, Encoding.Unicode);
                string str3 = "";
                str3 = "<Table>";
                str3 = str3 + "<tr>";
                for (int i = 0; i < dset.Tables[0].Columns.Count; i++)
                {
                    str3 = (str3 + "<th>") + dset.Tables[0].Columns[i].ColumnName + "</th>";
                }
                str3 = str3 + "</tr>";
                writer.Write(str3);
                for (int j = 0; j < dset.Tables[0].Rows.Count; j++)
                {
                    str3 = "<tr>";
                    for (int k = 0; k < dset.Tables[0].Columns.Count; k++)
                    {
                        str3 = (str3 + "<td>") + dset.Tables[0].Rows[j][k].ToString() + "</td>";
                    }
                    str3 = str3 + "</tr>";
                    writer.Write(str3);
                }
                str3 = "</Table>";
                writer.Write(str3);
                writer.Close();
                return str2;
            }
            catch 
            {
                //this.upd_error(exception.Message, this.sComputer, tenfile);
                return "";
            }
        }
        public string ngaygiohienhanh_server
        {
            get
            {
                return get_data("select to_char(sysdate,'dd/mm/yyyy hh24:mi') as ngay from dual").Tables[0].Rows[0]["ngay"].ToString();
            }
        }
        public string ngayhienhanh_server
        {
            get
            {
                return get_data("select to_char(sysdate,'dd/mm/yyyy') as ngay from dual").Tables[0].Rows[0]["ngay"].ToString();
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
        public string Ngayhienhanh_Client
        {
            get { return DateTime.Now.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Year.ToString(); }
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
        public System.DateTime StringToDate(string s)
        {
            string[] format = { "dd/MM/yyyy" };
            return System.DateTime.ParseExact(s.Substring(0, 10), format, System.Globalization.DateTimeFormatInfo.CurrentInfo, System.Globalization.DateTimeStyles.None);
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

        public bool upd_hddt_tonghop_mysql_API(string m_sohoso, decimal m_sotien,string m_tenkhachhang,string m_diachi,string m_NGUOITHU, string m_HOTENNGUOITHU, string m_ngayhd, 
            string m_sochungtu,int m_loaibn,string m_khoadieutri,string m_dotdieutri,string m_makhachhang)
        {
            //sql = "update tb_master set tenkhachhang=N'" + m_tenkhachhang + "',diachi=N'" + m_diachi + "',NGUOITHU='" + m_NGUOITHU + "',HOTENNGUOITHU=N'" + m_HOTENNGUOITHU + "',loaibn="+ m_loaibn + ", tongtientt=" + m_sotien + ",sochungtu ='" + m_sochungtu + "',ngaylap=STR_TO_DATE('" + m_ngayhd + "', '%d-%m-%Y %H:%i')";
            //sql += ",DOTDIEUTRI=N'" + m_dotdieutri + "',KHOADIEUTRI=N'" + m_khoadieutri + "',makhachhang='" + m_makhachhang + "'";
            //sql += " where sohoso='" + m_sohoso + "'";
            //if (!thucThiSql(sql))
            //{
                sql = "INSERT INTO tb_master(id,sohoso,sochungtu,makhachhang,tenkhachhang,loaibn,DOTDIEUTRI,KHOADIEUTRI,diachi,nguoithu,HOTENNGUOITHU,tongtientt,ngaylap)";
                sql += " values ('" + m_sohoso + "','" + m_sohoso + "','" + m_sochungtu + "','" + m_makhachhang + "',N'" + m_tenkhachhang + "',"+ m_loaibn + ",N'" + m_dotdieutri + "',N'" + m_khoadieutri + "',N'" + m_diachi + "','" + m_NGUOITHU + "',N'" + m_HOTENNGUOITHU + "'," + m_sotien + ",STR_TO_DATE('" + m_ngayhd + "', '%d-%m-%Y %H:%i'))";
                if (thucThiSql(sql))
                    return true;
                else
                    return false;
            //}
            
            //else
            //    return true;
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
            sql = "update tb_master set tenkhachhang=:m_tenkhachhang,diachi=:m_diachi,tongtientt=:m_sotien,nguoithu=:m_idthungan,hotennguoithu=:m_hotenthungan,ngaylap=STR_TO_DATE(:m_ngayhd, '%d-%m-%Y %H:%i')";
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
                    sql += " values (:m_sohoso,:m_tenkhachhang,:m_diachi,:m_sotien,STR_TO_DATE(:m_ngayhd, '%d-%m-%Y %H:%i'))";

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

        public bool upd_hddt_chitiet_mysql_API(string m_sohoso, string m_ma, string m_loaidv, string m_tenhang, double m_dongia, string m_donvitinh,
            double m_soluong, double m_thanhtien, double m_TYLEBH, double m_MUCBHTRA, double m_BHXHTRA)
        {

            sql = "update tb_detail set MAHANG='" + m_ma + "',LOAIDV='" + m_loaidv + "',tenhang='" + m_tenhang + "',dongia=" + m_dongia + ",donvitinh='" + m_donvitinh + "',";
            sql += " soluong=" + m_soluong + ",thanhtien=" + m_thanhtien + ",TONGTIEN=" + m_thanhtien + ",TYLEBH=" + m_TYLEBH + ",MUCBHTRA=" + m_MUCBHTRA + ",BHXHTRA=" + m_BHXHTRA;
            sql += " where SOHOSO='" + m_sohoso + "'";
            if (!thucThiSql(sql))
            {
                sql = "insert into tb_detail (MASTERID,SOHOSO,MAHANG,LOAIDV,tenhang,donvitinh,soluong,dongia,thanhtien,TYLEBH,MUCBHTRA,BHXHTRA)";
                sql += " values ('" + m_sohoso + "','" + m_sohoso + "','" + m_ma + "','" + m_loaidv + "','" + m_tenhang + "','" + m_donvitinh + "'," + m_soluong + "," + m_dongia + "," + m_thanhtien + "," + m_TYLEBH + "," + m_MUCBHTRA + "," + m_BHXHTRA + ")";
                if (thucThiSql(sql))
                    return true;
                else
                    return false;
            }

            else
                return true;
        }
        public bool upd_hddt_chitiet_mysql_API(string para)
        {
                sql = "insert into tb_detail (MASTERID,SOHOSO,MAHANG,MANHOMDV,TENNHOMDV,LOAIDV,tenhang,donvitinh,soluong,dongia,thanhtien,TYLEBH,MUCBHTRA,BHXHTRA)";
            sql += " values " + para;
         return thucThiSql(sql);


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
        public DataSet read_tables_vp(string v_user)
        {
            DataSet ads = new DataSet();
            ads.Tables.Add("Table");
            ads.Tables[0].Columns.Add("TABLE_NAME");
            ads.Tables[0].Rows.Add(new string[] { "V_BHYT" });
            ads.Tables[0].Rows.Add(new string[] { "V_CHIDINH" });
            ads.Tables[0].Rows.Add(new string[] { "V_CONGNO" });
            ads.Tables[0].Rows.Add(new string[] { "V_ERROR" });
            ads.Tables[0].Rows.Add(new string[] { "V_GIUONG" });
            ads.Tables[0].Rows.Add(new string[] { "V_SUATAN" });
            ads.Tables[0].Rows.Add(new string[] { "V_HOANTRA" });
            ads.Tables[0].Rows.Add(new string[] { "V_HOANTRACT" });
            ads.Tables[0].Rows.Add(new string[] { "V_MIENNGTRU" });
            ads.Tables[0].Rows.Add(new string[] { "V_MIENNOITRU" });
            ads.Tables[0].Rows.Add(new string[] { "V_PHIEUCHICT" });
            ads.Tables[0].Rows.Add(new string[] { "V_PHIEUCHILL" });
            ads.Tables[0].Rows.Add(new string[] { "V_TAMUNGCT" });
            ads.Tables[0].Rows.Add(new string[] { "V_TAMUNG" });
            ads.Tables[0].Rows.Add(new string[] { "V_THNGAYCT" });
            ads.Tables[0].Rows.Add(new string[] { "V_THNGAYLL" });
            ads.Tables[0].Rows.Add(new string[] { "V_THVPBHYT" });
            ads.Tables[0].Rows.Add(new string[] { "V_THVPCT" });
            ads.Tables[0].Rows.Add(new string[] { "V_THVPLL" });
            ads.Tables[0].Rows.Add(new string[] { "V_TRONGOI" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVBHYT" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVCT" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVCT_chenhlech" });//ThanhCuong-17062011
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVDS" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVLL" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVNHOM" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVPTTT" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVPTTTCT" });
            ads.Tables[0].Rows.Add(new string[] { "V_VIENPHICT" });
            ads.Tables[0].Rows.Add(new string[] { "V_VIENPHILL" });
            ads.Tables[0].Rows.Add(new string[] { "V_VPKHOA" });
            ads.Tables[0].Rows.Add(new string[] { "V_HUYBIENLAI" });
            ads.Tables[0].Rows.Add(new string[] { "TIEPDON" });
            ads.Tables[0].Rows.Add(new string[] { "V_TTRVTHUE" });
            ads.Tables[0].Rows.Add(new string[] { "V_SOVANG" });
            ads.Tables[0].Rows.Add(new string[] { "V_DACHIHOAN" });
            return ads;
        }
        public DataSet read_field_name()
        {
            DataSet ads = new DataSet();
            ads.Tables.Add("Tables");
            ads.Tables[0].Columns.Add("loai");
            ads.Tables[0].Columns.Add("table_name");
            ads.Tables[0].Columns.Add("field_name");
            ads.Tables[0].Rows.Add(new string[] { "VP", "TIEPDON", "mabn,maql,makp,ngay,madoituong,sovaovien,tuoivao,done,bnmoi,noitiepdon,loai,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_BHYT", "id,sothe,maphu,mabv,noigioithieu" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_CHIDINH", "id,mabn,maql,idkhoa,ngay,loai,makp,madoituong,mavp,soluong,dongia,paid,done,userid,ngayud,vattu,tinhtrang,thuchien,computer" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_CONGNO", "mabn,maql,idkhoa,mavp,sotien,dathu" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_ERROR", "message,computer,tables,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_GIUONG", "id,mavp,ngay,dongia" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_HOANTRA", "id,quyenso,sobienlai,ngay,mabn,hoten,sotien,ghichu,userid,ngayud,loai,loaibn" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_HOANTRACT", "id,loaivp,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_MIENNGTRU", "id,sotien,ghichu,maduyet,lydo" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_MIENNOITRU", "id,lydo,ghichu,maduyet" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_PHIEUCHICT", "id,stt,mavp,sotien,soluong,dongia,bhyt,mien,tenvp" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_PHIEUCHILL", "id,quyenso,sobienlai,ngay,mabn,maql,idkhoa,makp,hoten,loai,loaibn,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_SOVANG", "id,quyenso,sobienlai,tong,tyle,sotien,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TAMUNG", "id,mabn,maql,idkhoa,quyenso,sobienlai,ngay,loai,makp,madoituong,sotien,userid,ngayud,done,lanin,loaibn,idttrv" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TAMUNGCT", "id,loaivp,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_THNGAYCT", "id,ngay,mavp,soluong,dongia,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_THNGAYLL", "id,madoituong,mabn,maql,ngayvao,tu,den,giuong,makp,conlai,sotien,datra,done" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_THVPBHYT", "id,sothe,maphu,noigioithieu,noicap,traituyen" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_THVPCT", "id,ngay,makp,madoituong,mavp,soluong,dongia,sotien,vattu,sothe,done,idttrv,bhyttra" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_THVPLL", "id,mabn,maql,idkhoa,idttrv,ngayvao,ngayra,giuong,makp,chandoan,maicd,sotien,tamung,bhyt,mien,done" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TRONGOI", "id,sotien,tamung,hoantra,pm,yc,ghichu" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVBHYT", "id,sothe,maphu,tungay,ngay,mabv,noigioithieu" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVCT", "id,stt,ngay,makp,madoituong,mavp,soluong,dongia,vat,vattu,sotien,sothe,bhyttra,mien" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVCT_chenhlech", "id,makp,ngay,sum(soluong*dongia) chenhlechdv" });//ThanhCuong-17062011
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVDS", "id,mabn,maql,idkhoa,giuong,ngayvao,ngayra,chandoan,maicd" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVLL", "ngay,makp,sotien,tamung,mien,bhyt,userid,ngayud,lanin,id,loaibn,loai,quyenso,sobienlai,nopthem,thieu,chenhlech" });//,nopthem,thieu,vattu,chenhlech,idtonghop"});
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVNHOM", "id,ma,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVPTTT", "id,ngay,songay_tpt,songay_spt,mavp,so,loaipt,tenpt" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVPTTTCT", "id,stt,songay,dongia,loaipt" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_VIENPHICT", "id,stt,madoituong,mavp,soluong,dongia,mien,thieu,vattu,mabs,makp,idchidinh" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_VIENPHILL", "id,quyenso,sobienlai,ngay,mabn,maql,idkhoa,makp,hoten,namsinh,diachi,loai,loaibn,lanin,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_VPKHOA", "id,mabn,maql,idkhoa,ngay,makp,madoituong,mavp,soluong,dongia,done,userid,ngayud,vattu" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_HUYBIENLAI", "id,tables,loai,mabn,hoten,makp,ngay,quyenso,sobienlai,lydo,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_TTRVTHUE", "id,quyenso,sobienlai,sotien,lanin" });
            ads.Tables[0].Rows.Add(new string[] { "VP", "V_DACHIHOAN", "idttrv,sophieuchi,ngaychi,user_thu,userid_chi,dachi" });
            ads.Tables[0].Rows.Add(new string[] { "D", "BHYTCLS", "id,stt,mavp,soluong,dongia,idchidinh,sttra,sobienlai" });
            ads.Tables[0].Rows.Add(new string[] { "D", "BHYTDS", "mabn,hoten,namsinh,diachi" });
            ads.Tables[0].Rows.Add(new string[] { "D", "BHYTKB", "mabn,maql,makp,chandoan,maicd,mabs,sothe,maphu,mabv,congkham,thuoc,cls,bntra,bhyttra,mmyy,userid,ngayud,done,sotoa,id,nhom,quyenso,sobienlai,ngay" });
            ads.Tables[0].Rows.Add(new string[] { "D", "BHYTTHUOC", "id,stt,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,cachdung" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_BIENLAI", "id,sohieu,sobienlai,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_BUCSTT", "id,stt,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,sttduyet" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CAPSOTOA", "id,ngay,loai,sotoa" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CHANDOAN", "id,loai,maicd,chandoan" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CHUYENCT", "id,stt,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,nguonchuyen,stttchuyen" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CHUYENLL", "id,nhom,sophieu,ngay,lydo,mmyy,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CTGHISOCT", "id,makho,makp,no,co,sotien,diengiai" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_CTGHISOLL", "id,nhom,soct,ngay,mmyy,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_DUTRUCAPCT", "id,manguon,mabd,slyeucau,slthuc" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_DUTRUCAPLL", "id,nhom,sophieu,ngay,loai,khox,khon,userid,ngayud,done" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_ERROR", "message,computer,tables,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_HUYBANCT", "id,stt,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_HUYBANLL", "id,nhom,mabn,hoten,namsinh,ngay,mabs,makp,loai,mmyy,done,userid,ngayud,lanin,sotoa,maql" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_KIEMTRA", "nhom,ngay,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NGTRUCT", "id,stt,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NGTRULL", "id,nhom,mabn,hoten,namsinh,ngay,mabs,makp,loai,mmyy,done,userid,ngayud,lanin,sotoa,maql" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NHAPCT", "vat,soluong,dongia,sotien,giaban,giamua,sl1,sl2,tyle,cuocvc,chaythu,namsx,tailieu,baohanh,nguongoc,tinhtrang,sothe,giabancu,id,stt,mabd,handung,losx" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NHAPCT2", "id,stt,mabd,soluong,sotien" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NHAPLL", "id,nhom,sophieu,ngaysp,sohd,ngayhd,bbkiem,ngaykiem,loai,nguoigiao,madv,makho,manguon,nhomcc,no,co,mmyy,userid,ngayud,paid,lydo" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NHAPSLCT", "id,stt,mabd,handung,losx,soluong,sl1,sl2" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_NHAPSLLL", "id,nhom,sophieu,ngaysp,sohd,ngayhd,loai,nguoigiao,madv,makho,mmyy,done,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_PHIEUXUAT", "id,soct,ngay,makp,nhom,loai,phieu,kho,sotien,no,co,diengiai,mmyy,userid,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_SOPHIEU", "mmyy,nhom,ngay,makp,loai,phieu,loaiin,makho,madoituong,so,lanin,manguon,nguoilinh" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THANHTOAN", "ngay,no,co,sotien,datra,mmyy,userid,ngayud,id" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THEODOI", "id,mabd,manguon,nhomcc,handung,losx,sothe,namsx,namsd,baohanh,nguongoc,tinhtrang,giamua,giaban,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THEODOIGIA", "mabd,ngay,dongia,ngayud" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THEODOITSCD", "makho,sttt,nam,namsx,namsd,tyle,phanloai,baohanh,nguongoc,tinhtrang,sothe" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THUCBUCSTT", "id,sttt,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,stt" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THUCXUAT", "id,sttt,madoituong,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,sothe,namsx,namsd,stt" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_THUCXUAT2", "id,sttt,makho,mabd,soluong,sotien,giamua,stt" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TIENTHUOC", "mabn,maql,idkhoa,ngay,makp,madoituong,mabd,soluong,sotien,giaban,giamua,done" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TONKHOCT", "mmyy,makho,manguon,nhomcc,stt,idn,sttn,mabd,handung,losx,tondau,sttondau,slnhap,stnhap,slxuat,stxuat,giaban,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TONKHOKEMTHEO", "makhot,sttt,makho,stt,idn,sttn,mabd,tondau,sttondau,slnhap,stnhap,slxuat,stxuat,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TONKHOTH", "mmyy,makho,mabd,tondau,slnhap,slxuat,slyeucau,manguon" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TUTRUCCT", "mmyy,makp,makho,manguon,nhomcc,stt,mabd,handung,losx,tondau,sttondau,slnhap,stnhap,slxuat,stxuat,giaban,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TUTRUCKEMTHEO", "makhot,sttt,makp,makho,stt,mabd,tondau,sttondau,slnhap,stnhap,slxuat,stxuat,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_TUTRUCTH", "mmyy,makp,makho,mabd,tondau,slnhap,slxuat,slyeucau,manguon" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_XUATCT", "sothe,mabs,hotenbn,id,stt,sttt,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_XUATLL", "id,nhom,sophieu,ngay,loai,khox,khon,lydo,mmyy,userid,ngayud,idduyet" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_XUATSDCT", "id,stt,sttt,madoituong,makho,manguon,nhomcc,mabd,handung,losx,soluong,sotien,giaban,giamua,sttduyet,sothe,namsx,namsd" });
            ads.Tables[0].Rows.Add(new string[] { "D", "D_XUATSDLL", "id,nhom,mabn,maql,idkhoa,ngay,loai,phieu,makp,mmyy,userid,ngayud,idduyet,thuoc,makhoa,lydo,lk,ghichu" });
            return ads;
        }
        public string s_user_vp(string v_mmyy_only)
        {
            return (user + "" + v_mmyy_only);
        }
        public string s_user_d(string v_mmyy)
        {
            return (user + "d" + v_mmyy);
        }
        public string get_field_vp(string v_table)
        {
            string rt = "";
            try
            {
                rt = dsfield.Tables[0].Select("table_name='" + v_table.ToUpper() + "' and loai='VP'")[0]["field_name"].ToString();
            }
            catch
            {
                rt = "*";
            }
            if (rt.Trim() == "") rt = "*";
            return rt;
        }
        public bool s_iscreated(string v_user)
        {
            return (get_data("select * from m_table where mmyy='" + v_user.Substring(v_user.Length - 4) + "'").Tables[0].Rows.Count > 0);
        }
        #region ana

        public bool upd_ana_BANGDULIEUVIENPHI_API(string LOAIPHIEU, string MASO, string GIOPHUT, string COMPUTERNAME, string HOTEN, string TENKHOA, string MABN, string TENBN, string NAMSINH, string DOITUONG, string QUYENSO,
            string SOBIENLAI, string SOCHUNGTU, decimal TONGCONG, decimal BHXH, decimal TONGTAMUNG, decimal HOANUNG, decimal MIEN, decimal THUCTHU, decimal SOTIEN1, decimal SOTIEN2, decimal SOTIEN3, decimal SOTIEN4, decimal SOTIEN5, decimal SOTIEN6, 
            decimal SOTIEN7, decimal SOTIEN8, decimal SOTIEN9, decimal SOTIEN10, decimal SOTIEN11, decimal SOTIEN12, decimal SOTIEN13, decimal SOTIEN14, decimal SOTIEN15, string NGAYCT, decimal STT, string CUATHU, string SOHOADON, string SOHOADONCHUAN)
        {
            //sql = "update BANGDULIEUVIENPHI_TAM set LOAIPHIEU=N'" + LOAIPHIEU + "',MASO='"+ MASO + "',GIOPHUT='"+ GIOPHUT + "',COMPUTERNAME='"+ COMPUTERNAME + "',HOTEN=N'" + HOTEN + "',TENKHOA=N'" + TENKHOA + "',MABN='" + MABN + "',TENBN=N'" + TENBN + "',NAMSINH='" + NAMSINH + "',DOITUONG='" + DOITUONG + "',QUYENSO='" + QUYENSO + "',";
            //sql += "SOBIENLAI='" + SOBIENLAI + "',TONGCONG=" + TONGCONG + ",BHXH=" + BHXH + ",HOANUNG=" + HOANUNG + ",MIEN=" + MIEN + ",THUCTHU=" + THUCTHU + ",SOTIEN1=" + SOTIEN1 + ",SOTIEN2=" + SOTIEN2 + ",SOTIEN3=" + SOTIEN3 + ",SOTIEN4=" + SOTIEN4 + ",SOTIEN5=" + SOTIEN5 + ",SOTIEN6=" + SOTIEN6 + ",";
            //sql += "SOTIEN7=" + SOTIEN7 + ",SOTIEN8=" + SOTIEN8 + ",SOTIEN9=" + SOTIEN9 + ",SOTIEN10=" + SOTIEN10 + ",SOTIEN11=" + SOTIEN11 + ",SOTIEN12=" + SOTIEN12 + ",SOTIEN13=" + SOTIEN13 + ",SOTIEN14=" + SOTIEN14 + ",SOTIEN15=" + SOTIEN15 + ",TIT1=N'" + ANA_TIT1 + "',TIT2=N'" + ANA_TIT2 + "',TIT3=N'" + ANA_TIT3 + "',TIT4=N'" + ANA_TIT4 + "',TIT5=N'" + ANA_TIT5 + "',";
            //sql += "TIT6=N'" + ANA_TIT6 + "',TIT7=N'" + ANA_TIT7 + "',TIT8=N'" + ANA_TIT9 + "',TIT10=N'" + ANA_TIT10 + "',TIT11=N'" + ANA_TIT11 + "',TIT12=N'" + ANA_TIT12 + "',TIT13=N'" + ANA_TIT13 + "',TIT14=N'" + ANA_TIT14 + "',TIT15=N'" + ANA_TIT15 + "',NGAYCT='" + NGAYCT + "',STT=" + STT + ",CUATHU=N'" + CUATHU + "',SOHOADON='" + SOHOADON + "',SOHOADONCHUAN='" + SOHOADONCHUAN + "'";
            //sql += " where SOCHUNGTU='" + SOCHUNGTU + "'";
            //if (!thucThiSql(sql,apiExcuteMsSql))
            //{

                sql = "INSERT INTO BANGDULIEUVIENPHI_TAM(LOAIPHIEU,MASO,GIOPHUT,COMPUTERNAME,HOTEN,TENKHOA,MABN,TENBN,NAMSINH,DOITUONG,QUYENSO,SOBIENLAI,SOCHUNGTU,TONGCONG,BHXH,TONGTAMUNG,HOANUNG,MIEN,THUCTHU";
                sql += ",SOTIEN1,SOTIEN2,SOTIEN3,SOTIEN4,SOTIEN5,SOTIEN6,SOTIEN7,SOTIEN8,SOTIEN9,SOTIEN10,SOTIEN11,SOTIEN12,SOTIEN13,SOTIEN14,SOTIEN15,NGAYCT,STT,CUATHU,SOHOADON,SOHOADONCHUAN,TIT1,TIT2,TIT3,TIT4,TIT5,TIT6,TIT7,TIT8,TIT9,TIT10,TIT11,TIT12,TIT13,TIT14,TIT15)";
                sql += " values (N'" + LOAIPHIEU + "','" + MASO + "','" + GIOPHUT + "','" + COMPUTERNAME + "',N'" + HOTEN + "',N'" + TENKHOA + "','" + MABN + "',N'" + TENBN + "','" + NAMSINH + "','" + DOITUONG + "','" + QUYENSO + "','" + SOBIENLAI + "','" + SOCHUNGTU + "'";
                sql += "," + TONGCONG + "," + BHXH + ","+ TONGTAMUNG + "," + HOANUNG + "," + MIEN + "," + THUCTHU + "," + SOTIEN1 + "," + SOTIEN2 + "," + SOTIEN3 + "," + SOTIEN4 + "," + SOTIEN5 + "," + SOTIEN6 + "," + SOTIEN7 + "," + SOTIEN8 + "," + SOTIEN9 + "," + SOTIEN10 + "," + SOTIEN11 + "," + SOTIEN12 + "," + SOTIEN13 + "," + SOTIEN14 + "," + SOTIEN15 + ",";
                sql += "convert(datetime, '" + NGAYCT + "', 105)," + STT + ",N'" + CUATHU + "','" + SOHOADON + "','" + SOHOADONCHUAN + "'";
                sql += ",N'" + ANA_TIT1 + "',N'" + ANA_TIT2 + "',N'" + ANA_TIT3 + "',N'" + ANA_TIT4 + "',N'" + ANA_TIT5 + "',N'" + ANA_TIT6 + "',N'" + ANA_TIT7 + "',N'" + ANA_TIT8 + "',N'" + ANA_TIT9 + "',N'" + ANA_TIT10 + "',N'" + ANA_TIT11 + "',N'" + ANA_TIT12 + "',N'" + ANA_TIT13 + "',N'" + ANA_TIT14 + "',N'" + ANA_TIT15 + "')";
                if (thucThiSql(sql, apiExcuteMsSql))
                    return true;
                else
                    return false;
            //}

            //else
            //    return true;
        }
        public bool upd_ana_BANGDULIEUKYQUY_API(string LOAIPHIEU, string MASO, string GIOPHUT, string COMPUTERNAME, string HOTEN, string MAKHOA, string TENKHOA, string MABN, string TENBN, string NAMSINH, string DOITUONG, string QUYENSO, string SOBIENLAI, decimal TONGCONG, decimal SOTIENHUY, string NGAYCT, decimal STT, string NGAYHOADON, string CUATHU, string thoigianhoantra, int dahoantra)
        {
            sql = "insert into BANGDULIEUKYQUY_TAM (LOAIPHIEU,MASO,GIOPHUT,COMPUTERNAME,HOTEN,MAKHOA,TENKHOA,MABN,TENBN,NAMSINH,DOITUONG,QUYENSO,SOBIENLAI,TONGCONG,SOTIENHUY,NGAYCT,STT,NGAYHOADON,CUATHU,thoigianhoanquy,dahoanquy)";
            sql += "values (N'" + LOAIPHIEU + "','" + MASO + "','" + GIOPHUT + "','" + COMPUTERNAME + "',N'" + HOTEN + "','" + MAKHOA + "',N'" + TENKHOA + "','" + MABN + "',N'" + TENBN + "','" + NAMSINH + "',N'" + DOITUONG + "','" + QUYENSO + "','" + SOBIENLAI + "'," + TONGCONG + "," + SOTIENHUY + ",'" + NGAYCT + "'," + STT + ",'" + NGAYHOADON + "','" + CUATHU + "','" + thoigianhoantra + "'," + dahoantra + ")";
            if (thucThiSql(sql, apiExcuteMsSql))
                return true;
            else
            {
                ghiloi(sql);
                return false;
            }
        }
         public bool upd_ana_BANGDULIEUBAOHIEM_API(string IDMED, decimal STT, string LOAIPHIEU, string GIOPHUT, string NGAYTHANHTOAN, string NGAYDULIEU, string HOTEN, 
             string MAKHOA, string TENKHOA, string MATHE, string MADKBD, string NGAYDAUTHE, string NGAYCUOITHE, decimal TYLETHE, string MAICD, string MABN, 
             string TENBN, string NAMSINHNAM, string NAMSINHNU, string NAMSINH, string DIACHI, string NGAYCT, string NGAYVAO, string NGAYRA, decimal SONGAY, 
             decimal TONGCONG, decimal TONGBNTRA, decimal TONGBHTRA, decimal TIENXETNGHIEM, decimal TIENCDHA, decimal TIENTHUOC, decimal TIENMAU, decimal TIENTTPT, 
             decimal TIENVTYT, decimal TIENKHAM, decimal TIENGIUONG, decimal TIENVANCHUYEN, decimal TIENXETNGHIEMBH, decimal TIENCDHABH, decimal TIENTHUOCBH, 
             decimal TIENMAUBH, decimal TIENTTPTBH, decimal TIENVTYTBH, decimal TIENKHAMBH, decimal TIENGIUONGBH, decimal TIENVANCHUYENBH,string NAM, string THANG)
        {
            sql = "insert into BANGDULIEUBAOHIEM_TAM(IDMED, STT, LOAIPHIEU, GIOPHUT, NGAYTHANHTOAN, NGAYDULIEU, HOTEN, MAKHOA, TENKHOA, MATHE, MADKBD, NGAYDAUTHE, NGAYCUOITHE, TYLETHE,"
                + "MAICD, MABN, TENBN, NAMSINHNAM, NAMSINHNU, NAMSINH, DIACHI, NGAYCT, NGAYVAO, NGAYRA, SONGAY, TONGCONG, TONGBNTRA, TONGBHTRA, TIENXETNGHIEM, "
                + "TIENCDHA, TIENTHUOC, TIENMAU, TIENTTPT, TIENVTYT, TIENKHAM, TIENGIUONG, TIENVANCHUYEN, TIENXETNGHIEMBH, TIENCDHABH, TIENTHUOCBH, TIENMAUBH, "
                + "TIENTTPTBH, TIENVTYTBH, TIENKHAMBH, TIENGIUONGBH, TIENVANCHUYENBH,NAM,THANG) ";
            sql += " values (" + IDMED + "," + STT + ",N'" + LOAIPHIEU + "','" + GIOPHUT + "','" + NGAYTHANHTOAN + "',convert(datetime, '" + NGAYDULIEU + "', 105),"
                + "N'" + HOTEN + "','" + MAKHOA + "',N'" + TENKHOA + "','" + MATHE + "','" + MADKBD + "','" + NGAYDAUTHE + "',"
                + "'" + NGAYCUOITHE + "'," + TYLETHE + ",'" + MAICD + "','" + MABN + "',N'" + TENBN + "','" + NAMSINHNAM + "','" + NAMSINHNU + "',"
                + "'" + NAMSINH + "',N'" + DIACHI + "',convert(datetime, '" + NGAYCT + "', 105), '" + NGAYVAO + "','" + NGAYRA + "', " + SONGAY + ", " + TONGCONG + ", " + TONGBNTRA + ", " + TONGBHTRA + ", " + TIENXETNGHIEM
                + ", " + TIENCDHA + ", " + TIENTHUOC + ", " + TIENMAU + ", " + TIENTTPT + ", " + TIENVTYT + ", " + TIENKHAM + ", " + TIENGIUONG + ", " + TIENVANCHUYEN + ", " + TIENXETNGHIEMBH + ", " + TIENCDHABH + ", " + TIENTHUOCBH + ", " + TIENMAUBH + ", " + TIENTTPTBH + ", " + TIENVTYTBH + ", " + TIENKHAMBH + ", " + TIENGIUONGBH + ", " + TIENVANCHUYENBH + ",'" + NAM + "','" + THANG + "')";
            sql += "";
            if (thucThiSql(sql, apiExcuteMsSql))
                return true;
            else
                return false;
        }

        #endregion
    }
}