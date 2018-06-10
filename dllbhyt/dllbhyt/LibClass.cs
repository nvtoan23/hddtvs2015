namespace dllbhyt
{
    using System;
    using System.Data;
    using System.Data.OleDb;
    using System.Data.OracleClient;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Windows.Forms;
    using System.Xml;

    public class LibClass
    {
        private OracleCommand cmd;
        private OracleConnection con;
        private OracleDataAdapter dest;
        private DataSet ds = new DataSet();
        public const string links_pass = "715501920";
        public const string links_userid = "links";
        public string Msg = "B\x00e1o c\x00e1o BHYT";
        private string sComputer = null;
        private string sConn = "Data Source=medisoft;user id=medibv;password=medibv";
        private string service_name = "medisoft";
        private string sql = "";
        private string userid = "medibv";

        public LibClass()
        {
            if (this.Maincode("Con") != "")
            {
                this.sConn = this.Maincode("Con");
            }
            this.sComputer = Environment.MachineName.Trim().ToUpper();
            this.userid = this.sConn.Substring(this.sConn.LastIndexOf("=") + 1).Trim();
            this.service_name = this.sConn.Substring(this.sConn.IndexOf("=") + 1, (this.sConn.IndexOf(";") - 1) - this.sConn.IndexOf("=")).Trim();
        }

        public bool bcongkham_bhyt()
        {
            this.ds = this.get_data("select ten from d_thongso where id=82 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return false;
            }
            return (this.ds.Tables[0].Rows[0][0].ToString() == "1");
        }

        public decimal BHYT_traituyen()
        {
            this.ds = this.get_data("select ten from thongso where id=482 ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0M;
            }
            return decimal.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public bool bMaicd(string s)
        {
            return ((s == "") || (this.get_data("select cicd10 from icd10 where cicd10='" + s + "'").Tables[0].Rows.Count > 0));
        }

        public bool bMmyy(string m_mmyy)
        {
            try
            {
                return (this.get_data("select * from d_table where mmyy='" + ((m_mmyy.Trim().Length == 4) ? m_mmyy : this.mmyy(m_mmyy)) + "'").Tables[0].Rows.Count > 0);
            }
            catch
            {
                return false;
            }
        }

        public bool bNgay(string ngay)
        {
            try
            {
                if (ngay.IndexOf("_") != -1)
                {
                    return false;
                }
                int length = ngay.Length;
                if (length == 0)
                {
                    return false;
                }
                string s = ngay.Substring(0, 2);
                string str2 = ngay.Substring(3, 2);
                string str3 = ngay.Substring(6, 4);
                string str4 = "01+03+05+07+08+10+12+";
                string str5 = "04+06+09+11+";
                string str6 = ((int.Parse(str3) % 4) == 0) ? "29" : "28";
                if (int.Parse(str3.Substring(0, 1)) < 1)
                {
                    return false;
                }
                if ((int.Parse(str2) < 1) || (int.Parse(str2) > 12))
                {
                    return false;
                }
                if (str4.IndexOf(str2 + "+") > -1)
                {
                    if ((int.Parse(s) < 1) || (int.Parse(s) > 0x1f))
                    {
                        return false;
                    }
                }
                else if (str5.IndexOf(str2 + "+") > -1)
                {
                    if ((int.Parse(s) < 1) || (int.Parse(s) > 30))
                    {
                        return false;
                    }
                }
                else if ((int.Parse(s) < 1) || (int.Parse(s) > int.Parse(str6)))
                {
                    return false;
                }
                if (length > 10)
                {
                    string str7 = ngay.Substring(11, 2);
                    string str8 = ngay.Substring(14, 2);
                    if (int.Parse(str7) > 0x17)
                    {
                        return false;
                    }
                    if (int.Parse(str8) > 0x3b)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void check_process_Excel()
        {
            try
            {
                Process[] processes = Process.GetProcesses();
                if (processes.Length > 1)
                {
                    int num = 0;
                    for (int i = 0; i <= (processes.Length - 1); i++)
                    {
                        if (processes[i].ProcessName == "EXCEL")
                        {
                            num++;
                            processes[i].Kill();
                        }
                    }
                }
            }
            catch
            {
            }
        }

        public decimal Congkham()
        {
            this.ds = this.get_data("select ten from d_thongso where id=47 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 3000M;
            }
            return decimal.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public int d_dongia_le(int d_nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=58 and nhom=" + d_nhom);
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 2;
            }
            return int.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public int d_giaban_le(int d_nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=79 and nhom=" + d_nhom);
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0;
            }
            return int.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public int d_soluong_le(int d_nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=57 and nhom=" + d_nhom);
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 2;
            }
            return int.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public int d_thanhtien_le(int d_nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=59 and nhom=" + d_nhom);
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 2;
            }
            return int.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public string DateToString(string format, DateTime date)
        {
            if (date.Equals(null))
            {
                return "";
            }
            return date.ToString(format, DateTimeFormatInfo.CurrentInfo);
        }

        public void delrec(DataTable dt, string exp)
        {
            try
            {
                DataRow[] rowArray = dt.Select(exp);
                for (int i = 0; i < rowArray.Length; i++)
                {
                    rowArray[i].Delete();
                }
            }
            catch
            {
            }
        }

        public void execute_data(string sql)
        {
            try
            {
                this.con = new OracleConnection(this.sConn);
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                this.con.Close();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim(), this.sComputer, "?");
            }
            finally
            {
                this.cmd.Dispose();
                this.con.Dispose();
            }
        }

        public void execute_data_d(string mmyy, string sql)
        {
            try
            {
                this.con = new OracleConnection(this.get_conn_d(mmyy));
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                this.con.Close();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim(), this.sComputer, "?");
            }
            finally
            {
                this.cmd.Dispose();
                this.con.Dispose();
            }
        }

        public void execute_data_mmyy(string str, string tu, string den)
        {
            DateTime time = this.StringToDate(tu);
            DateTime time2 = this.StringToDate(den);
            int year = time.Year;
            int month = time.Month;
            int num3 = time2.Year;
            int num4 = time2.Month;
            string str2 = "";
            for (int i = year; i <= num3; i++)
            {
                int num5 = (i == year) ? month : 1;
                int num6 = (i == num3) ? num4 : 12;
                for (int j = num5; j <= num6; j++)
                {
                    str2 = j.ToString().PadLeft(2, '0') + i.ToString().Substring(2, 2);
                    if (this.bMmyy(str2))
                    {
                        this.sql = str.Replace("xxx", this.user + str2);
                        this.execute_data(this.sql);
                    }
                }
            }
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
            catch (Exception exception)
            {
                this.upd_error(exception.Message, this.sComputer, tenfile);
                return "";
            }
        }

        public string Export_Excel(DataTable dt, string tenfile)
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
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    str3 = (str3 + "<th>") + dt.Columns[i].ColumnName + "</th>";
                }
                str3 = str3 + "</tr>";
                writer.Write(str3);
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    str3 = "<tr>";
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        str3 = (str3 + "<td>") + dt.Rows[j][k].ToString() + "</td>";
                    }
                    str3 = str3 + "</tr>";
                    writer.Write(str3);
                }
                str3 = "</Table>";
                writer.Write(str3);
                writer.Close();
                return str2;
            }
            catch (Exception exception)
            {
                this.upd_error(exception.Message, this.sComputer, tenfile);
                return "";
            }
        }

        public string Export_Excel(DataSet dset, string tenfile, string s_dieukien, string s_FieldSort)
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
                foreach (DataRow row in dset.Tables[0].Select(s_dieukien, s_FieldSort))
                {
                    str3 = "<tr>";
                    for (int j = 0; j < dset.Tables[0].Columns.Count; j++)
                    {
                        str3 = (str3 + "<td>") + row[j].ToString() + "</td>";
                    }
                    str3 = str3 + "</tr>";
                    writer.Write(str3);
                }
                str3 = "</Table>";
                writer.Write(str3);
                writer.Close();
                return str2;
            }
            catch (Exception exception)
            {
                this.upd_error(exception.Message, this.sComputer, tenfile);
                return "";
            }
        }

        public void f_Capnhat_Cautruc()
        {
            try
            {
                this.f_taocautruc();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public void f_end_process_Excel(string idprocess)
        {
            try
            {
                Process[] processesByName = Process.GetProcessesByName("EXCEL");
                if (processesByName.Length > 1)
                {
                    string[] strArray = idprocess.Split(new char[] { ',' });
                    for (int i = 0; i < strArray.Length; i++)
                    {
                        for (int j = 0; j <= (processesByName.Length - 1); j++)
                        {
                            if (processesByName[j].Id.ToString() == strArray[i])
                            {
                                processesByName[j].Kill();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        public void f_end_process_Excel(string idprocessfirst, string idprocesslast)
        {
            if (idprocesslast.Length > 0)
            {
                idprocesslast = "," + idprocesslast + ",";
            }
            string[] strArray = idprocessfirst.Split(new char[] { ',' });
            for (int i = 0; i < strArray.Length; i++)
            {
                if (idprocesslast.IndexOf("," + strArray[i] + ",") > -1)
                {
                    idprocesslast = idprocesslast.Replace("," + strArray[i] + ",", ",");
                }
            }
            if (idprocesslast.Trim(new char[] { ',' }).Length > 0)
            {
                this.f_end_process_Excel(idprocesslast.Trim(new char[] { ',' }));
            }
        }

        public DataSet f_get_dskhoaphong()
        {
            string user = this.user;
            string sql = "select makp,tenkp,makp_byt from " + user + ".btdkp_bv ";
            this.ds = this.get_data(sql);
            return this.ds;
        }

        public DataSet f_get_dsthuoc_cls()
        {
            string user = this.user;
            string sql = "";
            string str3 = "";
            string columnName = "nhathau";
            try
            {
                columnName = this.get_data("select thongtinthau from v_giavp where 1=0").Tables[0].Columns[0].ColumnName;
            }
            catch
            {
            }
            sql = "select a.id,cast(a.ma as varchar2(50)) as ma,a.ten,a.dang as dvt,a.kythuat,a.bhyt,e.idnhombhyt as nhombhyt,0 as loaivp,case when a.vattuthaythe=1 then 14 else h.idnhombhytmedisoft end as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,1 as thuoc,a.hamluong,a.sodk,a.maduongdung as duongdung,a.duongdung as tenduongdung,a.lieuluong as lieuluong,a.mavattubyt as mavattu,a.masobyt as mabyt,a.tenbyt,e.mabhyt2348 as maloaibhyt,0 as gia_bh,a.dvtbyt,a.donvi,a.gia_bh_toida,a.ma_giuong_byt," + columnName + " nhathau from " + user + ".d_dmbd a inner join " + user + ".d_dmnhom b on a.manhom=b.id inner join " + user + ".v_nhomvp e on b.nhomvp=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id ";
            str3 = "  select c.id,cast(c.ma as varchar2(50)) as ma,cast(c.ten as nvarchar2(2000)) as ten,c.dvt,c.kythuat,c.bhyt,e.idnhombhyt as nhombhyt,d.id as loaivp,case when c.vattuthaythe=1 then 14 else h.idnhombhytmedisoft end as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,0 as thuoc,null as hamluong,null as sodk,null as duongdung,null as tenduongdung,null as lieuluong,c.mavattubyt as mavattu,c.masobyt as mabyt,c.tenbyt,d.masobyt as maloaibhyt,c.gia_bh,dvt as dvtbyt,null as donvi,c.gia_bh_toida,c.ma_giuong_byt," + columnName + " as nhathau from " + user + ".v_giavp c inner join " + user + ".v_loaivp d on c.id_loai=d.id inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id ";
            this.execute_data("create or replace view vi_dmvpbhyt as " + (sql + " union all " + str3));
            this.ds = this.get_data("select * from vi_dmvpbhyt");
            if ((this.ds == null) || (this.ds.Tables.Count == 0))
            {
                try
                {
                    this.ds = this.get_data(sql);
                    this.ds.Merge(this.get_data(str3));
                }
                catch
                {
                }
            }
            return this.ds;
        }

        public int f_get_lydo_vaovien(string makp, bool traituyen, int nhantu)
        {
            int num = 1;
            try
            {
                if (nhantu == 1)
                {
                    return 2;
                }
                if (traituyen)
                {
                    num = 3;
                }
            }
            catch
            {
            }
            return num;
        }

        public string f_get_mabvchuyenden(string mabn, string mavaovien)
        {
            string str = ";";
            string sql = "";
            string str3 = mavaovien.Substring(2, 2) + mavaovien.Substring(0, 2);
            string str5 = "select * from (select distinct b.mabv from " + this.user + str3 + ".tiepdon a1 inner join " + this.user + str3 + ".noigioithieu a on a1.maql=a.maql inner join dmnoicapbhyt b on a.mabv=b.mabvmedi WHERE a1.mabn='" + mabn + "' and a1.mavaovien=" + mavaovien;
            sql = str5 + " union all select distinct b.mabv from " + this.user + str3 + ".tiepdon a1 inner join noigioithieu a on a1.maql=a.maql inner join dmnoicapbhyt b on a.mabv=b.mabvmedi WHERE a1.mabn='" + mabn + "' and a1.mavaovien=" + mavaovien + ") b order by b.mabv desc";
            try
            {
                str = this.get_data(sql).Tables[0].Rows[0]["mabv"].ToString();
            }
            catch
            {
            }
            return str.Trim(new char[] { ';' });
        }

        public DataSet f_get_sothebhyt_tungay(string sothe, string ngayvao, string ngayra)
        {
            string str = "";
            if (ngayvao.IndexOf('/') == -1)
            {
                ngayvao = ngayvao.Substring(6, 2) + "/" + ngayvao.Substring(4, 2) + "/" + ngayvao.Substring(0, 4);
                ngayra = ngayra.Substring(6, 2) + "/" + ngayra.Substring(4, 2) + "/" + ngayra.Substring(0, 4);
            }
            DataSet set = new DataSet();
            string str2 = this.user + str;
            DateTime time = new DateTime(int.Parse(ngayra.Substring(6, 4)), int.Parse(ngayra.Substring(3, 2)), int.Parse(ngayra.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time3 = new DateTime(int.Parse(ngayvao.Substring(6, 4)), int.Parse(ngayvao.Substring(3, 2)), 1).AddMonths(-1);
            string str4 = "";
            while (time3 <= time)
            {
                str = time3.ToString("MMyy");
                if (this.bMmyy(str))
                {
                    str2 = this.user + str;
                    string str5 = str4;
                    str4 = str5 + "#select to_char(tungay,'dd/mm/yyyy') tungay,to_char(denngay,'dd/mm/yyyy') denngay from " + str2 + ".bhyt  where sothe='" + sothe + "'# select to_char(tungay,'dd/mm/yyyy') tungay,to_char(denngay,'dd/mm/yyyy') denngay from bhyt  where sothe='" + sothe + "'";
                }
                time3 = time3.AddMonths(1);
            }
            return this.get_data("select * from (" + str4.Trim(new char[] { '#' }).Replace("#", " union all ") + ") a order by a.tungay desc");
        }

        public DataTable f_getdata_fromfileExcel(string filePath, string sheetname)
        {
            DataSet dataSet = new DataSet();
            string[] strArray = filePath.Split(new char[] { '.' });
            string connectionString = "";
            if ((filePath.Length > 1) && (strArray[1] == "xls"))
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=Excel 8.0;";
            }
            else
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;";
            }
            OleDbConnection selectConnection = new OleDbConnection(connectionString);
            try
            {
                selectConnection.Open();
                new OleDbDataAdapter("select * from [" + sheetname + "$]", selectConnection).Fill(dataSet);
            }
            catch
            {
                return null;
            }
            finally
            {
                selectConnection.Close();
            }
            if (dataSet != null)
            {
                try
                {
                    dataSet.WriteXml("dsexcel.xml");
                }
                catch
                {
                }
            }
            return dataSet.Tables[0];
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

        public string f_modify(string tenfile)
        {
            return File.GetLastWriteTime(tenfile).ToString("dd/MM/yyyy HH:mm");
        }

        public string f_size(string tenfile)
        {
            FileInfo info = new FileInfo(tenfile);
            long num = info.Length / 0x400L;
            return num.ToString();
        }

        private void f_taocautruc()
        {
            try
            {
                this.sql = "create table " + this.user + ".dmnhomdv_bhyt(manhom number(3),nhombv varchar2(8),tennhom nvarchar2(200),mabv varchar2(8),tenbv nvarchar2(200),stt number(3),constraint pk_dmnhomdv_bhyt primary key(manhom,stt));\n";
                this.sql = this.sql + "create table dmbaocao_bhyt(id number(3),stt number(3),ten nvarchar2(100),idnhomvp number(3),idloaivp number(3),constraint pk_dmbaocao_bhyt primary key(id));\n";
                this.sql = this.sql + "create table doituongth_bhyt(id number(4),stt number(4),math nvarchar2(100),tenth nvarchar2(100),maloai number(1),tenloai nvarchar2(100),ghichu nvarchar2(100),constraint pk_doituongth_bhyt primary key(id,math));\n";
                this.sql = this.sql + "alter table doituong_bhyt modify(ma13 nvarchar2(100),tenth nvarchar2(100));\n";
                this.sql = this.sql + "exit;";
                string path = @"..\..\..\xml\tao.sql";
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
                StreamWriter writer = new StreamWriter(stream);
                writer.Write(this.sql);
                writer.Close();
                string fileName = "sqlplus.exe";
                string arguments = this.user + "/" + this.user + "@" + this.service_name + " @" + path;
                new run(fileName, arguments, true).Launch();
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        public void f_taosolieu_bcbhyt_dmnhombhytct()
        {
            this.execute_data("create table bcbhyt_dmnhombhytct(id numeric(10) default 0,idvp numeric(10) default 0,ma varchar2(100),ten nvarchar2(200),idnhombhytmedisoft numeric(3) default 0,constraint pk_bcbhyt_dmnhombhytct primary key(id))");
        }

        public void f_taosolieu_bcbhyt_phanquyen()
        {
            this.execute_data("create table bcbhyt_phanquyen(id numeric(10) default 0,ten nvarchar2(100),constraint pk_bcbhyt_phanquyen primary key(id))");
        }

        public void f_taosolieu_bcbhyt_tt2348()
        {
            this.execute_data("alter table v_nhomvp add mabhyt2348 varchar2(20)");
            this.execute_data("alter table btdkp_bv add makp_byt varchar2(20)");
            this.execute_data("alter table ttxk add mabhyt2348 numeric(3) default 1");
            this.execute_data("alter table ketqua add mabhyt2348 numeric(3) default 1");
        }

        public void f_taosolieu_d_dmbd()
        {
            int count = 0;
            try
            {
                count = this.get_data("select mavattubyt from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add mavattubyt varchar2(20)");
            }
            try
            {
                count = this.get_data("select masobyt from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add masobyt varchar2(20)");
            }
            try
            {
                count = this.get_data("select lieuluong from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add lieuluong nvarchar2(1000)");
            }
            try
            {
                count = this.get_data("select tenbyt from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add tenbyt nvarchar2(1000)");
            }
            try
            {
                count = this.get_data("select maduongdung from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add maduongdung varchar2(20)");
            }
            try
            {
                count = this.get_data("select nhathau from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add nhathau nvarchar2(200)");
            }
            try
            {
                count = this.get_data("select nhathau from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add nhathau varchar2(100)");
            }
            try
            {
                count = this.get_data("select soquyetdinh from d_dmbd where 1=0").Tables[0].Rows.Count;
                this.execute_data("alter table d_dmbd modify soquyetdinh nvarchar2(100)");
            }
            catch
            {
                this.execute_data("alter table d_dmbd add soquyetdinh varchar2(100)");
            }
            try
            {
                count = this.get_data("select ngaycongboqd from d_dmbd where 1=0").Tables[0].Rows.Count;
                this.execute_data("alter table d_dmbd modify ngaycongboqd nvarchar2(100)");
            }
            catch
            {
                this.execute_data("alter table d_dmbd add ngaycongboqd date");
            }
            try
            {
                count = this.get_data("select dvtbyt from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add dvtbyt nvarchar2(50)");
            }
            try
            {
                count = this.get_data("select donggoibyt from d_dmbd where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add donggoibyt nvarchar2(50)");
            }
            try
            {
                int num2 = this.get_data("select gia_bh_toida from d_dmbd where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add gia_bh_toida number(15,2) default 0");
            }
            try
            {
                int num3 = this.get_data("select vattuthaythe from d_dmbd where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add vattuthaythe number(1) default 0");
            }
            try
            {
                int num4 = this.get_data("select ma_giuong_byt from d_dmbd where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table d_dmbd add ma_giuong_byt nvarchar2(100)");
            }
        }

        public void f_taosolieu_d_dmnhom()
        {
            this.execute_data("alter table d_dmnhom add thuocyhct numeric(1,0) default 0");
        }

        public void f_taosolieu_v_giavp()
        {
            int count = 0;
            try
            {
                count = this.get_data("select masobyt from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add masobyt varchar2(20)");
            }
            try
            {
                count = this.get_data("select madvkt from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add madvkt varchar2(20)");
            }
            try
            {
                count = this.get_data("select stttheothongtu from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add stttheothongtu varchar2(20)");
            }
            try
            {
                count = this.get_data("select sttbyt from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add sttbyt varchar2(20)");
            }
            try
            {
                count = this.get_data("select mavattubyt from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add mavattubyt varchar2(20)");
            }
            try
            {
                count = this.get_data("select tenbyt from v_giavp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add tenbyt nvarchar2(400)");
            }
            try
            {
                count = this.get_data("select masobyt from v_loaivp where 1=0").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_loaivp add masobyt varchar2(40)");
            }
            try
            {
                int num2 = this.get_data("select gia_bh_toida from v_giavp where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add gia_bh_toida number(15,2) default 0");
            }
            try
            {
                int num3 = this.get_data("select vattuthaythe from v_giavp where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add vattuthaythe number(1) default 0");
            }
            try
            {
                int num4 = this.get_data("select ma_giuong_byt from v_giavp where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this.execute_data("alter table v_giavp add ma_giuong_byt nvarchar2(100)");
            }
        }

        public void f_write_log(string mesg)
        {
            try
            {
                StreamWriter writer = new StreamWriter("log.txt", true);
                string str = "-------------------------------------\r\n" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "\r\n" + mesg;
                writer.WriteLine(str);
                writer.Close();
            }
            catch
            {
            }
        }

        public string file_exe(string tenfilegoc)
        {
            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory());
            for (int i = 0; i < files.GetLength(0); i++)
            {
                if ((files[i].ToString().ToUpper().IndexOf(".EXE") != -1) && (this.tenfile_goc(files[i].ToString()) == tenfilegoc.ToLower()))
                {
                    return files[i].ToString();
                }
            }
            return "";
        }

        public string format_dongia(int d_nhom)
        {
            string str = "###,###,###,##0";
            int num = this.d_dongia_le(d_nhom);
            if (num > 0)
            {
                str = str + ".";
            }
            for (int i = 0; i < num; i++)
            {
                str = str + "0";
            }
            return str;
        }

        public string format_giaban(int d_nhom)
        {
            string str = "###,###,###,##0";
            int num = this.d_giaban_le(d_nhom);
            if (num > 0)
            {
                str = str + ".";
            }
            for (int i = 0; i < num; i++)
            {
                str = str + "0";
            }
            return str;
        }

        public string format_soluong(int d_nhom)
        {
            string str = "###,###,###,##0";
            int num = this.d_soluong_le(d_nhom);
            if (num > 0)
            {
                str = str + ".";
            }
            for (int i = 0; i < num; i++)
            {
                str = str + "0";
            }
            return str;
        }

        public string format_sotien(int d_nhom)
        {
            string str = "###,###,###,##0";
            int num = this.d_thanhtien_le(d_nhom);
            if (num > 0)
            {
                str = str + ".";
            }
            for (int i = 0; i < num; i++)
            {
                str = str + "0";
            }
            return str;
        }

        private string get_conn(string mmyy)
        {
            return ("Data Source=" + this.service_name + ";user id=" + this.userid + "d" + mmyy + ";password=" + this.userid + "d" + mmyy);
        }

        private string get_conn_d(string mmyy)
        {
            return ("Data Source=" + this.service_name + ";user id=" + this.userid + "d" + mmyy + ";password=" + this.userid + "d" + mmyy);
        }

        private string get_conn_v(string mmyy)
        {
            return ("Data Source=" + this.service_name + ";user id=" + this.userid + mmyy + ";password=" + this.userid + mmyy);
        }

        public DataSet get_data(string sql)
        {
            try
            {
                this.con = new OracleConnection(this.sConn);
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.dest = new OracleDataAdapter(this.cmd);
                this.ds = new DataSet();
                this.dest.Fill(this.ds);
                this.cmd.Dispose();
                this.con.Close();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim() + "\n\r" + sql, this.sComputer, "?");
            }
            finally
            {
                this.ds.Dispose();
                this.cmd.Dispose();
                this.con.Dispose();
            }
            return this.ds;
        }

        public DataSet get_data(string mmyy, string sql)
        {
            try
            {
                this.con = new OracleConnection(this.get_conn(mmyy));
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.dest = new OracleDataAdapter(this.cmd);
                this.ds = new DataSet();
                this.dest.Fill(this.ds);
                this.cmd.Dispose();
                this.con.Close();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim() + "\n\r" + sql, this.sComputer, "?");
            }
            finally
            {
                this.ds.Dispose();
                this.cmd.Dispose();
                this.con.Dispose();
            }
            return this.ds;
        }

        public DataSet get_data_mmyy(string str, string tu, string den)
        {
            DataSet set = new DataSet();
            DateTime time = this.StringToDate(tu).AddDays((double) -this.iNgaykiemke);
            DateTime time2 = this.StringToDate(den).AddDays((double) this.iNgaykiemke);
            int year = time.Year;
            int month = time.Month;
            int num3 = time2.Year;
            int num4 = time2.Month;
            string str2 = "";
            bool flag = true;
            for (int i = year; i <= num3; i++)
            {
                int num5 = (i == year) ? month : 1;
                int num6 = (i == num3) ? num4 : 12;
                for (int j = num5; j <= num6; j++)
                {
                    str2 = j.ToString().PadLeft(2, '0') + i.ToString().Substring(2, 2);
                    if (this.bMmyy(str2))
                    {
                        this.sql = str.Replace("xxx", this.user + "d" + str2);
                        if (flag)
                        {
                            set = this.get_data(this.sql);
                            flag = false;
                        }
                        else
                        {
                            set.Merge(this.get_data(this.sql));
                        }
                    }
                }
            }
            return set;
        }

        public DataSet get_data_mmyy(string str, string tu, string den, bool khoangcach)
        {
            DataSet set = new DataSet();
            DateTime time = khoangcach ? this.StringToDate(tu).AddDays((double) -this.iNgaykiemke) : this.StringToDate(tu);
            DateTime time2 = khoangcach ? this.StringToDate(den).AddDays((double) this.iNgaykiemke) : this.StringToDate(den);
            int year = time.Year;
            int month = time.Month;
            int num3 = time2.Year;
            int num4 = time2.Month;
            string str2 = "";
            bool flag = true;
            for (int i = year; i <= num3; i++)
            {
                int num5 = (i == year) ? month : 1;
                int num6 = (i == num3) ? num4 : 12;
                for (int j = num5; j <= num6; j++)
                {
                    str2 = j.ToString().PadLeft(2, '0') + i.ToString().Substring(2, 2);
                    if (this.bMmyy(str2))
                    {
                        this.sql = str.Replace("xxx", this.user + str2);
                        if (flag)
                        {
                            set = this.get_data(this.sql);
                            flag = false;
                        }
                        else
                        {
                            set.Merge(this.get_data(this.sql));
                        }
                    }
                }
            }
            return set;
        }

        public DataSet get_data_v(string mmyy, string sql)
        {
            try
            {
                this.con = new OracleConnection(this.get_conn_v(mmyy));
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.dest = new OracleDataAdapter(this.cmd);
                this.ds = new DataSet();
                this.dest.Fill(this.ds);
                this.cmd.Dispose();
                this.con.Close();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim(), this.sComputer, "?");
            }
            return this.ds;
        }

        public DataTable get_dmnhombc_bhyt(int i_nhombc)
        {
            if (i_nhombc == 0)
            {
                this.sql = " select distinct b.ten,a.idnhombhytmedisoft as idnhomvp,'' as idloaivp, b.stt as stt,b.id from v_nhombhyt a inner join v_nhombhyt_medisoft b on a.idnhombhytmedisoft=b.id order by b.stt ";
            }
            else
            {
                this.sql = "select * from " + this.user + ".dmbaocao_bhyt order by stt ";
            }
            return this.get_data(this.sql).Tables[0];
        }

        public DataTable get_dmnhombc_bhyt_vp(int i_nhombc)
        {
            if (i_nhombc == 0)
            {
                this.sql = "select a.ten, a.id as idnhomvp, ' ' as idloaivp, a.stt, a.id from " + this.user + ".v_loaivp a order by stt ";
            }
            else
            {
                this.sql = "select * from " + this.user + ".dmbaocao_bhyt order by stt ";
            }
            return this.get_data(this.sql).Tables[0];
        }

        public int get_id_nhombhyt()
        {
            try
            {
                this.sql = "select nvl(max(manhom),0) id from dmnhomdv_bhyt";
                return (int.Parse(this.get_data(this.sql).Tables[0].Rows[0]["id"].ToString()) + 1);
            }
            catch
            {
                return 1;
            }
        }

        private int get_loai_doituong(int madoituong)
        {
            this.ds = this.get_data("select loai from d_doituong where madoituong=" + madoituong);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return int.Parse(this.ds.Tables[0].Rows[0]["loai"].ToString());
            }
            return 0;
        }

        public string get_sql_mmyy(string str, string tu, string den)
        {
            DataSet set = new DataSet();
            DateTime time = this.StringToDate(tu).AddDays((double) -this.iNgaykiemke);
            DateTime time2 = this.StringToDate(den).AddDays((double) this.iNgaykiemke);
            string str2 = "";
            this.sql = "";
            while (int.Parse(time.ToString("yyMM")) <= int.Parse(time2.ToString("yyMM")))
            {
                str2 = time.ToString("MMyy");
                if (this.bMmyy(str2))
                {
                    this.sql = this.sql + "#" + str.Replace("xxxd", this.user + "d" + str2).Replace("xxx", this.user + str2);
                }
                time = time.AddMonths(1);
            }
            return this.sql.Trim(new char[] { '#' }).Replace("#", " union all ");
        }

        public string get_stt_nhombc(DataTable table, string s_idnhom, string s_idloai)
        {
            string str = "0";
            string filterExpression = " idnhomvp=" + s_idnhom;
            foreach (DataRow row in table.Select(filterExpression, ""))
            {
                str = row["id"].ToString();
                if (("," + row["idloaivp"].ToString()).IndexOf("," + s_idloai.Trim() + ",") >= 0)
                {
                    return row["stt"].ToString();
                }
            }
            return str;
        }

        public string getid_process_Excel()
        {
            try
            {
                Process[] processesByName = Process.GetProcessesByName("EXCEL");
                if (processesByName.Length > 1)
                {
                    string str = "";
                    for (int i = 0; i <= (processesByName.Length - 1); i++)
                    {
                        str = str + processesByName[i].Id.ToString() + ",";
                    }
                    return str.Trim(new char[] { ',' });
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        public string getIndex(int i)
        {
            string[] strArray = new string[] { 
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", 
                "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", 
                "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", 
                "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", 
                "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", 
                "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", 
                "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", 
                "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", 
                "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", 
                "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", 
                "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", 
                "FU", "FV", "FW", "FX", "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", 
                "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", 
                "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", 
                "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "IA", "IB", "IC", "ID", "IE", "IF", 
                "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", 
                "IW", "IX", "IY", "IZ", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", 
                "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ", "KA", "KB", 
                "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", 
                "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ", "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", 
                "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", 
                "LY", "LZ", "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", 
                "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NB", "NC", "ND", 
                "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", 
                "NU", "NV", "NW", "NX", "NY", "NZ", "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OI", "OJ", 
                "OK", "OL", "OM", "ON", "OO", "OP", "OQ", "OR", "OS", "OT", "OU", "OV", "OW", "OX", "OY", "OZ", 
                "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", 
                "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", 
                "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", 
                "QW", "QX", "QY", "QZ", "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL", 
                "RM", "RN", "RO", "RP", "RQ", "RR", "RS", "RT", "RU", "RV", "RW", "RX", "RY", "RZ", "SA", "SB", 
                "SC", "SD", "SE", "SF", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SP", "SQ", "SR", 
                "SS", "ST", "SU", "SV", "SW", "SX", "SY", "SZ", "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", 
                "TI", "TJ", "TK", "TL", "TM", "TN", "TO", "TP", "TQ", "TR", "TS", "TT", "TU", "TV", "TW", "TX", 
                "TY", "TZ", "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UI", "UJ", "UK", "UL", "UM", "UN", 
                "UO", "UP", "UQ", "UR", "US", "UT", "UU", "UV", "UW", "UX", "UY", "UZ", "VA", "VB", "VC", "VD", 
                "VE", "VF", "VG", "VH", "VI", "VJ", "VK", "VL", "VM", "VN", "VO", "VP", "VQ", "VR", "VS", "VT", 
                "VU", "VV", "VW", "VX", "VY", "VZ", "WA", "WB", "WC", "WD", "WE", "WF", "WG", "WH", "WI", "WJ", 
                "WK", "WL", "WM", "WN", "WO", "WP", "WQ", "WR", "WS", "WT", "WU", "WV", "WW", "WX", "WY", "WZ", 
                "XA", "XB", "XC", "XD", "XE", "XF", "XG", "XH", "XI", "XJ", "XK", "XL", "XM", "XN", "XO", "XP", 
                "XQ", "XR", "XS", "XT", "XU", "XV", "XW", "XX", "XY", "XZ", "YA", "YB", "YC", "YD", "YE", "YF", 
                "YG", "YH", "YI", "YJ", "YK", "YL", "YM", "YN", "YO", "YP", "YQ", "YR", "YS", "YT", "YU", "YV", 
                "YW", "YX", "YY", "YZ", "ZA", "ZB", "ZC", "ZD", "ZE", "ZF", "ZG", "ZH", "ZI", "ZJ", "ZK", "ZL", 
                "ZM", "ZN", "ZO", "ZP", "ZQ", "ZR", "ZS", "ZT", "ZU", "ZV", "ZW", "ZX", "ZY", "ZZ"
             };
            return strArray[i];
        }

        public DataRow getrowbyid(DataTable dt, string exp)
        {
            try
            {
                return dt.Select(exp)[0];
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string Giamdoc(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=7 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public int iMavp_congkham(int d_nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=127 and nhom=" + d_nhom);
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0;
            }
            return int.Parse(this.ds.Tables[0].Rows[0][0].ToString().Trim());
        }

        public string Ketoan(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=15 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public string Ketoantruong(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=32 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public string Maincode(string sql)
        {
            XmlDocument document = new XmlDocument();
            document.Load(@"..\..\..\xml\maincode.xml");
            return document.GetElementsByTagName(sql).Item(0).InnerText;
        }

        public string mmyy(string ngay)
        {
            return (ngay.Substring(3, 2) + ngay.Substring(8, 2));
        }

        public int Ngay_toa_bhyt()
        {
            this.ds = this.get_data("select ten from d_thongso where id=99 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return int.Parse(this.ds.Tables[0].Rows[0]["ten"].ToString());
            }
            return 1;
        }

        public string Phutrach(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=8 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public void Run_fileExcel(string tenfile)
        {
            try
            {
                int id = 0;
                Process[] processes = Process.GetProcesses();
                if (processes.Length > 1)
                {
                    for (int i = 0; i <= (processes.Length - 1); i++)
                    {
                        if ((processes[i].ProcessName == "EXCEL") && (processes[i].Id > id))
                        {
                            id = processes[i].Id;
                        }
                    }
                    for (int j = 0; j <= (processes.Length - 1); j++)
                    {
                        if ((processes[j].ProcessName == "EXCEL") && (processes[j].Id == id))
                        {
                            processes[j].Kill();
                            break;
                        }
                    }
                }
                Process.Start(tenfile);
            }
            catch
            {
            }
        }

        public string sNoitrutinhnhuphongkham()
        {
            this.ds = this.get_data("select ten from thongso where id=493");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "0";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string sothe(int madoituong)
        {
            this.ds = this.get_data("select sothe,ngay,mabv,mien from doituong where madoituong=" + madoituong);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return (this.ds.Tables[0].Rows[0]["sothe"].ToString().Trim().PadLeft(2, '0') + this.ds.Tables[0].Rows[0]["ngay"].ToString().Trim() + this.ds.Tables[0].Rows[0]["mabv"].ToString().Trim() + this.ds.Tables[0].Rows[0]["mien"].ToString().Trim());
            }
            return "00000";
        }

        public bool sothe_doituong(int madoituong)
        {
            return (this.get_data("select * from doituong where sothe>0 and madoituong=" + madoituong).Tables[0].Rows.Count > 0);
        }

        public string sothemoi_18_95()
        {
            this.ds = this.get_data("select ten from thongso where id=435");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string sothemoi_80()
        {
            this.ds = this.get_data("select ten from thongso where id=50");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string sothemoi15_80()
        {
            this.ds = this.get_data("select ten from thongso where id=49");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string sothemoi15_95()
        {
            this.ds = this.get_data("select ten from thongso where id=-50");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string sTraituyentrenBHYTtra()
        {
            this.ds = this.get_data("select ten from thongso where id=490");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "0";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public DateTime StringToDate(string s)
        {
            string[] formats = new string[] { "dd/MM/yyyy" };
            return DateTime.ParseExact(s.Substring(0, 10), formats, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None);
        }

        public DateTime StringToDateTime(string s)
        {
            string[] strArray = new string[] { "dd/MM/yyyy" };
            string[] strArray2 = new string[] { "dd/MM/yyyy HH:mm" };
            return DateTime.ParseExact(s.ToString(), (s.Length == 10) ? strArray : strArray2, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None);
        }

        public DateTime StringToDateTime(string s, string f)
        {
            string[] formats = new string[] { f };
            return DateTime.ParseExact(s.ToString(), formats, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.None);
        }

        public void Taobangma_chuyen()
        {
            DataRow row;
            string str = "";
            string str2 = "";
            string str3 = "";
            str = "a+\x00b8+\x00b5+\x00b6+\x00b7+\x00b9+\x00a8+\x00be+\x00bb+\x00bc+\x00bd+\x00a8+\x00a9+\x00ca+\x00c7+\x00c8+\x00cb+b+c+d+\x00ae+e+\x00d0+\x00cc+\x00ce+\x00cf+\x00d1+\x00aa+\x00d5+\x00d2+\x00d3+\x00d4+\x00d6+i+\x00dd+\x00d7+\x00d8+\x00dc+\x00de+k+l+m+n+g+h+u+\x00f3+\x00ef+\x00f1+\x00f2+\x00f4+\x00ad+\x00f8+\x00f5+\x00f6+\x00f7+\x00f9+o+\x00ab+\x00e8+\x00e5+\x00e6+\x00e7+\x00e9+\x00ac+\x00ed+\x00ea+\x00eb+\x00ec+\x00ee+y+\x00fd+\x00fa+\x00fb+\x00fc+\x00fe+p+v+t+\x00e3+\x00df+\x00e1+\x00e2+\x00e4";
            str3 = "a+a\x00f9+a\x00f8+a\x00fb+a\x00f5+a\x00ef+a\x00ea+a\x00e9+a\x00e8+a\x00fa+a\x00fc+a\x00ea+a\x00e2+a\x00e1+a\x00e0+a\x00e5+a\x00e4+b+c+d+\x00f1+e+e\x00f9+e\x00f8+e\x00fb+e\x00f5+e\x00ef+e\x00e2+e\x00e1+e\x00e0+e\x00e5+e\x00e3+e\x00e4+i+\x00ed+\x00ec+\x00e6+\x00f3+\x00f2+k+l+m+n+g+h+u+u\x00f9+u\x00f8+u\x00fb+u\x00f5+u\x00ef+\x00f6+\x00f6\x00f9+\x00f6\x00f8+\x00f6\x00fb+\x00f6\x00f5+\x00f6\x00ef+o+o\x00e2+o\x00e1+o\x00e0+o\x00e5+o\x00e3+o\x00e4+\x00f4+\x00f4\x00f9+\x00f4\x00f8+\x00f4\x00fb+\x00f4\x00f5+\x00f4\x00ef+y+y\x00f9+y\x00f8+y\x00fb+y\x00f5+\x00ee+p+v+t+o\x00f9+o\x00f8+o\x00fb+o\x00f5+o\x00ef";
            str2 = "a+\x00e1+\x00e0+ả+\x00e3+ạ+ă+ắ+ằ+ẳ+ẵ+ă+\x00e2+ấ+ầ+ẩ+ậ+b+c+d+đ+e+\x00e9+\x00e8+ẻ+ẽ+ẹ+\x00ea+ế+ề+ể+ễ+ệ+i+\x00ed+\x00ec+ỉ+ĩ+ị+k+l+m+n+g+h+u+\x00fa+\x00f9+ủ+ũ+ụ+ư+ứ+ừ+ử+ữ+ự+o+\x00f4+ố+ồ+ổ+ỗ+ộ+ơ+ớ+ờ+ở+ỡ+ợ+y+\x00fd+ỳ+ỷ+ỹ+ỵ+p+v+t+\x00f3+\x00f2+ỏ+\x00f5+ọ";
            DataSet set = new DataSet();
            DataTable table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("ten", typeof(string));
            set.Tables.Add(table);
            DataSet set2 = new DataSet();
            DataSet set3 = new DataSet();
            DataSet set4 = new DataSet();
            set2 = set.Clone();
            string[] strArray = str.Split(new char[] { '+' });
            if (strArray.Length > 0)
            {
                for (int i = 0; i < strArray.Length; i++)
                {
                    row = set2.Tables[0].NewRow();
                    row["id"] = i + 1;
                    row["ten"] = strArray[i].ToString();
                    set2.Tables[0].Rows.Add(row);
                }
            }
            set2.AcceptChanges();
            set2.WriteXml("..//FontXml//abc.xml", XmlWriteMode.WriteSchema);
            set3 = set.Clone();
            strArray = str2.Split(new char[] { '+' });
            if (strArray.Length > 0)
            {
                for (int j = 0; j < strArray.Length; j++)
                {
                    row = set3.Tables[0].NewRow();
                    row["id"] = j + 1;
                    row["ten"] = strArray[j].ToString();
                    set3.Tables[0].Rows.Add(row);
                }
            }
            set3.AcceptChanges();
            set3.WriteXml("..//FontXml//unicode.xml", XmlWriteMode.WriteSchema);
            set4 = set.Clone();
            strArray = str3.Split(new char[] { '+' });
            if (strArray.Length > 0)
            {
                for (int k = 0; k < strArray.Length; k++)
                {
                    row = set4.Tables[0].NewRow();
                    row["id"] = k + 1;
                    row["ten"] = strArray[k].ToString();
                    set4.Tables[0].Rows.Add(row);
                }
            }
            set4.AcceptChanges();
            set4.WriteXml("..//FontXml//vni.xml", XmlWriteMode.WriteSchema);
        }

        public string tenfile_goc(string tenfile)
        {
            return Assembly.LoadFrom(tenfile).GetName().Name.ToString().ToLower();
        }

        public decimal themoi_chitra()
        {
            return 80M;
        }

        public decimal themoi_sotien()
        {
            this.ds = this.get_data("select ten from d_thongso where id=51 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0M;
            }
            return decimal.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public int themoi15_dodai()
        {
            this.ds = this.get_data("select ten from d_thongso where id=-50");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 15;
            }
            return int.Parse((this.ds.Tables[0].Rows[0][0].ToString() == "") ? "15" : this.ds.Tables[0].Rows[0][0].ToString());
        }

        public decimal themoi15_sotien()
        {
            this.ds = this.get_data("select ten from d_thongso where id=50 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0M;
            }
            return decimal.Parse(this.ds.Tables[0].Rows[0][0].ToString());
        }

        public string themoi15_vitri()
        {
            string str = this.vitrithe_moi15();
            return ((int.Parse(str.Substring(0, str.IndexOf(","))) + 1) + "," + int.Parse(str.Substring(str.IndexOf(",") + 1)));
        }

        public string thetrongtinh()
        {
            this.ds = this.get_data("select ten from d_thongso where id=83 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return "50";
            }
            return this.ds.Tables[0].Rows[0][0].ToString();
        }

        public string thetrongtinh_vitri()
        {
            string str = "2,2";
            this.ds = this.get_data("select ten from d_thongso where id=84 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                str = this.ds.Tables[0].Rows[0][0].ToString().Trim();
                int index = str.IndexOf(",");
                if (index != -1)
                {
                    str = ((int.Parse(str.Substring(0, index)) - 1)).ToString() + "," + int.Parse(str.Substring(index + 1)).ToString();
                }
            }
            return str;
        }

        public string Thongke(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=9 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public string Thongsobhyt(string sql)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\thongsobhyt.xml");
                return document.GetElementsByTagName(sql.ToUpper()).Item(0).InnerText;
            }
            catch
            {
                this.ds = new DataSet();
                this.ds.ReadXml(@"..\..\..\xml\thongsobhyt.xml");
                DataColumn column = new DataColumn();
                column.ColumnName = sql.ToUpper();
                column.DataType = System.Type.GetType("System.String");
                this.ds.Tables[0].Columns.Add(column);
                this.ds.Tables[0].Rows[0][sql.ToUpper()] = "0";
                this.ds.WriteXml(@"..\..\..\xml\thongsobhyt.xml");
                return "0";
            }
        }

        public string Thukho(int nhom)
        {
            this.ds = this.get_data("select ten from d_thongso where id=16 and nhom=" + nhom);
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                return this.ds.Tables[0].Rows[0]["ten"].ToString().Trim();
            }
            return "";
        }

        public decimal tile_traituyen()
        {
            this.ds = this.get_data("select ten from thongso where id=483 ");
            if (this.ds.Tables[0].Rows.Count == 0)
            {
                return 0M;
            }
            return decimal.Parse((this.ds.Tables[0].Rows[0][0].ToString() == "") ? "0" : this.ds.Tables[0].Rows[0][0].ToString());
        }

        public string Unicode_TCVN3_ABC(string sUnicode)
        {
            DataSet set = new DataSet();
            DataSet set2 = new DataSet();
            if (!File.Exists("..//FontXml//unicode.xml") || !File.Exists("..//FontXml//abc.xml"))
            {
                this.Taobangma_chuyen();
            }
            set.ReadXml("..//FontXml//unicode.xml", XmlReadMode.ReadSchema);
            set2.ReadXml("..//FontXml//abc.xml", XmlReadMode.ReadSchema);
            string str = sUnicode.ToLower().Trim();
            string exp = "";
            string str3 = "";
            for (int i = 0; i < str.Length; i++)
            {
                char ch = str[i];
                if (ch.ToString() != " ")
                {
                    exp = " ten ='" + str[i].ToString() + "'";
                    DataRow row = this.getrowbyid(set.Tables[0], exp);
                    if (row != null)
                    {
                        int num = int.Parse(row["id"].ToString());
                        exp = "id=" + num;
                        DataRow row2 = this.getrowbyid(set2.Tables[0], exp);
                        if (row2 != null)
                        {
                            str3 = str3 + row2["ten"].ToString();
                        }
                        else
                        {
                            str3 = str3 + str[i].ToString();
                        }
                    }
                    else
                    {
                        str3 = str3 + str[i].ToString();
                    }
                }
                else
                {
                    str3 = str3 + " ";
                }
            }
            return str3;
        }

        public void upd_bcbhyt(int l_id, int i_stt, string s_ten, int i_idnhomvp, string i_idloaivp)
        {
            string commandText = "";
            string str2 = "";
            int num = 0;
            commandText = "update dmbaocao_bhyt set stt=:i_stt,ten=:s_ten,idnhomvp=:i_idnhomvp,idloaivp=:i_idloaivp where id=:l_id";
            str2 = "insert into dmbaocao_bhyt(id,stt,ten,idnhomvp,idloaivp) values(:l_id,:i_stt,:s_ten,:i_idnhomvp,:i_idloaivp)";
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add("l_id", OracleType.Number, 3).Value = l_id;
            command.Parameters.Add("i_stt", OracleType.Number, 3).Value = i_stt;
            command.Parameters.Add("s_ten", OracleType.NVarChar, 100).Value = s_ten;
            command.Parameters.Add("i_idnhomvp", OracleType.Number, 3).Value = i_idnhomvp;
            command.Parameters.Add("i_idloaivp", OracleType.VarChar, 80).Value = i_idloaivp;
            if (command.ExecuteNonQuery() == 0)
            {
                command.Dispose();
                command = new OracleCommand(str2, connection);
                command.Parameters.Add("l_id", OracleType.Number, 3).Value = l_id;
                command.Parameters.Add("i_stt", OracleType.Number, 3).Value = i_stt;
                command.Parameters.Add("s_ten", OracleType.NVarChar, 100).Value = s_ten;
                command.Parameters.Add("i_idnhomvp", OracleType.Number, 3).Value = i_idnhomvp;
                command.Parameters.Add("i_idloaivp", OracleType.VarChar, 80).Value = i_idloaivp;
                try
                {
                    num = command.ExecuteNonQuery();
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString());
                }
            }
            command.Dispose();
            connection.Dispose();
        }

        public void upd_bcbhyt_dmnhombhytct(int id, int idvp, string ma, string ten, int idnhommedi)
        {
            string commandText = "";
            commandText = string.Concat(new object[] { "insert into bcbhyt_dmnhombhytct(id,idvp,ma,ten,idnhombhytmedisoft)  values(", id, ",", idvp, ",:sma,:sten,", idnhommedi, ")" });
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            try
            {
                command = new OracleCommand(commandText, connection);
                command.Parameters.Add("sma", OracleType.VarChar, 50).Value = ma;
                command.Parameters.Add("sten", OracleType.NVarChar, 200).Value = ten;
                command.ExecuteNonQuery();
                command.Dispose();
            }
            catch (Exception exception)
            {
                this.upd_error(exception.ToString(), this.sComputer, "bcbhyt_dmnhombhytct");
            }
            connection.Dispose();
        }

        public void upd_bhytcls(string mmyy, long id, int stt, int mavp, int soluong, decimal dongia, int idchidinh, int sttra, int sobienlai)
        {
            string commandText = "";
            string str2 = "";
            int num = 0;
            commandText = "update " + this.user + "d" + mmyy + ".bhytcls set mavp=:mavp,dongia=:dongia where id=:id and stt=:stt";
            str2 = "insert into " + this.user + "d" + mmyy + ".bhytcls(id,stt,mavp,soluong,dongia,idchidinh,sttra,sobienlai) values(:id,:stt,:mavp,:soluong,:dongia,:idchidinh,:sttra,:sobienlai)";
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add("id", OracleType.Number, 3).Value = id;
            command.Parameters.Add("stt", OracleType.Number, 3).Value = stt;
            command.Parameters.Add("mavp", OracleType.Number, 3).Value = mavp;
            command.Parameters.Add("dongia", OracleType.Number, 3).Value = dongia;
            if (command.ExecuteNonQuery() == 0)
            {
                command.Dispose();
                command = new OracleCommand(str2, connection);
                command.Parameters.Add("id", OracleType.Number, 3).Value = id;
                command.Parameters.Add("stt", OracleType.Number, 3).Value = stt;
                command.Parameters.Add("mavp", OracleType.Number, 3).Value = mavp;
                command.Parameters.Add("soluong", OracleType.Number, 3).Value = soluong;
                command.Parameters.Add("dongia", OracleType.Number, 3).Value = dongia;
                command.Parameters.Add("idchidinh", OracleType.Number, 3).Value = idchidinh;
                command.Parameters.Add("sttra", OracleType.Number, 3).Value = sttra;
                command.Parameters.Add("sobienlai", OracleType.Number, 3).Value = sobienlai;
                try
                {
                    num = command.ExecuteNonQuery();
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString());
                }
            }
            command.Dispose();
            connection.Dispose();
        }

        public void upd_boiduong_pttt(string b_pttt, string b_loai, string b_tenloai, decimal b_ptv, decimal b_phu1, decimal b_phu2, decimal b_bsgayme, decimal b_ktvgayme, decimal b_hoisuc, decimal b_dungcu)
        {
            string commandText = "";
            string str2 = "";
            int num = 0;
            commandText = "update boiduong_pttt set ptv=:b_ptv,tenloai=:b_tenloai";
            commandText = (commandText + ",phu1=:b_phu1,phu2=:b_phu2,bsgayme=:b_bsgayme,ktvgayme=:b_ktvgayme") + ",hoisuc=:b_hoisuc,dungcu=:b_dungcu  " + " where pttt=:b_pttt and loai=:b_loai";
            str2 = string.Concat(new object[] { 
                "insert into boiduong_pttt(pttt,loai,tenloai,ptv,phu1,phu2,bsgayme,ktvgayme,hoisuc,dungcu)  values('", b_pttt, "','", b_loai, "',:b_tenloai,", b_ptv, ",", b_phu1, ",", b_phu2, ",", b_bsgayme, ",", b_ktvgayme, ",", b_hoisuc, 
                ",", b_dungcu, ")"
             });
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add("b_tenloai", OracleType.NVarChar, 60).Value = b_tenloai;
            command.Parameters.Add("b_ptv", OracleType.Number).Value = b_ptv;
            command.Parameters.Add("b_phu1", OracleType.Number).Value = b_phu1;
            command.Parameters.Add("b_phu2", OracleType.Number).Value = b_phu2;
            command.Parameters.Add("b_bsgayme", OracleType.Number).Value = b_bsgayme;
            command.Parameters.Add("b_ktvgayme", OracleType.Number).Value = b_ktvgayme;
            command.Parameters.Add("b_hoisuc", OracleType.Number).Value = b_hoisuc;
            command.Parameters.Add("b_dungcu", OracleType.Number).Value = b_dungcu;
            command.Parameters.Add("b_pttt", OracleType.VarChar, 1).Value = b_pttt;
            command.Parameters.Add("b_loai", OracleType.VarChar, 10).Value = b_loai;
            if (command.ExecuteNonQuery() == 0)
            {
                command.Dispose();
                command = new OracleCommand(str2, connection);
                command.Parameters.Add("b_tenloai", OracleType.NVarChar, 60).Value = b_tenloai;
                try
                {
                    num = command.ExecuteNonQuery();
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString());
                }
            }
            command.Dispose();
            connection.Dispose();
        }

        public void upd_d_thongso(int id, string ten, string tendef, string tencur)
        {
            int nhomkho = this.nhomkho;
            this.sql = "update d_thongso set ten=:ten,tendef=:tendef,tencur=:tencur ";
            this.sql = this.sql + " where id=:id and nhom=:nhom";
            this.con = new OracleConnection(this.sConn);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(this.sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("ten", OracleType.NVarChar, 0xfe).Value = ten;
                this.cmd.Parameters.Add("tendef", OracleType.NVarChar, 0xfe).Value = tendef;
                this.cmd.Parameters.Add("tencur", OracleType.NVarChar, 0xfe).Value = tencur;
                this.cmd.Parameters.Add("id", OracleType.Number).Value = id;
                this.cmd.Parameters.Add("nhom", OracleType.Number).Value = nhomkho;
                int num2 = this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                if (num2 == 0)
                {
                    this.sql = "insert into d_thongso(id,ten,nhom,tendef,tencur) values (:id,:ten,:nhom,:tendef,:tencur)";
                    this.cmd = new OracleCommand(this.sql, this.con);
                    this.cmd.CommandType = CommandType.Text;
                    this.cmd.Parameters.Add("id", OracleType.Number).Value = id;
                    this.cmd.Parameters.Add("ten", OracleType.NVarChar, 0xfe).Value = ten;
                    this.cmd.Parameters.Add("nhom", OracleType.Number).Value = nhomkho;
                    this.cmd.Parameters.Add("tendef", OracleType.NVarChar, 0xfe).Value = tendef;
                    this.cmd.Parameters.Add("tencur", OracleType.NVarChar, 0xfe).Value = tencur;
                    num2 = this.cmd.ExecuteNonQuery();
                    this.cmd.Dispose();
                }
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim(), this.sComputer, "d_Thongso");
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
        }

        public void upd_dmnhomdv_bhyt(int i_manhom, int i_stt, string s_nhombv, string s_tennhom, string s_mabv, string s_tenbv, bool bXoa)
        {
            string commandText = "";
            string str2 = "";
            string str3 = "";
            int num = 0;
            str3 = "delete from dmnhomdv_bhyt where manhom=:i_manhom and stt=:i_stt and mabv=:s_mabv";
            commandText = "update dmnhomdv_bhyt set nhombv=:s_nhombv,tennhom=:s_tennhom,mabv=:s_mabv,tenbv=:s_tenbv where manhom=:i_manhom and stt=:i_stt and mabv=:s_mabv";
            str2 = "insert into dmnhomdv_bhyt(manhom,nhombv,tennhom,mabv,tenbv,stt) values(:i_manhom,:s_nhombv,:s_tennhom,:s_mabv,:s_tenbv,:i_stt)";
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            try
            {
                if (bXoa)
                {
                    command = new OracleCommand(str3, connection);
                    command.CommandType = CommandType.Text;
                    command.Parameters.Add("i_manhom", OracleType.Number, 3).Value = i_manhom;
                    command.Parameters.Add("s_mabv", OracleType.NVarChar, 8).Value = s_mabv;
                    command.Parameters.Add("i_stt", OracleType.Number, 3).Value = i_stt;
                    command.ExecuteNonQuery();
                    command.Dispose();
                }
                else
                {
                    command = new OracleCommand(commandText, connection);
                    command.CommandType = CommandType.Text;
                    command.Parameters.Add("i_manhom", OracleType.Number, 3).Value = i_manhom;
                    command.Parameters.Add("s_nhombv", OracleType.NVarChar, 8).Value = s_nhombv;
                    command.Parameters.Add("s_tennhom", OracleType.NVarChar, 200).Value = s_tennhom;
                    command.Parameters.Add("s_mabv", OracleType.NVarChar, 8).Value = s_mabv;
                    command.Parameters.Add("s_tenbv", OracleType.NVarChar, 200).Value = s_tenbv;
                    command.Parameters.Add("i_stt", OracleType.Number, 3).Value = i_stt;
                    num = command.ExecuteNonQuery();
                    command.Dispose();
                    if (num == 0)
                    {
                        command.Dispose();
                        command = new OracleCommand(str2, connection);
                        command.Parameters.Add("i_manhom", OracleType.Number, 3).Value = i_manhom;
                        command.Parameters.Add("s_nhombv", OracleType.NVarChar, 8).Value = s_nhombv;
                        command.Parameters.Add("s_tennhom", OracleType.NVarChar, 200).Value = s_tennhom;
                        command.Parameters.Add("s_mabv", OracleType.NVarChar, 8).Value = s_mabv;
                        command.Parameters.Add("s_tenbv", OracleType.NVarChar, 200).Value = s_tenbv;
                        command.Parameters.Add("i_stt", OracleType.Number, 3).Value = i_stt;
                        command.ExecuteNonQuery();
                        command.Dispose();
                    }
                }
            }
            catch (Exception exception)
            {
                this.upd_error(exception.ToString(), this.sComputer, "dmnhomdv_bhyt");
            }
            connection.Dispose();
        }

        public bool upd_doituong_bhyt(string s_ma_old, string s_ma13, string s_ma_dt, string s_ten_dt, int n_ty_le, string s_ghichu, double n_tien1, double n_tien2, int n_loai, string s_tenth, int n_tonghop, int n_admin)
        {
            int num = 0;
            bool flag = true;
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            this.sql = "update doituong_bhyt set ma_old=:s_ma_old,ma13=:s_ma13,ten_dt=:s_ten_dt,ty_le=:n_ty_le,ghichu=:s_ghichu,tien1=:n_tien1, tien2=:n_tien2,loai=:n_loai,tenth=:s_tenth,tonghop=:n_tonghop,readonly=:n_admin where ma_dt=:s_ma_dt";
            OracleCommand command = new OracleCommand(this.sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add("s_ma_old", OracleType.VarChar, 2).Value = s_ma_old;
            command.Parameters.Add("s_ma13", OracleType.VarChar, 100).Value = s_ma13;
            command.Parameters.Add("s_ma_dt", OracleType.VarChar).Value = s_ma_dt;
            command.Parameters.Add("s_ten_dt", OracleType.NVarChar, 100).Value = s_ten_dt;
            command.Parameters.Add("n_ty_le", OracleType.Number).Value = n_ty_le;
            command.Parameters.Add("s_ghichu", OracleType.NVarChar, 0x7d0).Value = s_ghichu;
            command.Parameters.Add("n_tien1", OracleType.Number).Value = n_tien1;
            command.Parameters.Add("n_tien2", OracleType.Number).Value = n_tien2;
            command.Parameters.Add("n_loai", OracleType.Number).Value = n_loai;
            command.Parameters.Add("s_tenth", OracleType.NVarChar, 100).Value = s_tenth;
            command.Parameters.Add("n_tonghop", OracleType.Number).Value = n_tonghop;
            command.Parameters.Add("n_admin", OracleType.Number).Value = n_admin;
            if (command.ExecuteNonQuery() == 0)
            {
                this.sql = "insert into doituong_bhyt (ma_old,ma13,ten_dt,ty_le,ghichu,tien1, tien2,loai,tenth,tonghop,ma_dt,readonly) values (:s_ma_old,:s_ma13,:s_ten_dt,:n_ty_le,:s_ghichu,:n_tien1,:n_tien2,:n_loai,:s_tenth,:n_tonghop,:s_ma_dt,:n_admin)";
                command.Dispose();
                command = new OracleCommand(this.sql, connection);
                command.Parameters.Add("s_ma_old", OracleType.VarChar, 2).Value = s_ma_old;
                command.Parameters.Add("s_ma13", OracleType.VarChar, 100).Value = s_ma13;
                command.Parameters.Add("s_ma_dt", OracleType.VarChar).Value = s_ma_dt;
                command.Parameters.Add("s_ten_dt", OracleType.NVarChar, 100).Value = s_ten_dt;
                command.Parameters.Add("n_ty_le", OracleType.Number).Value = n_ty_le;
                command.Parameters.Add("s_ghichu", OracleType.NVarChar, 0x7d0).Value = s_ghichu;
                command.Parameters.Add("n_tien1", OracleType.Number).Value = n_tien1;
                command.Parameters.Add("n_tien2", OracleType.Number).Value = n_tien2;
                command.Parameters.Add("n_loai", OracleType.Number).Value = n_loai;
                command.Parameters.Add("s_tenth", OracleType.NVarChar, 100).Value = s_tenth;
                command.Parameters.Add("n_tonghop", OracleType.Number).Value = n_tonghop;
                command.Parameters.Add("n_admin", OracleType.Number).Value = n_admin;
                try
                {
                    num = command.ExecuteNonQuery();
                    flag = true;
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString() + "\n -" + s_ma_dt);
                    flag = false;
                }
            }
            else
            {
                flag = true;
            }
            command.Dispose();
            connection.Dispose();
            return flag;
        }

        public bool upd_doituongth_bhyt(int s_id, int s_stt, string s_math, string s_tenth, int s_maloai, string s_tenloai, string s_ghichu)
        {
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            this.sql = "update doituongth_bhyt set stt=:s_stt,math=:s_math,tenth=:s_tenth,maloai=:s_maloai,tenloai=:s_tenloai,ghichu=:s_ghichu where id=:s_id";
            OracleCommand command = new OracleCommand(this.sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add("s_id", OracleType.Number, 4).Value = s_id;
            command.Parameters.Add("s_stt", OracleType.Number, 4).Value = s_stt;
            command.Parameters.Add("s_math", OracleType.NVarChar, 100).Value = s_math;
            command.Parameters.Add("s_tenth", OracleType.NVarChar, 100).Value = s_tenth;
            command.Parameters.Add("s_maloai", OracleType.Number, 1).Value = s_maloai;
            command.Parameters.Add("s_tenloai", OracleType.NVarChar, 100).Value = s_tenloai;
            command.Parameters.Add("s_ghichu", OracleType.NVarChar, 100).Value = s_ghichu;
            if (command.ExecuteNonQuery() == 0)
            {
                this.sql = "insert into doituongth_bhyt(id,stt,math,tenth,maloai,tenloai,ghichu) values(:s_id,:s_stt,:s_math,:s_tenth,:s_maloai,:s_tenloai,:s_ghichu)";
                command.Dispose();
                command = new OracleCommand(this.sql, connection);
                command.Parameters.Add("s_id", OracleType.Number, 4).Value = s_id;
                command.Parameters.Add("s_stt", OracleType.Number, 4).Value = s_stt;
                command.Parameters.Add("s_math", OracleType.NVarChar, 100).Value = s_math;
                command.Parameters.Add("s_tenth", OracleType.NVarChar, 100).Value = s_tenth;
                command.Parameters.Add("s_maloai", OracleType.Number, 1).Value = s_maloai;
                command.Parameters.Add("s_tenloai", OracleType.NVarChar, 100).Value = s_tenloai;
                command.Parameters.Add("s_ghichu", OracleType.NVarChar, 100).Value = s_ghichu;
                try
                {
                    int num = command.ExecuteNonQuery();
                    return true;
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString() + "\n -" + s_id);
                    return false;
                }
            }
            return true;
        }

        public void upd_error(string m_message, string m_computer, string m_table)
        {
            this.cmd.Dispose();
            this.con.Close();
            this.sql = "insert into error(message,computer,tables,ngayud) values (:m_message,:m_computer,:m_table,sysdate)";
            this.con = new OracleConnection(this.sConn);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(this.sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("m_message", OracleType.VarChar, 0xfe).Value = m_message;
                this.cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                this.cmd.Parameters.Add("m_table", OracleType.VarChar, 20).Value = m_table;
                this.cmd.ExecuteNonQuery();
            }
            catch
            {
            }
            finally
            {
                this.cmd.Dispose();
                this.con.Close();
            }
        }

        public void upd_quyen(int userid, string quyen)
        {
            string commandText = "";
            string str2 = "";
            int num = 0;
            commandText = "update v_loginbh set right_='" + quyen + "' where id=" + userid.ToString();
            str2 = "insert into v_loginbh(id, right_) values(" + userid.ToString() + ",'" + quyen + "')";
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            if (command.ExecuteNonQuery() == 0)
            {
                command.Dispose();
                command = new OracleCommand(str2, connection);
                try
                {
                    num = command.ExecuteNonQuery();
                }
                catch (OracleException exception)
                {
                    MessageBox.Show(exception.ToString());
                }
            }
            command.Dispose();
            connection.Dispose();
        }

        public string upd_table_fieldTen(string m_table, string m_fieldName, string m_giatri, string m_fieldKey, string m_Keyvalue)
        {
            this.sql = "update " + m_table + " set " + m_fieldName + "=:" + m_fieldName;
            string sql = this.sql;
            this.sql = sql + " where " + m_fieldKey + "=" + m_Keyvalue;
            this.con = new OracleConnection(this.sConn);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(this.sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add(m_fieldName, OracleType.NVarChar, 0x3e8).Value = m_giatri;
                int num = this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message, this.sComputer, m_table);
                return m_fieldName;
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
            return "";
        }

        public string upd_tenvien(string m_table, string m_mabv, string m_tenbv, string m_diachi, string m_sodt, string m_matuyen, string m_maloai, string m_mahang, string m_mavung, string m_matinh)
        {
            this.sql = "update " + m_table + " set tenbv=:m_tenbv ";
            this.sql = this.sql + " where mabv=:m_mabv";
            this.con = new OracleConnection(this.sConn);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(this.sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("m_tenbv", OracleType.NVarChar, 100).Value = m_tenbv;
                this.cmd.Parameters.Add("m_mabv", OracleType.VarChar, 8).Value = m_mabv;
                int num = this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                if (num == 0)
                {
                    this.sql = "insert into " + m_table + " (mabv,tenbv,diachi,sodt,matuyen,maloai,mahang,mavung,matinh) values (:m_mabv,:m_tenbv,:m_diachi,:m_sodt,:m_matuyen,:m_maloai,:m_mahang,:m_mavung,:m_matinh)";
                    this.cmd = new OracleCommand(this.sql, this.con);
                    this.cmd.CommandType = CommandType.Text;
                    this.cmd.Parameters.Add("m_mabv", OracleType.VarChar, 8).Value = m_mabv;
                    this.cmd.Parameters.Add("m_tenbv", OracleType.NVarChar, 100).Value = m_tenbv;
                    this.cmd.Parameters.Add("m_diachi", OracleType.NVarChar, 100).Value = m_diachi;
                    this.cmd.Parameters.Add("m_sodt", OracleType.VarChar, 20).Value = m_sodt;
                    this.cmd.Parameters.Add("m_matuyen", OracleType.VarChar, 1).Value = m_matuyen;
                    this.cmd.Parameters.Add("m_maloai", OracleType.VarChar, 5).Value = m_maloai;
                    this.cmd.Parameters.Add("m_mahang", OracleType.VarChar, 1).Value = m_mahang;
                    this.cmd.Parameters.Add("m_mavung", OracleType.VarChar, 1).Value = m_mavung;
                    this.cmd.Parameters.Add("m_matinh", OracleType.VarChar, 3).Value = m_matinh;
                    num = this.cmd.ExecuteNonQuery();
                    this.cmd.Dispose();
                }
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message, this.sComputer, m_table);
                return m_mabv;
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
            return "";
        }

        public void upd_thongso(int id, string ten, string tendef, string tencur)
        {
            this.sql = "update thongso set ten=:ten,tendef=:tendef,tencur=:tencur ";
            this.sql = this.sql + " where id=:id";
            this.con = new OracleConnection(this.sConn);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(this.sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("ten", OracleType.NVarChar, 0xfe).Value = ten;
                this.cmd.Parameters.Add("tendef", OracleType.NVarChar, 0xfe).Value = tendef;
                this.cmd.Parameters.Add("tencur", OracleType.NVarChar, 0xfe).Value = tencur;
                this.cmd.Parameters.Add("id", OracleType.Number).Value = id;
                int num = this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                if (num == 0)
                {
                    this.sql = "insert into thongso(id,ten,tendef,tencur) values (:id,:ten,:tendef,:tencur)";
                    this.cmd = new OracleCommand(this.sql, this.con);
                    this.cmd.CommandType = CommandType.Text;
                    this.cmd.Parameters.Add("id", OracleType.Number).Value = id;
                    this.cmd.Parameters.Add("ten", OracleType.NVarChar, 0xfe).Value = ten;
                    this.cmd.Parameters.Add("tendef", OracleType.NVarChar, 0xfe).Value = tendef;
                    this.cmd.Parameters.Add("tencur", OracleType.NVarChar, 0xfe).Value = tencur;
                    num = this.cmd.ExecuteNonQuery();
                    this.cmd.Dispose();
                }
            }
            catch (OracleException exception)
            {
                this.upd_error(exception.Message.ToString().Trim(), this.sComputer, "Thongso");
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
        }

        public void upd_xuatsd_glivec(string mabn, int mabd, decimal soluong, decimal dongia, string makp_bhyt, string manhom_bhyt, string ngay)
        {
            string commandText = "";
            string str2 = "";
            int num = 0;
            commandText = string.Concat(new object[] { 
                "update ", this.user, ".d_xuatsd_glivec set soluong=", soluong, ",dongia=", dongia, ",manhom_bhyt='", manhom_bhyt, "',ngayud=sysdate where mabn='", mabn, "' and mabd=", mabd.ToString(), " and makp_bhyt='", makp_bhyt, "' and to_char(ngay,'dd/mm/yyyy')='", ngay, 
                "'"
             });
            str2 = string.Concat(new object[] { 
                "insert into ", this.user, ".d_xuatsd_glivec(mabn,mabd,manhom_bhyt,soluong,dongia,makp_bhyt,ngay) values('", mabn, "',", mabd, ",'", manhom_bhyt, "',", soluong, ",", dongia, ",'", makp_bhyt, "',to_date('", ngay, 
                "','dd/mm/yyyy'))"
             });
            OracleConnection connection = new OracleConnection(this.ConStr);
            connection.Open();
            OracleCommand command = new OracleCommand(commandText, connection);
            command.CommandType = CommandType.Text;
            if (command.ExecuteNonQuery() == 0)
            {
                command.Dispose();
                command = new OracleCommand(str2, connection);
                try
                {
                    num = command.ExecuteNonQuery();
                }
                catch (OracleException exception)
                {
                    this.upd_error(exception.ToString(), Environment.MachineName, "d_glivec");
                }
            }
            command.Dispose();
            connection.Dispose();
        }

        public void update_table_column()
        {
            DataTable table;
            string sql = "";
            try
            {
                sql = "select stt_40 from dmbd where 1=2";
                table = this.get_data(sql).Tables[0];
            }
            catch
            {
                sql = "alter table dmbd add stt_40 nvarchar2(20)";
                this.execute_data(sql);
            }
            try
            {
                sql = "select mathuoc from d_theodoi where 1=2";
                table = this.get_data(sql).Tables[0];
            }
            catch
            {
                sql = "alter table dmbd add stt_40 nvarchar2(20)";
                this.execute_data(sql);
            }
        }

        public string vitrithe_13()
        {
            string str = "0,2";
            this.ds = this.get_data("select ten from d_thongso where id=52 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                str = this.ds.Tables[0].Rows[0][0].ToString().Trim();
                int index = str.IndexOf(",");
                if (index != -1)
                {
                    str = ((int.Parse(str.Substring(0, index)) - 1)).ToString() + "," + int.Parse(str.Substring(index + 1)).ToString();
                }
            }
            return str;
        }

        public string vitrithe_moi()
        {
            string str = "5,2";
            this.ds = this.get_data("select ten from d_thongso where id=53 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                str = this.ds.Tables[0].Rows[0][0].ToString().Trim();
                int index = str.IndexOf(",");
                if (index != -1)
                {
                    str = ((int.Parse(str.Substring(0, index)) - 1)).ToString() + "," + int.Parse(str.Substring(index + 1)).ToString();
                }
            }
            return str;
        }

        public string vitrithe_moi15()
        {
            string str = "3,1";
            this.ds = this.get_data("select ten from d_thongso where id=52 and nhom=" + this.nhomkho + " ");
            if (this.ds.Tables[0].Rows.Count > 0)
            {
                str = this.ds.Tables[0].Rows[0][0].ToString().Trim();
                int index = str.IndexOf(",");
                if (index != -1)
                {
                    str = ((int.Parse(str.Substring(0, index)) - 1)).ToString() + "," + int.Parse(str.Substring(index + 1)).ToString();
                }
            }
            return str;
        }

        public void writeXml(string tenfile, string cot, string s)
        {
            this.ds = new DataSet();
            try
            {
                if (File.Exists(@"..\..\..\xml\" + tenfile + ".xml"))
                {
                    this.ds.ReadXml(@"..\..\..\xml\" + tenfile + ".xml");
                    this.ds.Tables[0].Rows[0][cot.ToUpper()] = s.ToUpper();
                }
                else
                {
                    this.sql = "select '" + s + "' as " + cot + " from dual ";
                    this.ds = this.get_data(this.sql);
                }
            }
            catch
            {
                DataColumn column = new DataColumn();
                column.ColumnName = cot;
                column.DataType = System.Type.GetType("System.String");
                this.ds.Tables[0].Columns.Add(column);
                this.ds.Tables[0].Rows[0][cot.ToUpper()] = s.ToUpper();
            }
            this.ds.WriteXml(@"..\..\..\xml\" + tenfile + ".xml");
        }

        public int baocaobhyt
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("baocaobhyt").Item(0).InnerText.ToString());
            }
        }

        public int baocaote
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("baocaote").Item(0).InnerText.ToString());
            }
        }

        public bool bBHYTngtr_noitru
        {
            get
            {
                try
                {
                    return (int.Parse(this.get_data("select ten from thongso where id=183").Tables[0].Rows[0][0].ToString()) == 1);
                }
                catch
                {
                    return false;
                }
            }
        }

        public int bGiabhyt
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("giabhyt").Item(0).InnerText.ToString());
            }
        }

        public string BHXH_TINH
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("dv2").Item(0).InnerText;
            }
        }

        public string BHXH_VN
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("dv1").Item(0).InnerText;
            }
        }

        public int congkhamchuyenkham
        {
            get
            {
                try
                {
                    return int.Parse(this.get_data("select ten from d_thongso where id=187 and nhom=" + this.nhomkho + "").Tables[0].Rows[0][0].ToString());
                }
                catch
                {
                    return 0;
                }
            }
        }

        public string ConStr
        {
            get
            {
                return this.sConn;
            }
        }

        public string Diachi
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\maincode.xml");
                return document.GetElementsByTagName("Diachi").Item(0).InnerText;
            }
        }

        public string DV_KCB
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("dv3").Item(0).InnerText;
            }
        }

        public int iCongkham
        {
            get
            {
                try
                {
                    XmlDocument document = new XmlDocument();
                    document.Load(@"..\..\..\xml\bhxh.xml");
                    return int.Parse(document.GetElementsByTagName("congkham").Item(0).InnerText.ToString());
                }
                catch
                {
                    return 0;
                }
            }
        }

        public int iKhambenh
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("tienkham").Item(0).InnerText.ToString());
            }
        }

        public string iManhomKTCao
        {
            get
            {
                try
                {
                    XmlDocument document = new XmlDocument();
                    document.Load(@"..\..\..\xml\bhxh.xml");
                    return document.GetElementsByTagName("manhomktcao").Item(0).InnerText.ToString();
                }
                catch
                {
                    return "";
                }
            }
        }

        public int iMuctinhbhytmoi
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("mucbhytmoi").Item(0).InnerText.ToString());
            }
        }

        public int iNgaykiemke
        {
            get
            {
                int num = 0;
                this.ds = this.get_data("select ten from d_thongso where id=105");
                foreach (DataRow row in this.ds.Tables[0].Rows)
                {
                    if (num < int.Parse(row["ten"].ToString()))
                    {
                        num = int.Parse(row["ten"].ToString());
                    }
                }
                if (num == 0)
                {
                    num = 7;
                }
                return num;
            }
        }

        public int iNhombaocao
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("nhombaocao").Item(0).InnerText.ToString());
            }
        }

        public int iNhomvpthuoc
        {
            get
            {
                try
                {
                    return int.Parse(this.get_data("select ten from thongso where id=371").Tables[0].Rows[0][0].ToString());
                }
                catch
                {
                    return 0;
                }
            }
        }

        public int iPttt_mien_bhyt
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("pttt").Item(0).InnerText.ToString());
            }
        }

        public int iTienkham
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("tienkham").Item(0).InnerText.ToString());
            }
        }

        public int iTonghop
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("tonghop").Item(0).InnerText.ToString());
            }
        }

        public int iTreem6tuoi
        {
            get
            {
                try
                {
                    return int.Parse(this.get_data("select ten from thongso where id=147").Tables[0].Rows[0][0].ToString());
                }
                catch
                {
                    return 0;
                }
            }
        }

        public int laysolieututrucpk
        {
            get
            {
                try
                {
                    return int.Parse(this.get_data("select ten from d_thongso where id=142 and nhom=" + this.nhomkho + "").Tables[0].Rows[0][0].ToString());
                }
                catch
                {
                    return 0;
                }
            }
        }

        public string MABHXH
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("mabhxh").Item(0).InnerText;
            }
        }

        public string Mabv
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                XmlNodeList elementsByTagName = document.GetElementsByTagName("mabv");
                return ((elementsByTagName.Item(0).InnerText == "") ? "701.1.01" : elementsByTagName.Item(0).InnerText);
            }
        }

        public string MAYTCQ
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("maytcq").Item(0).InnerText;
            }
        }

        public int nhom_duoc
        {
            get
            {
                try
                {
                    return int.Parse(this.get_data("select ten from thongso where id=166").Tables[0].Rows[0][0].ToString());
                }
                catch
                {
                    return 1;
                }
            }
        }

        public int nhomkho
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("nhomkho").Item(0).InnerText.ToString());
            }
        }

        public string passbhyt
        {
            get
            {
                return "vpbhyt";
            }
        }

        public string pTenBV
        {
            get
            {
                try
                {
                    return this.get_data("select ten from thongso where id=3").Tables[0].Rows[0][0].ToString();
                }
                catch
                {
                    return "";
                }
            }
        }

        public string s_sothe_80_09
        {
            get
            {
                try
                {
                    XmlDocument document = new XmlDocument();
                    document.Load(@"..\..\..\xml\bhxh.xml");
                    return document.GetElementsByTagName("sothe80_09").Item(0).InnerText;
                }
                catch
                {
                    return "UC+TC+VC+YC+XC";
                }
            }
        }

        public string sGiobaocao
        {
            get
            {
                try
                {
                    return this.get_data("select ten from thongso where id=138").Tables[0].Rows[0][0].ToString();
                }
                catch
                {
                    return "00:00";
                }
            }
        }

        public string sNgaytinhmoibhytmoi
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("ngaytinhbhyt").Item(0).InnerText;
            }
        }

        public string Syte
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\maincode.xml");
                return document.GetElementsByTagName("Syte").Item(0).InnerText;
            }
        }

        public string Tenbv
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\maincode.xml");
                return document.GetElementsByTagName("Tenbv").Item(0).InnerText;
            }
        }

        public string thetrongtinh_vitri_old
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return document.GetElementsByTagName("vitri").Item(0).InnerText.ToString();
            }
        }

        public int TOANGOAITRU
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("toangoaitru").Item(0).InnerText.ToString());
            }
        }

        public int tonghopngay
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("tonghopngay").Item(0).InnerText.ToString());
            }
        }

        public int TRE_EM_DUOI_6
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("doituongtreem").Item(0).InnerText.ToString());
            }
        }

        public int updloaibn
        {
            get
            {
                XmlDocument document = new XmlDocument();
                document.Load(@"..\..\..\xml\bhxh.xml");
                return int.Parse(document.GetElementsByTagName("updloaibn").Item(0).InnerText.ToString());
            }
        }

        public string user
        {
            get
            {
                return this.userid;
            }
        }

        public string userbhyt
        {
            get
            {
                return "vpbhyt";
            }
        }
    }
}

