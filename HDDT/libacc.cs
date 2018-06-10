using dllbhyt;
using System;
using System.Data;
using libHDDT;
using LibMedi;
namespace HDDT
{


    public class libacc
    {
        private LibClass _cacc = new LibClass();
        private AccessDataApi m = new AccessDataApi();
        private AccessDataAPI hddt = new AccessDataAPI();
        private string _sUser = "";

        public libacc()
        {
            this._sUser = this.m.user;
        }

        public void f_alter_table(string mmyy)
        {
            string sql = "";
            string str2 = sql;
            sql = str2 + "create table " + this._sUser + mmyy + ".d_thuoc_glivec(MA_BENH_NHAN varchar2(10),MA_THUOC varchar2(20),MA_NHOM varchar2(20),DON_VI_TINH nvarchar2(20),SO_LUONG number(10) default 0,DON_GIA number(15,2) default 0,THANH_TIEN number(15,2) default 0,MA_KHOA varchar2(10),NGAY_YL date, constraint pk_" + mmyy + "_glivec primary key(MA_BENH_NHAN,MA_THUOC,NGAY_YL))";
            try
            {
                this.m.thucThiSql(sql);
            }
            catch
            {
            }
        }

        public DataSet f_get_chiphi(string tungay, string denngay, string mabn, bool vbngoai, bool vbnoi, bool laytheongayxuatkhoa)
        {
            string str = "";
            DataSet set = new DataSet();
            string str2 = this._sUser + str;
            string sql = "";
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str = time3.ToString("MMyy");
                if (this._cacc.bMmyy(str))
                {
                    str2 = this._sUser + str;
                    sql = "";
                    string str4 = "select e.id from " + str2 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.AddMonths(-2).ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    if (vbngoai)
                    {
                        sql = "select b.id,a.maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,case when b1.tungay is null then to_char(a1.tungay,'yyyymmdd') else to_char(b1.tungay,'yyyymmdd') end as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,1 as songaydt,1 as tinhtrangrv,2 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id  inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id left join " + str2 + ".bhyt a1 on a1.maql=a.maql inner join " + this._sUser + ".btdbn a2 on a.mabn=a2.mabn left join " + this._sUser + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + this._sUser + ".btdquan a22 on a2.maqu=a22.maqu left join " + this._sUser + ".btdtt a23 on a2.matt=a23.matt left join " + this._sUser + ".btdkp_bv b2 on b2.makp=b.makp where  b.loaibn in(3,2,4) and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy') and b.bhyt>0" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    if (vbnoi)
                    {
                        string str5 = sql;
                        sql = str5 + ((sql == "") ? "" : " union all ") + " select b.id,a.maql,to_char(a.ngayvao,'yymmddhh24mi')||a.mabn as mavaovien,a.mabn,a2.hoten,case when a2.ngaysinh is null then a2.namsinh else to_char(a2.ngaysinh,'yyyymmdd') end as namsinh,case when a2.phai=0 then 1 else 2 end as phai ,to_char(a.ngayvao,'yyyymmddhh24mi') as ngayvao,to_char(a.ngayra,'yyyymmddhh24mi') as ngayra,to_char(b.ngay,'yyyymmddhh24mi') as ngaythu,0 as sophieu,b.sotien,b.bhyt as bhtra,b.sotien-b.bhyt as bntra,0 bntt,0 nguonkhac,b1.sothe,b1.mabv as madkkcb,to_char(b1.tungay,'yyyymmdd') as tungay,to_char(b1.ngay,'yyyymmdd') as denngay,a.maicd,b.makp,a2.sonha||' '||a2.thon||' '||a21.tenpxa||','||a22.tenquan||','||a23.tentt as diachi,b1.traituyen,a.chandoan,round(to_date(to_char(ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy'))+1 as songaydt,a43.mabhyt2348 as tinhtrangrv,a42.mabhyt2348 as ketqua,b.loaibn,b2.makp_byt,b2.tenkp from " + str2 + ".v_ttrvll b  inner join " + str2 + ".v_ttrvds a on a.id=b.id inner join " + str2 + ".v_ttrvbhyt b1 on b1.id=b.id inner join " + this._sUser + ".btdbn a2 on a.mabn=a2.mabn left join " + this._sUser + ".btdpxa a21 on a2.maphuongxa=a21.maphuongxa left join " + this._sUser + ".btdquan a22 on a2.maqu=a22.maqu left join " + this._sUser + ".btdtt a23 on a2.matt=a23.matt left join " + this._sUser + ".xuatvien a4 on a.maql=a4.maql left join " + this._sUser + ".btdkp_bv b2 on b2.makp=b.makp left join " + this._sUser + ".ketqua a42 on a42.ma=a4.ketqua left join " + this._sUser + ".ttxk a43 on a43.ma=a4.ttlucrv where  b.loaibn in(1) and b.bhyt>0 and b.id not in(" + str4 + ") and to_date(to_char(" + (laytheongayxuatkhoa ? "a.ngayra" : "b.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((mabn == "") ? "" : (" and a.mabn in('" + mabn.Replace(",", "','") + "')"));
                    }
                    try
                    {
                        if (set.Tables[0].Rows.Count > 0)
                        {
                            set.Merge(this._cacc.get_data(sql));
                        }
                        else
                        {
                            set = this._cacc.get_data(sql);
                        }
                    }
                    catch
                    {
                        set = this._cacc.get_data(sql);
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
            string str = "select a.id,a.ma,a.ten,a.dang as dvt,a.donvi,a.kythuat,a.bhyt,e.idnhombhyt as nhombhyt,0 as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,1 as thuoc,a.hamluong,a.sodk,a.maduongdung duongdung,a.tenhc,'' as lieuluong,'' as mavattu,a.masobyt,a.tenbyt,a.gia_bh_toida from " + this._sUser + ".d_dmbd a inner join " + this._sUser + ".d_dmnhom b on a.manhom=b.id inner join " + this._sUser + ".v_nhomvp e on b.nhomvp=e.ma inner join " + this._sUser + ".v_nhombhyt h on e.idnhombhyt=h.id  union all select c.id,cast(c.ma as varchar2(50)) as ma,cast(c.ten as nvarchar2(100)) as ten,c.dvt,null donvi,c.kythuat,c.bhyt,e.idnhombhyt as nhombhyt,d.id as loaivp,h.idnhombhytmedisoft as nhombhytmedi,e.ma as nhomvp,e.mabhyt2348 as manhombhyt,0 as thuoc,null as hamluong,null as sodk,null as duongdung,null tenhc,null as lieuluong,c.mavattubyt as mavattu,c.masobyt,c.tenbyt,c.gia_bh_toida from " + this._sUser + ".v_giavp c inner join " + this._sUser + ".v_loaivp d on c.id_loai=d.id inner join " + this._sUser + ".v_nhomvp e on d.id_nhom=e.ma inner join " + this._sUser + ".v_nhombhyt h on e.idnhombhyt=h.id ";
            string str2 = "";
            string str3 = this._sUser + str2;
            DateTime time = new DateTime(int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)), int.Parse(denngay.Substring(0, 2)));
            time = time.AddMonths(1);
            DateTime time2 = new DateTime(int.Parse(tungay.Substring(6, 4)), int.Parse(tungay.Substring(3, 2)), 1);
            for (DateTime time3 = time2.AddMonths(-1); time3 <= time; time3 = time3.AddMonths(1))
            {
                str2 = time3.ToString("MMyy");
                if (this._cacc.bMmyy(str2))
                {
                    str3 = this._sUser + str2;
                    string str4 = "select e.id from " + str3 + ".v_ttrvll e inner join (select * from (" + this.f_get_sqlfull("select quyenso,sobienlai from xxx.v_hoantra", time3.ToString("dd/MM/yyyy"), time.AddMonths(2).AddMonths(6).ToString("dd/MM/yyyy")) + ")) f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai";
                    string sql = "select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,c2.maql,c.loaibn,a.id,b.id as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(round(a.soluong,2)) as soluong,round(a.dongia,2) dongia,sum(round(round(a.soluong,2)*round(dongia,2),2)) as sotien1,b.loaivp,sum(round(a.bhyttra,2)) as bhyttra,sum(round(round(a.soluong,2)*round(dongia,2),2)-round(a.bhyttra,2)) as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi,case when c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%'  and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%'  then b.bhyt else case when c3.traituyen=1 then b.bhyt else 100 end end as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt,case when b.bhyt<>0 and b.bhyt<>100 and c3.sothe not like 'TE1%' and c3.sothe not like 'CC1%' and c3.sothe not like 'CA5%' and c3.sothe not like 'QN5%' then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  case when c3.traituyen=1 then (round(round(a.dongia,2)*sum(round(a.soluong,2)),2)*b.bhyt)/100 else  round(round(a.dongia,2)*sum(round(a.soluong,2)),2) end end as sotien,0 bntt,0 muchuong,0 nguonkhac from " + str3 + ".v_ttrvct a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + this._sUser + ".btdkp_bv a2 on a2.makp=a.makp inner join (" + str + ") b on a.mavp=b.id where a.madoituong=1  and a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id where    to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy')  between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ") and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,c3.traituyen,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.bhyt,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    if (flag)
                    {
                        string str6 = sql;
                        sql = str6 + " union all select to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn as mavaovien,c2.maql,c.loaibn,a.id,b.id as idvp,b.ma as mavp,b.ten as ten,b.dvt,sum(a.soluong) as soluong,a.dongia,sum(a.soluong*a.dongia) as sotien1,b.loaivp,sum(a.dongia*a.soluong) as bhyttra,0 as bntra,0 as stt,b.manhombhyt,a2.makp_byt as makp,b.hamluong,b.duongdung,b.sodk,tenhc,b.lieuluong,to_char(a.ngay,'yyyymmddhh24mi') as ngayylenh,b.donvi, 100  as bhyt,b.thuoc,b.nhombhytmedi ,b.nhombhytmedi nhombhyt,b.mavattu,b.masobyt,b.tenbyt, sum(a.dongia*a.soluong)  as sotien,0 bntt,0 muchuong, 0  as nguonkhac from d_xuatsd_glivec a inner join " + str3 + ".v_ttrvll c on a.id=c.id left join " + str3 + ".v_ttrvbhyt c3 on c3.id=c.id inner join " + str3 + ".v_ttrvds c2 on a.id=c2.id left join " + this._sUser + ".btdkp_bv a2 on a2.makp=c.makp inner join (" + str + ") b on a.mabd=b.id where  a.id in( select distinct a.id from " + str3 + ".v_ttrvll a inner join " + str3 + ".v_ttrvbhyt b on a.id=b.id where    to_date(to_char(" + (laytheongayxuatkhoa ? "c2.ngayra" : "a.ngay") + ",'dd/mm/yyyy'),'dd/mm/yyyy')  between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')" + ((loaibn == "") ? "" : (" and a.loaibn in(" + loaibn + ")")) + ") and a.id not in(" + str4 + ")" + ((mabn == "") ? "" : (" and c2.mabn in('" + mabn.Replace(",", "','") + "')")) + ((nhomvpbhyt == "") ? "" : (" and b.nhombhytmedi in(" + nhomvpbhyt + ")")) + " group by to_char(c2.ngayvao,'yymmddhh24mi')||c2.mabn ,c2.maql,c.loaibn,a.id,b.id,b.ma,b.ten,b.dvt,a.dongia,b.loaivp,b.manhombhyt,a2.makp_byt,b.donvi,tenhc,c3.sothe,to_char(a.ngay,'yyyymmddhh24mi'),b.hamluong,b.duongdung,b.sodk,b.lieuluong,b.thuoc,b.nhombhytmedi,b.mavattu,b.masobyt,b.tenbyt";
                    }
                    if (laybaocaothongke)
                    {
                        sql = "select loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,sum(soluong) soluong,sum(soluong)*dongia*bhyt/100 sotien,sum(sotien) sotien1 from (" + sql + ") group by loaibn,idvp,mavp,ten,dvt,dongia,manhombhyt,donvi,tenhc,hamluong,duongdung,sodk,lieuluong,nhombhyt,mavattu,masobyt,tenbyt,bhyt";
                    }
                    try
                    {
                        set.Merge(this._cacc.get_data(sql));
                    }
                    catch
                    {
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
            sql = str4 + "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + this._sUser + "d" + maql.Substring(2, 2) + maql.Substring(0, 2) + ".bhytkb WHERE maql in(" + maql + ")";
            try
            {
                foreach (DataRow row in this._cacc.get_data(sql).Tables[0].Rows)
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
                sql = "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + this._sUser + "d" + id.Substring(2, 2) + id.Substring(0, 2) + ".bhytkb WHERE maql =" + maql + " or idttrv=" + id;
                foreach (DataRow row2 in this._cacc.get_data(sql).Tables[0].Rows)
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
            string str5 = str4 + "#select distinct to_char(ngay,'yyyymmddhh24mi') ngay from " + this._sUser + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE maql in(" + maql + ")";
            sql = (str5 + "#select distinct to_char(ngay,'yyyymmddhh24mi') ngay from " + this._sUser + ".benhandt WHERE maql =" + maql + " ").Trim(new char[] { '#' }).Replace("#", " union all ");
            try
            {
                foreach (DataRow row in this._cacc.get_data(sql).Tables[0].Rows)
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
                sql = "select distinct to_char(ngayud,'yyyymmddhh24mi') ngay from " + this._sUser + maql.Substring(2, 2) + maql.Substring(0, 2) + ".benhandt WHERE to_char(maql) like '" + maql.Substring(0, 6) + "%' and mabn='" + mabn + "'";
                foreach (DataRow row2 in this._cacc.get_data(sql).Tables[0].Rows)
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
                if (this._cacc.bMmyy(str))
                {
                    str2 = str2 + vsql.Replace("xxxd.", this._sUser + "d" + str + ".").Replace("xxx.", this._sUser + str + ".") + "#";
                }
            }
            return str2.Trim(new char[] { '#' }).Replace("#", " union all ");
        }

        public string f_get_tungay(string maql)
        {
            string str = "";
            string sql = "";
            string str4 = sql;
            string str5 = str4 + "#select distinct to_char(tungay,'yyyymmdd') tungay from " + this._sUser + maql.Substring(2, 2) + maql.Substring(0, 2) + ".bhyt WHERE maql in(" + maql + ")";
            sql = (str5 + "#select distinct to_char(tungay,'yyyymmdd') tungay from " + this._sUser + ".bhyt WHERE maql =" + maql + " ").Trim(new char[] { '#' }).Replace("#", " union all ");
            try
            {
                foreach (DataRow row in this._cacc.get_data(sql).Tables[0].Rows)
                {
                    if (str == "")
                    {
                        str = row["tungay"].ToString();
                    }
                }
            }
            catch
            {
            }
            return str;
        }

        public void f_upd_thuocGlivec(string m_mabn, string m_mathuoc, string m_manhom, string m_donvitinh, decimal m_soluong, decimal m_dongia, string m_makhoa, string m_ngayylenh)
        {
            try
            {
                this.m.thucThiSql("insert into " + this._sUser + (m_ngayylenh.Substring(3, 2) + m_ngayylenh.Substring(8, 2)) + ".d_thuoc_glivec(MA_BENH_NHAN,MA_THUOC,MA_NHOM,");
            }
            catch (Exception exception)
            {
                this._cacc.f_write_log("update thuoc glivec " + exception.ToString());
            }
        }
    }
}

