using System;
using System.Data;

namespace LibMedi
{
	/// <summary>
	/// Summary description for Export.
	/// </summary>
	public class Export
	{
		private AccessData m=new AccessData();
		private DataColumn dc;
		private DataSet ds;
		private DataTable dt;
		private string sql;
		public Export()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		private void tongcong(string exp,string stt,int k)
		{
			Decimal l_tong=0;
			DataRow[] r1=ds.Tables[0].Select(exp);
			Int16 iRec=Convert.ToInt16(r1.Length);
			for(int j=k;j<ds.Tables[0].Columns.Count;j++)
			{
				l_tong=0;
				for(Int16 i=0;i<iRec;i++) 
				{
					try
					{
						l_tong+=(r1[i][j].ToString()=="")?0:Decimal.Parse(r1[i][j].ToString());
					}
					catch{l_tong+=0;}
				}
				m.updrec_02(ds.Tables[0],stt,j,l_tong);
			}
		}

		private void tongcong(string exp,int ma,int k)
		{
			Decimal l_tong=0;
			DataRow[] r1=ds.Tables[0].Select(exp);
			Int16 iRec=Convert.ToInt16(r1.Length);
			for(int j=k;j<ds.Tables[0].Columns.Count;j++)
			{
				l_tong=0;
				for(Int16 i=0;i<iRec;i++) l_tong+=Decimal.Parse(r1[i][j].ToString());
				m.updrec_145(ds.Tables[0],ma,j,l_tong);
			}
		}

		public DataSet bieu_02(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			ds=new DataSet();
			DataSet ds1=new DataSet();
			DataRow[] dr;
			Int64 c02,c03,c04,c05,c06,c07,ma;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			ds1=m.get_data("select * from btdkp order by makp");
			DataRow r1,r2;
			if (benhan)
			{			
				sql="SELECT a.ngay as ngayvv,a.madoituong,a.nhantu,b.ttlucrv,a.loaiba,";
				sql+="b.ngay as ngayrv,c.makpbo as ma,d.mien"; 
				sql+=" from xxx.benhandt a,xxx.xuatvien b,btdkp_bv c,doituong d,btdbn e,icd10 f ";
				sql+=" where a.maql=b.maql and a.makp=c.makp and a.madoituong=d.madoituong and a.mabn=e.mabn";
				sql+=" and b.maicd=f.cicd10 and length(trim(f.stt))>0";
				sql+=" and a.loaiba=3 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				foreach(DataRow r in m.get_data_mmyy(sql,s_tu1,s_den,false).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						c02=(r["madoituong"].ToString()=="1")?1:0;
						c03=(r["madoituong"].ToString()!="1" && r["mien"].ToString()=="0")?1:0;
						c04=(c02+c03==0)?1:0;
						c05=(r["nhantu"].ToString()=="1")?1:0;
						c06=0;//(r["ttlucrv"].ToString()=="5")?1:0;
						c07=(r["ttlucrv"].ToString()=="6")?1:0;
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+c04;
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
						if (r["loaiba"].ToString()=="2") dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
					}
				}
				//
				sql="SELECT a.ngay as ngayvv,a.madoituong,a.nhantu,b.ttlucrv,a.loaiba,";
				sql+="b.ngay as ngayrv,c.makpbo as ma,d.mien"; 
				sql+=" from benhandt a,xuatvien b,btdkp_bv c,doituong d,btdbn e,icd10 f ";
				sql+=" where a.maql=b.maql and a.makp=c.makp and a.madoituong=d.madoituong and a.mabn=e.mabn";
				sql+=" and b.maicd=f.cicd10 and length(trim(f.stt))>0";
				sql+=" and a.loaiba not in (1,3) and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						c02=(r["madoituong"].ToString()=="1")?1:0;
						c03=(r["madoituong"].ToString()!="1" && r["mien"].ToString()=="0")?1:0;
						c04=(c02+c03==0)?1:0;
						c05=(r["nhantu"].ToString()=="1")?1:0;
						c06=0;//(r["ttlucrv"].ToString()=="5")?1:0;
						c07=(r["ttlucrv"].ToString()=="6")?1:0;
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+c04;
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
						if (r["loaiba"].ToString()=="2") dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
					}
				}

				sql="SELECT c.MAKPBO ma,sum(case when a.khoachuyen='01' then 1 else 0 end) C06";
				sql+=" FROM NHAPKHOA a,BTDKP_BV c,BTDBN d,BENHANDT e ";
				sql+=" WHERE a.MAKP=c.MAKP and a.maql=e.maql and a.mabn=d.mabn and a.maba<20 ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" GROUP BY c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["ma"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r["ma"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						c06=long.Parse(r["c06"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
					}
				}
				//ngaydt
				sql="SELECT c.makpbo as makp,";
				sql+="sum(to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdkp_bv c,btdbn e where  a.maql = b.maql and a.mabn=e.mabn and a.loaiba=2 ";
				sql+="and a.makp=c.makp and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.makp is not null ";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				#region 270107
				/*
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdkp_bv c,btdbn e where  a.maql = b.maql and a.mabn=e.mabn and a.makp=c.makp and a.loaiba=2 ";
				sql+="and (to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+="or to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy'))";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdkp_bv c,btdbn e where  a.mabn=e.mabn and a.maql = b.maql(+) and a.makp=c.makp and a.loaiba=2 ";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<to_date('"+s_tu1+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdkp_bv c,btdbn e where  a.mabn=e.mabn and a.maql = b.maql and a.makp=c.makp and a.loaiba=2 ";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null ";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				*/
				#endregion
				tongcong("ma>1","A",3);
			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09";
				sql+=" from bieu_02 a,dm_02 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09=0");
			ds.AcceptChanges();
			return ds;
		}


		public DataSet bieu_02_bv(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			ds=new DataSet();
			bool bBieu2_nhapvien_la_nhapkhoa=m.bBieu2_nhapvien_la_nhapkhoa;
			DataRow[] dr;
			DataRow r1,r2;
			Int64 c02,c03,c04,c05,c06,c07,ma;
			ds=m.get_data("select * from dm_02 where ma<2 order by ma");
			DataSet ds1=new DataSet();
			ds1=m.get_data("select * from btdkp_bv order by makp");			
			if (benhan)
			{			

				sql="SELECT a.ngay as ngayvv,a.madoituong,a.nhantu,b.ttlucrv,a.loaiba,";
				sql+="b.ngay as ngayrv,c.idkp as ma,c.tenkp,d.mien"; 
				sql+=" from xxx.benhandt a,xxx.xuatvien b,btdkp_bv c,doituong d,btdbn e,icd10 f ";
				sql+=" where a.maql=b.maql and a.makp=c.makp and a.madoituong=d.madoituong and a.mabn=e.mabn";
				sql+=" and b.maicd=f.cicd10 and length(trim(f.stt))>0";
				sql+=" and a.loaiba=3 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				DataSet lds=m.get_data_mmyy(sql,s_tu1,s_den,false);
				if(lds!=null && lds.Tables.Count>0)
				{
					foreach(DataRow r in lds.Tables[0].Rows)
					{
						ma=int.Parse(r["ma"].ToString());
						r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
						if (r1==null) m.updrec_bieu(ds.Tables[0],int.Parse(r["ma"].ToString()),r["tenkp"].ToString(),9);
						dr=ds.Tables[0].Select("ma="+ma);
						if (dr.Length>0)
						{
							c02=(r["madoituong"].ToString()=="1")?1:0;
							c03=(r["madoituong"].ToString()!="1" && r["mien"].ToString()=="0")?1:0;
							c04=(c02+c03==0)?1:0;
							c05=(r["nhantu"].ToString()=="1")?1:0;
							c06=(bBieu2_nhapvien_la_nhapkhoa==false && r["ttlucrv"].ToString()=="5")?1:0;//chi dinh nhap vien tinh la nhap vien
							c07=(r["ttlucrv"].ToString()=="6")?1:0;
							dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
							dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
							dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
							dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+c04;
							dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
							dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
							dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
							if (r["loaiba"].ToString()=="2")
								dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
						}
					}
				}

				sql="SELECT a.ngay as ngayvv,a.madoituong,a.nhantu,b.ttlucrv,a.loaiba,";
				sql+="b.ngay as ngayrv,c.idkp as ma,c.tenkp,d.mien"; 
				sql+=" from benhandt a,xuatvien b,btdkp_bv c,doituong d,btdbn e,icd10 f ";
				sql+=" where a.maql=b.maql and a.makp=c.makp and a.madoituong=d.madoituong and a.mabn=e.mabn";
				sql+=" and b.maicd=f.cicd10 and length(trim(f.stt))>0";
				sql+=" and a.loaiba not in (1,3) and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				//foreach(DataRow r in m.get_data_mmyy(sql,s_tu1,s_den,false).Tables[0].Rows)
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],int.Parse(r["ma"].ToString()),r["tenkp"].ToString(),9);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						c02=(r["madoituong"].ToString()=="1")?1:0;
						c03=(r["madoituong"].ToString()!="1" && r["mien"].ToString()=="0")?1:0;
						c04=(c02+c03==0)?1:0;
						c05=(r["nhantu"].ToString()=="1")?1:0;
						c06=(bBieu2_nhapvien_la_nhapkhoa==false && r["ttlucrv"].ToString()=="5")?1:0;//chi dinh nhap vien tinh la nhap vien
						c07=(r["ttlucrv"].ToString()=="6")?1:0;
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+c04;
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
						if (r["loaiba"].ToString()=="2")
							dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
					}
				}
				if(bBieu2_nhapvien_la_nhapkhoa)
				{
					sql="SELECT c.idKP ma,sum(case when a.khoachuyen='01' then 1 else 0 end) C06";
					sql+=" FROM NHAPKHOA a,BTDKP_BV c,BTDBN d,BENHANDT e ";
					sql+=" WHERE a.MAKP=c.MAKP and a.maql=e.maql and a.mabn=d.mabn and a.maba<20 ";
					sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
					sql+=" GROUP BY c.idKP";
					foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
					{
						ma=int.Parse(r["ma"].ToString());
						r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
						if (r1==null)
						{
							r2=m.getrowbyid(ds1.Tables[0],"idkp="+r["ma"].ToString());//"makp='"+r["ma"].ToString().PadLeft(2,'0')+"'");
							if (r2!=null)
								m.updrec_bieu(ds.Tables[0],int.Parse(r["ma"].ToString()),r2["tenkp"].ToString(),9);
						}
						dr=ds.Tables[0].Select("ma="+ma);
						if (dr.Length>0)
						{
							c06=long.Parse(r["c06"].ToString());
							dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
						}
					}
				}
				//ngaydt
				sql="SELECT d.idkp as makp,";
				sql+="sum(to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdbn e, btdkp_bv d where  a.maql = b.maql and a.mabn=e.mabn and a.makp=d.makp and a.loaiba=2 ";
				sql+="and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.makp is not null ";
				sql+="group by d.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"idkp="+r["makp"].ToString());//,"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				#region 270107_bo
				/*
				sql="SELECT a.makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdbn e where  a.maql = b.maql and a.mabn=e.mabn and a.loaiba=2 ";
				sql+="and (to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy') ";
				sql+="or to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')) ";
				sql+=" and a.makp is not null ";
				sql+="group by a.makp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT a.makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdbn e where  a.maql = b.maql and a.mabn=e.mabn and a.loaiba=2 ";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<to_date('"+s_tu1+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.makp is not null ";
				sql+="group by a.makp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT a.makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+1) c15 ";
				sql+="FROM benhandt a,xuatvien b,btdbn e where  a.mabn=e.mabn and a.maql = b.maql and a.loaiba=2 ";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null ";
				sql+=" and a.makp is not null ";
				sql+="group by a.makp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null)
							m.updrec_bieu(ds.Tables[0],int.Parse(r2["makp"].ToString()),r2["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				*/
				#endregion
				tongcong("ma>1","A",3);
			}
			#region 29102004
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09";
				sql+=" from bieu_02 a where to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString());
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						//r1=m.getrowbyid(ds1.Tables[0],"makp='"+ma.ToString().PadLeft(2,'0')+"'");
						r1=m.getrowbyid(ds1.Tables[0],"idkp='"+ma.ToString()+"'");
						if (r1!=null) m.updrec_bieu(ds.Tables[0],int.Parse(r1["idkp"].ToString()),r1["tenkp"].ToString(),9);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
					}
				}
			}
			
			#endregion
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet bieu_031(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh,bool benhvien)
		{
			int iadd=m.iNgaydieutri_ngayra_ngayvao_1;
			DataRow[] dr;
			DataRow r1,r2;
			Int64 c08,c09,ma,i_makp;
			DataColumn dc;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			ds.AcceptChanges();
			dc=new DataColumn();
			dc.ColumnName="C12";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C13";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C14";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C15";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			foreach(DataRow r in ds.Tables[0].Rows)
			{
				r["c12"]=0;r["c13"]=0;r["c14"]=0;r["c15"]=0;
			}
			DataSet ds1=new DataSet();
			ds1=m.get_data("select * from btdkp order by makp");
			if (benhan)
			{	
				sql="SELECT a.khoachuyen,a.ngay as ngayvv,b.ttlucrk as ttlucrv,c.namsinh,";
				sql+="b.ngay as ngayrv,d.makpbo as ma,c.phai";
				sql+=" FROM ((nhapkhoa a inner JOIN xuatkhoa b ON a.id=b.id)";
				sql+=" INNER JOIN BTDBN c ON a.MABN=c.MABN) INNER JOIN BTDKP_BV d ON a.MAKP=d.MAKP";
				sql+=" where  to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.maba<20 and b.ttlucrk=7 ";
				if (benhvien) sql+=" and d.kehoach>0";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						c08=(DateTime.Now.Year-int.Parse(r["namsinh"].ToString())<15)?1:0;
						c09=(m.DateToString("dd/MM/yyyy",DateTime.Parse(r["ngayvv"].ToString()))==m.DateToString("dd/MM/yyyy",DateTime.Parse(r["ngayrv"].ToString())))?1:0;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+1;
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+c08;
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
					}
				}
				//phai=1
				sql="SELECT a.khoachuyen,a.ngay as ngayvv,b.ttlucrk as ttlucrv,c.namsinh,";
				sql+="b.ngay as ngayrv,d.makpbo as ma,c.phai";
				sql+=" FROM ((nhapkhoa a inner JOIN xuatkhoa b ON a.id=b.id)";
				sql+=" INNER JOIN BTDBN c ON a.MABN=c.MABN) INNER JOIN BTDKP_BV d ON a.MAKP=d.MAKP";
				sql+=" where  to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.maba<20 and c.phai=1 and b.ttlucrk=7 and a.khoachuyen='01'";
				if (benhvien) sql+=" and d.kehoach>0";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["ma"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						c08=(DateTime.Now.Year-int.Parse(r["namsinh"].ToString())<15)?1:0;
						c09=(m.DateToString("dd/MM/yyyy",DateTime.Parse(r["ngayvv"].ToString()))==m.DateToString("dd/MM/yyyy",DateTime.Parse(r["ngayrv"].ToString())))?1:0;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+1;
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+c08;
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
					}
				}
				//dau ky + cuoi ky
				sql="SELECT c.MAKPBO ma,sum(case when (to_date(a.ngay,'dd/mm/yy')<to_date('"+s_tu1+"','dd/mm/yy') and (b.ngay is null or to_date(b.ngay,'dd/mm/yy')>=to_date('"+s_tu1+"','dd/mm/yy'))) then 1 else 0 end) C03,";
				sql+="sum(case when a.khoachuyen='01' and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C04,";
				sql+="sum(case when a.khoachuyen<>'01' and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C05,";
				sql+="sum(case when a.khoachuyen='01' and to_char(sysdate,'yyyy')-d.namsinh<15 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C041,";
				sql+="sum(case when a.khoachuyen<>'01' and to_char(sysdate,'yyyy')-d.namsinh<15 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C042,";
				sql+="sum(case when a.khoachuyen='01' and e.nhantu=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C051,";
				sql+="sum(case when a.khoachuyen<>'01' and e.nhantu=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C052,";
				sql+="sum(case when a.khoachuyen='01' and e.madoituong=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C101,";
				sql+="sum(case when a.khoachuyen<>'01' and e.madoituong=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C102,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=1 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C06,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=2 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C07,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=3 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C08,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=4 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C09,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=5 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C10,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=6 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C11,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=7 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C12,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C13";
				sql+=" FROM NHAPKHOA a,XUATKHOA b,BTDKP_BV c,BTDBN d,BENHANDT e ";
				sql+=" WHERE a.MAKP=c.MAKP and a.ID=b.ID(+) and a.maql=e.maql and a.mabn=d.mabn and a.maba<20 ";
				if (benhvien) sql+=" and c.kehoach>0";
				foreach(DataRow r in m.get_data(sql+"GROUP BY c.makpbo ORDER BY c.makpbo").Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,m.get_data("select tenkp from btdkp where makp='"+r["ma"].ToString()+"'").Tables[0].Rows[0][0].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						dr[0]["c02"]=Decimal.Parse(r["c03"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c041"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c051"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c101"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c13"]=Decimal.Parse(dr[0]["c13"].ToString())+Decimal.Parse(r["c042"].ToString());
						dr[0]["c14"]=Decimal.Parse(dr[0]["c14"].ToString())+Decimal.Parse(r["c052"].ToString());
						dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c102"].ToString());
						dr[0]["c11"]=Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())+Decimal.Parse(r["c05"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c10"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					}
				}
				//phai=1
				sql+=" and d.phai=1 and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql+"GROUP BY c.makpbo ORDER BY c.makpbo").Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,m.get_data("select tenkp from btdkp where makp='"+r["ma"].ToString()+"'").Tables[0].Rows[0][0].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						dr[0]["c02"]=decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c041"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c051"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c101"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c13"]=Decimal.Parse(dr[0]["c13"].ToString())+Decimal.Parse(r["c042"].ToString());
						dr[0]["c14"]=Decimal.Parse(dr[0]["c14"].ToString())+Decimal.Parse(r["c052"].ToString());
						dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c102"].ToString());
						dr[0]["c11"]=decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())+Decimal.Parse(r["c05"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c10"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					}
				}
				/*
				#region linh 15/11/2007
				//ngaydt linh 15/11/2007
				sql="select e.makpbo makp,sum(";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) end end end end) c15";
				//nu
				sql+=",sum(decode(d.phai,1,";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) end end end end,0)) c15_nu";
				sql+=" from nhapkhoa a,xuatkhoa b,benhandt c,btdbn d,btdkp_bv e ";
				sql+=" where a.id=b.id(+) and a.maql=c.maql and a.makp=e.makp and a.maba<20 and c.loaiba=1 and a.mabn=d.mabn and a.makp is not null ";
				sql+=" and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') <=to_date('"+s_den+"','dd/mm/yyyy') and to_date(to_char(nvl(b.ngay,sysdate),'dd/mm/yyyy'),'dd/mm/yyyy') >=to_date('"+s_tu+"','dd/mm/yyyy')";
				if (benhvien) sql+=" and e.kehoach>0";
				sql+=" group by e.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
					//nu
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15_nu"].ToString());
				}
				//end linh
				*/
				/*
				//xuattam
				sql="select e.makpbo makp,sum(";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy') else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') end end end end) c15";
				//nu
				sql+=",sum(decode(d.phai,1,";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy') else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') end end end end,0)) c15_nu";
				sql+=" from nhapkhoa a,xuattam b,benhandt c,btdbn d,btdkp_bv e ";
				sql+=" where a.id=b.idkhoa and a.maql=c.maql and a.makp=e.makp and a.maba<20 and c.loaiba=1 and a.mabn=d.mabn and a.makp is not null ";
				sql+=" and to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') <=to_date('"+s_den+"','dd/mm/yyyy') and to_date(to_char(nvl(b.ngayvao,sysdate),'dd/mm/yyyy'),'dd/mm/yyyy') >=to_date('"+s_tu+"','dd/mm/yyyy')";
				if (benhvien) sql+=" and e.kehoach>0";
				sql+=" group by e.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
					//nu
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15_nu"].ToString());
				}
				*/
				//ngaydt
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";//sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID and a.makp=c.makp and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID and a.makp=c.makp and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID(+) and a.makp=c.makp and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null ";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				//ngaydt-nu
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID(+) and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				//xuat tam
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID and a.makp=c.makp and a.maba<20 ";
				sql+=" and to_date(b.ngayra,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID and a.makp=c.makp and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.makpbo makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID(+) and a.makp=c.makp and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngayra is null ";
				sql+="group by c.makpbo";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				//ngaydt-nu
				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 ";
				sql+=" and to_date(b.ngayra,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID(+) and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngayra is null and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				//
				int i_ma;
				foreach(DataRow r in m.get_data("select makpbo,sum(kehoach) as kh from btdkp_bv group by makpbo order by makpbo").Tables[0].Rows)
				{
					i_ma=int.Parse(r["makpbo"].ToString())+2;
					r2=m.getrowbyid(ds.Tables[0],"ma="+i_ma);
					if (r2!=null) r2["c01"]=r["kh"].ToString();
				}
				tongcong("ma>2","1",3);
				r1=m.getrowbyid(ds.Tables[0],"ma=1");
				if (r1!=null)
				{
					r1["c12"]=0;r1["c13"]=0;r1["c14"]=0;r1["c15"]=0;
				}
				foreach(DataRow r in ds.Tables[0].Rows)
				{
					r["c03"]=decimal.Parse(r["c03"].ToString())+decimal.Parse(r["c12"].ToString());
					r["c04"]=decimal.Parse(r["c04"].ToString())+decimal.Parse(r["c13"].ToString());
					r["c05"]=decimal.Parse(r["c05"].ToString())+decimal.Parse(r["c14"].ToString());
					r["c10"]=decimal.Parse(r["c10"].ToString())+decimal.Parse(r["c15"].ToString());
				}
				
			}
			
			if (thongke)
			{
				DataSet ds2=new DataSet();
				ds2=m.get_data("select a.makp,a.tenkp,b.makp as ma from btdkp a,btdkp_bv b where a.makp=b.makpbo and b.loai=1 order by a.makp");
				sql="select ma,sum(c01) as c01,sum(c02) as c02,sum(c03) as c03,sum(c04) as c04,";
				sql+="sum(c05) as c05,sum(c06) as c06,sum(c07) as c07,sum(c08) as c08,";
				sql+="sum(c09) as c09,sum(c10) as c10,sum(c11) as c11";
				sql+=" from bieu_031 where to_date(ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				if (benhvien) sql+=" and b.kehoach>0";
				sql+=" group by ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString());
					if (ma>2)
					{
						i_makp=ma-2;
						r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
						if (r1==null)
						{
							r1=m.getrowbyid(ds2.Tables[0],"ma='"+i_makp.ToString().PadLeft(2,'0')+"'");
							if (r1!=null)
								m.updrec_bieu(ds.Tables[0],ma,r1["tenkp"].ToString(),15);
						}
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						//dr[0]["c01"]=r["c01"].ToString();
						//dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
						//dr[0]["c11"]=Decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c11"].ToString());
					}
				}
				sql="select id from bieu_031 where to_date(ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" order by ngay";
				ds1=m.get_data(sql);
				if (ds1.Tables[0].Rows.Count!=0)
				{
					long id1=long.Parse(ds1.Tables[0].Rows[0][0].ToString());
					long id2=long.Parse(ds1.Tables[0].Rows[ds1.Tables[0].Rows.Count-1][0].ToString());
					foreach(DataRow r in m.get_data("select ma,c01,c02 from bieu_031 where id="+id1).Tables[0].Rows)
					{
						r1=m.getrowbyid(ds.Tables[0],"ma="+int.Parse(r["ma"].ToString()));
						if (r1!=null)
						{
							if (decimal.Parse(r["c01"].ToString())>0) r1["c01"]=decimal.Parse(r["c01"].ToString());
							r1["c02"]=decimal.Parse(r1["c02"].ToString())+decimal.Parse(r["c02"].ToString());
						}
					}
					foreach(DataRow r in m.get_data("select ma,c11 from bieu_031 where id="+id2).Tables[0].Rows)
					{
						r1=m.getrowbyid(ds.Tables[0],"ma="+int.Parse(r["ma"].ToString()));
						if (r1!=null) r1["c11"]=decimal.Parse(r1["c11"].ToString())+decimal.Parse(r["c11"].ToString());
					}
				}
			}
			
			if (phatsinh) m.delrec(ds.Tables[0],"c02+c03+c04+c05+c06+c07+c08+c09+c10+c11=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet bieu_031_bv(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			int iadd=m.iNgaydieutri_ngayra_ngayvao_1;
			DataRow[] dr;
			DataRow r1,r2;
			Int64 c08,c09,ma,i_makp;
			DataColumn dc;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
//			m.delrec(ds.Tables[0],"ma>3");
			ds.AcceptChanges();
			dc=new DataColumn();
			dc.ColumnName="C12";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C13";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C14";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C15";
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			foreach(DataRow r in ds.Tables[0].Rows)
			{
				r["c12"]=0;r["c13"]=0;r["c14"]=0;r["c15"]=0;
			}
			DataSet ds1=new DataSet();
			ds1=m.get_data("select * from btdkp_bv where kehoach>0 order by makp");
			if (benhan)
			{	
				sql="SELECT a.khoachuyen,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngayvv,b.ttlucrk as ttlucrv,c.namsinh,";
				sql+="to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayrv,d.idkp as ma,c.phai,d.tenkp";
				sql+=" FROM nhapkhoa a,xuatkhoa b,BTDBN c,BTDKP_BV d,benhandt e ";
				sql+=" where a.id=b.id and a.mabn=c.mabn and a.makp=d.makp and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.maba<20 and b.ttlucrk=7 and a.maql=e.maql and e.loaiba=1 ";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					//Thuy 20.03.2012
					ma=int.Parse(r["ma"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["ma"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					//ds.WriteXml("xx.xml",XmlWriteMode.WriteSchema);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						try{c08=(DateTime.Now.Year-int.Parse(r["namsinh"].ToString())<15)?1:0;}
						catch{c08=0;}
						if (m.songay(m.StringToDate(r["ngayrv"].ToString().Substring(0,10)),m.StringToDate(r["ngayvv"].ToString().Substring(0,10)),0)==1 && int.Parse(r["ngayrv"].ToString().Substring(11,2))<int.Parse(r["ngayvv"].ToString().Substring(11,2))) c09=1;
						else c09=(r["ngayvv"].ToString().Substring(0,10)==r["ngayrv"].ToString().Substring(0,10))?1:0;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+1;
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+c08;
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
					}
				}
				//phai=1
				sql="SELECT a.khoachuyen,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngayvv,b.ttlucrk as ttlucrv,c.namsinh,";
				sql+="to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayrv,d.idkp as ma,c.phai";
				sql+=" FROM nhapkhoa a,xuatkhoa b,BTDBN c,BTDKP_BV d,benhandt e ";
				sql+=" where  a.id=b.id and a.mabn=c.mabn and a.makp=d.makp and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and d.kehoach>0 and a.maba<20 and b.ttlucrk=7 and a.maql=e.maql and e.loaiba=1 and c.phai=1 and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["ma"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						try{c08=(DateTime.Now.Year-int.Parse(r["namsinh"].ToString())<15)?1:0;}
						catch{c08=0;}
						if (m.songay(m.StringToDate(r["ngayrv"].ToString().Substring(0,10)),m.StringToDate(r["ngayvv"].ToString().Substring(0,10)),0)==1 && int.Parse(r["ngayrv"].ToString().Substring(11,2))<int.Parse(r["ngayvv"].ToString().Substring(11,2))) c09=1;
						else c09=(r["ngayvv"].ToString().Substring(0,10)==r["ngayrv"].ToString().Substring(0,10))?1:0;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+1;
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+c08;
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
					}
				}
				//dau ky + cuoi ky
				sql="SELECT c.idKP ma,sum(case when (to_date(a.ngay,'dd/mm/yy')<to_date('"+s_tu1+"','dd/mm/yy') and (b.ngay is null or to_date(b.ngay,'dd/mm/yy')>=to_date('"+s_tu1+"','dd/mm/yy'))) then 1 else 0 end) C03,";
				sql+="sum(case when a.khoachuyen='01' and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C04,";
				sql+="sum(case when a.khoachuyen<>'01' and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C05,";
				sql+="sum(case when a.khoachuyen='01' and to_char(sysdate,'yyyy')-d.namsinh<15 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C041,";
				sql+="sum(case when a.khoachuyen<>'01' and to_char(sysdate,'yyyy')-d.namsinh<15 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C042,";
				sql+="sum(case when a.khoachuyen='01' and e.nhantu=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C051,";
				sql+="sum(case when a.khoachuyen<>'01' and e.nhantu=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C052,";
				sql+="sum(case when a.khoachuyen='01' and e.madoituong=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C101,";
				sql+="sum(case when a.khoachuyen<>'01' and e.madoituong=1 and to_date(a.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C102,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=1 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C06,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=2 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C07,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=3 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C08,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=4 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C09,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=5 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C10,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=6 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C11,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=7 And to_date(b.ngay,'dd/mm/yy') Between to_date('"+s_tu1+"','dd/mm/yy') And to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end end) C12,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') then 1 else 0 end) C13";
				sql+=" FROM NHAPKHOA a,XUATKHOA b,BTDKP_BV c,BTDBN d,benhandt e ";
				sql+=" WHERE a.MAKP=c.MAKP and a.ID=b.ID(+) and a.mabn=d.mabn and a.maql=e.maql and c.kehoach>0 and a.maba<20 and e.loaiba=1";
				
				foreach(DataRow r in m.get_data(sql+"GROUP BY c.idkp ORDER BY c.idkp").Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,m.get_data("select tenkp from btdkp_bv where idkp='"+r["ma"].ToString()+"'").Tables[0].Rows[0][0].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						dr[0]["c02"]=Decimal.Parse(r["c03"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c041"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c051"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c101"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c13"]=Decimal.Parse(dr[0]["c13"].ToString())+Decimal.Parse(r["c042"].ToString());
						dr[0]["c14"]=Decimal.Parse(dr[0]["c14"].ToString())+Decimal.Parse(r["c052"].ToString());
						dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c102"].ToString());
						dr[0]["c11"]=Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())+Decimal.Parse(r["c05"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c10"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					}
				}
				
				//phai=1
				sql+=" and d.phai=1 and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql+"GROUP BY c.idkp ORDER BY c.idkp").Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,m.get_data("select tenkp from btdkp_bv where idkp='"+r["ma"].ToString()+"'").Tables[0].Rows[0][0].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						dr[0]["c02"]=decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c041"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c051"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c101"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c13"]=Decimal.Parse(dr[0]["c13"].ToString())+Decimal.Parse(r["c042"].ToString());
						dr[0]["c14"]=Decimal.Parse(dr[0]["c14"].ToString())+Decimal.Parse(r["c052"].ToString());
						dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c102"].ToString());
						dr[0]["c11"]=decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())+Decimal.Parse(r["c05"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c10"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					}
				}
				
				#region linh 15/11/2007 - bo
				/*
				//ngaydt linh 15/11/2007
				sql="select a.makp,sum(";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) end end end end) c15";
				//nu
				sql+=",sum(decode(d.phai,1,";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngay is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')+decode(b.ttlucrk,5,0,6,0,7,0,1) end end end end,0)) c15_nu";
				sql+=" from nhapkhoa a,xuatkhoa b,benhandt c,btdbn d ";
				sql+=" where a.id=b.id(+) and a.maql=c.maql and a.maba<20 and c.loaiba=1 and a.mabn=d.mabn and a.makp is not null ";
				sql+=" and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') <=to_date('"+s_den+"','dd/mm/yyyy') and to_date(to_char(nvl(b.ngay,sysdate),'dd/mm/yyyy'),'dd/mm/yyyy') >=to_date('"+s_tu+"','dd/mm/yyyy')";
				sql+=" group by a.makp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
					//nu
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15_nu"].ToString());
				}
				//end linh
				//ngaydt
				//xuattam
				sql="select a.makp,sum(";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy') else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') end end end end) c15";
				//nu
				sql+=",sum(decode(d.phai,1,";
				//vao<tu and (den<ra or ra is null) => den-tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy')+1 else";
				//vao<tu and (ra<=den)              => ra - tu
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')< to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date('"+s_tu+"','dd/mm/yyyy') else ";
				//vao>=tu and (den<ra or ra is null)=> den-vao
				sql+=" case when to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and (to_date('"+s_den+"','dd/mm/yyyy')< to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy') or b.ngayvao is null) then to_date('"+s_den+"','dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')+1 else ";
				//vao>=tu and (ra<=den)             => ra - vao
				sql+=" case when to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')>= to_date('"+s_tu+"','dd/mm/yyyy') ";
				sql+=" and to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('"+s_den+"','dd/mm/yyyy') then to_date(to_char(b.ngayvao,'dd/mm/yyyy'),'dd/mm/yyyy')-to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') end end end end,0)) c15_nu";
				sql+=" from nhapkhoa a,xuattam b,benhandt c,btdbn d ";
				sql+=" where a.id=b.idkhoa and a.maql=c.maql and a.maba<20 and c.loaiba=1 and a.mabn=d.mabn and a.makp is not null ";
				sql+=" and to_date(to_char(b.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') <=to_date('"+s_den+"','dd/mm/yyyy') and to_date(to_char(nvl(b.ngayvao,sysdate),'dd/mm/yyyy'),'dd/mm/yyyy') >=to_date('"+s_tu+"','dd/mm/yyyy')";
				sql+=" group by a.makp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
					//nu
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
					if (r1==null)
					{
						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15_nu"].ToString());
				}
				*/
				#endregion linh_ngaydieutri
				#region ngay_dieutri
				sql="SELECT c.idkp makp,";
				//ra null or ra > den  ==> DEN = den else DEN = ra
				//and vao> tu ==> TU = vao else TU = tu 
				//c15 = DEN - TU + (1: khi khoa dau la 01 va ngay vao >= tu, hoac +1: khi ngay vao < tu; nguoc lai +0
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end +(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end)) c15 ";//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID and a.makp=c.makp and c.kehoach>0 and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.idkp as makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end +(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID and a.makp=c.makp and c.kehoach>0 and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.idkp as makp,";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdkp_bv c,benhandt d where  a.maql=d.maql and a.ID = b.ID(+) and a.makp=c.makp and c.kehoach>0 and a.maba<20 and d.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null ";
				sql+="group by c.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				//ngaydt-nu
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end +(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID and a.mabn=c.mabn and d.kehoach>0 and a.makp=d.makp and c.phai=1 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and d.kehoach>0 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngay is null or to_date(b.ngay,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATKHOA b,btdbn c,btdkp_bv d,benhandt e where  a.maql=e.maql and a.ID = b.ID(+) and a.mabn=c.mabn and d.kehoach>0 and a.makp=d.makp and c.phai=1 and a.maba<20 and e.loaiba=1";
				sql+=" and to_date(a.ngay,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngay is null and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
				//xuat tam
				sql="SELECT c.idkp as makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID and a.makp=c.makp and c.kehoach>0 and a.maba<20 ";
				sql+=" and to_date(b.ngayra,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.idkp as makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID and a.makp=c.makp and c.kehoach>0 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+="group by c.idkp ";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT c.idkp as makp,";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',1,0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdkp_bv c where  a.ID = b.ID(+) and a.makp=c.makp and c.kehoach>0 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngayra is null ";
				sql+="group by c.idkp";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["makp"].ToString())+2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null) m.updrec_bieu(ds.Tables[0],ma,r["tenkp"].ToString(),15);
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0) dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				//ngaydt-nu
				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and d.kehoach>0  and a.maba<20 ";
				sql+=" and to_date(b.ngayra,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and d.kehoach>0 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}

				sql="SELECT ";
				sql+="sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+s_den+"','dd/mm/yy') ";
				sql+="then to_date('"+s_den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
				sql+="case when to_date(b.ngayvao,'dd/mm/yy')>to_date('"+s_tu1+"','dd/mm/yy')";
				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+s_tu1+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + s_tu1 + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',1,0)) c15 ";
				sql+="FROM NHAPKHOA a,XUATtam b,btdbn c,btdkp_bv d where  a.ID = b.ID(+) and a.mabn=c.mabn and a.makp=d.makp and c.phai=1 and d.kehoach>0 and a.maba<20 ";
				sql+=" and to_date(b.ngayvao,'dd/mm/yy')<=to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.ngayra is null and a.khoachuyen='01'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=2;
					r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
//					if (r1==null)
//					{
//						r2=m.getrowbyid(ds1.Tables[0],"makp='"+r["makp"].ToString()+"'");
//						if (r2!=null) m.updrec_bieu(ds.Tables[0],ma,r2["tenkp"].ToString(),15);
//					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0 && r["c15"].ToString()!="") dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())-Decimal.Parse(r["c15"].ToString());
				}
				#endregion ngay_dieutri
				//
				int i_ma;
				foreach(DataRow r in m.get_data("select idkp as makp, sum(kehoach) as kh from btdkp_bv group by idkp order by makp").Tables[0].Rows)
				{
					i_ma=int.Parse(r["makp"].ToString())+2;
					r2=m.getrowbyid(ds.Tables[0],"ma="+i_ma);
					if (r2!=null) r2["c01"]=r["kh"].ToString();
				}
				tongcong("ma>2","1",3);
				r1=m.getrowbyid(ds.Tables[0],"ma=1");
				if (r1!=null)
				{
					r1["c12"]=0;r1["c13"]=0;r1["c14"]=0;r1["c15"]=0;
				}
				foreach(DataRow r in ds.Tables[0].Rows)
				{
					r["c03"]=decimal.Parse(r["c03"].ToString())+decimal.Parse(r["c12"].ToString());
					r["c04"]=decimal.Parse(r["c04"].ToString())+decimal.Parse(r["c13"].ToString());
					r["c05"]=decimal.Parse(r["c05"].ToString())+decimal.Parse(r["c14"].ToString());
					r["c10"]=decimal.Parse(r["c10"].ToString())+decimal.Parse(r["c15"].ToString());
				}
			}
			#region 29102004
			
			if (thongke)
			{
				sql="select ma,sum(c01) as c01,sum(c02) as c02,sum(c03) as c03,sum(c04) as c04,";
				sql+="sum(c05) as c05,sum(c06) as c06,sum(c07) as c07,sum(c08) as c08,";
				sql+="sum(c09) as c09,sum(c10) as c10,sum(c11) as c11";
				sql+=" from bieu_031 where to_date(ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					ma=int.Parse(r["ma"].ToString());
					if (ma>2)
					{
						i_makp=ma-2;
						r1=m.getrowbyid(ds.Tables[0],"ma="+ma);
						if (r1==null) m.updrec_bieu(ds.Tables[0],ma,m.get_data("select tenkp from btdkp_bv where makp='"+i_makp.ToString().PadLeft(2,'0')+"'").Tables[0].Rows[0][0].ToString(),15);
					}
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						//dr[0]["c01"]=r["c01"].ToString();
						//dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
						//dr[0]["c11"]=Decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c11"].ToString());
					}
				}
				sql="select id from bieu_031 where to_date(ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" order by ngay";
				ds1=m.get_data(sql);
				if (ds1.Tables[0].Rows.Count!=0)
				{
					long id1=long.Parse(ds1.Tables[0].Rows[0][0].ToString());
					long id2=long.Parse(ds1.Tables[0].Rows[ds1.Tables[0].Rows.Count-1][0].ToString());
					foreach(DataRow r in m.get_data("select ma,c01,c02 from bieu_031 where id="+id1).Tables[0].Rows)
					{
						r1=m.getrowbyid(ds.Tables[0],"ma="+int.Parse(r["ma"].ToString()));
						if (r1!=null)
						{
							r1["c01"]=decimal.Parse(r1["c01"].ToString())+decimal.Parse(r["c01"].ToString());
							r1["c02"]=decimal.Parse(r1["c02"].ToString())+decimal.Parse(r["c02"].ToString());
						}
					}
					foreach(DataRow r in m.get_data("select ma,c11 from bieu_031 where id="+id2).Tables[0].Rows)
					{
						r1=m.getrowbyid(ds.Tables[0],"ma="+int.Parse(r["ma"].ToString()));
						if (r1!=null) r1["c11"]=decimal.Parse(r1["c11"].ToString())+decimal.Parse(r["c11"].ToString());
					}
				}
			}
			#endregion
			if (phatsinh) m.delrec(ds.Tables[0],"c02+c03+c04+c05+c06+c07+c08+c09+c10+c11=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet bieu_04(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			DataRow[] dr;
			int c02,c03,c05,c06,c07,c09,c10,ma,so;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			if (benhan)
			{	
				sql="SELECT a.mapt,b.loaipt,a.tinhhinh,a.taibien,a.tuvong";
				sql+=" FROM PTTT a,DMPTTT b where a.MAPT = b.MAPT";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					so=(r["mapt"].ToString().Substring(0,1)=="P")?1:10;
					ma=int.Parse(r["loaipt"].ToString())+so;
					dr=ds.Tables[0].Select("ma="+ma);
					if (dr.Length>0)
					{
						c02=(r["tinhhinh"].ToString()=="2")?1:0;
						c03=(r["tinhhinh"].ToString()!="2")?1:0;
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
						if (r["taibien"].ToString()!="0")
						{
							c05=(r["taibien"].ToString()=="2")?1:0;
							c06=(r["taibien"].ToString()=="3")?1:0;
							c07=(r["taibien"].ToString()!="2" || r["taibien"].ToString()!="3")?1:0;
							dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+1;
							dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
							dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
							dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
						}
						if (r["tuvong"].ToString()!="0")
						{
							c09=(r["tuvong"].ToString()=="1")?1:0;
							c10=(r["tuvong"].ToString()=="2")?1:0;
							dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
							dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
							dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+c10;
						}
					}
				}
				tongcong("ma>1 and ma<10","A",3);
				tongcong("ma>10","B",3);
			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09,sum(a.c10) as c10";
				sql+=" from bieu_04 a,dm_04 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet bieu_04_khoa(string s_makp,string s_tu,string s_tu1,string s_den,string s_table,int i_loaiba,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				s_tu1=s_tu1+" "+m.sGiobaocao;
				s_den=s_den+" "+m.sGiobaocao;
			}
			DataRow[] dr;
			int c02,c03,c05,c06,c07,c09,c10,ma,so;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			sql="SELECT a.mapt,b.loaipt,a.tinhhinh,a.taibien,a.tuvong";
			sql+=" FROM PTTT a,DMPTTT b,benhandt c where a.MAPT = b.MAPT and a.maql=c.maql(+)";
			if (i_loaiba==1) sql+=" and (c.loaiba=1 or c.loaiba is null)";
			else sql+=" and c.loaiba<>1";
			if (time) sql+=" and a.ngay between to_date('"+s_tu1+"',"+stime+") and to_date('"+s_den+"',"+stime+")";
			else sql+=" and to_date(a.ngay,"+stime+") between to_date('"+s_tu1+"',"+stime+") and to_date('"+s_den+"',"+stime+")";
			//if (s_makp!="") sql+=" and a.makp in ("+s_makp.Substring(0,s_makp.Length-1)+")";
			if (s_makp!="" )
			{
				string s=s_makp.Replace(",","','");
				sql+=" and a.makp in ('"+s.Substring(0,s.Length-3)+"')";
			}
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				so=(r["mapt"].ToString().Substring(0,1)=="P")?1:10;
				ma=int.Parse(r["loaipt"].ToString())+so;
				dr=ds.Tables[0].Select("ma="+ma);
				if (dr.Length>0)
				{
					c02=(r["tinhhinh"].ToString()=="2")?1:0;
					c03=(r["tinhhinh"].ToString()!="2")?1:0;
					dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+1;
					dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+c02;
					dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+c03;
					if (r["taibien"].ToString()!="0")
					{
						c05=(r["taibien"].ToString()=="2")?1:0;
						c06=(r["taibien"].ToString()=="3")?1:0;
						c07=(r["taibien"].ToString()!="2" || r["taibien"].ToString()!="3")?1:0;
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+1;
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+c05;
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+c06;
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+c07;
					}
					if (r["tuvong"].ToString()!="0")
					{
						c09=(r["tuvong"].ToString()=="1")?1:0;
						c10=(r["tuvong"].ToString()=="2")?1:0;
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+1;
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+c09;
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+c10;
					}
				}
			}
			tongcong("ma>1 and ma<10","A",3);
			tongcong("ma>10","B",3);
			return ds;
		}

		//ThanhCuong 14/03/2012 bieu 19
		public DataSet bieu_19(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			DataRow[] dr;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			if (benhan)
			{	
				sql="SELECT d.STT,";
				sql+="sum(decode(b.loaiba,1,0,1)) as c01,";
				sql+="sum(case when b.loaiba<>1 and c.phai=1 then 1 else 0 end) as c02,";
				sql+="sum(case when b.loaiba<>1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c03,";
				sql+="sum(case when b.loaiba<>1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(decode(b.loaiba,1,1,0)) as c05,";
				sql+="sum(case when b.loaiba=1 and c.phai=1 then 1 else 0 end) as c06,";
				sql+="sum(case when b.loaiba=1 and a.ttlucrv=7 then 1 else 0 end) as c07,";
				sql+="sum(case when b.loaiba=1 and c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c09,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c10,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 and a.ttlucrv=7 then 1 else 0 end) as c11,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c041,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c051,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c15,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 then 1 else 0 end) as c16,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c17,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c18 ";
				sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d";
				sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.MAICD = d.CICD10 and length(trim(d.stt))>0 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.loaiba<>3";
				sql+=" group by d.stt";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
					if (dr.Length>0)
					{
						dr[0]["c01"]=r["c01"].ToString();
						dr[0]["c02"]=r["c02"].ToString();
						dr[0]["c03"]=r["c03"].ToString();
						dr[0]["c04"]=r["c04"].ToString();
						dr[0]["c05"]=r["c05"].ToString();
						dr[0]["c06"]=r["c06"].ToString();
						dr[0]["c07"]=r["c07"].ToString();
						dr[0]["c08"]=r["c08"].ToString();
						dr[0]["c09"]=r["c09"].ToString();
						dr[0]["c10"]=r["c10"].ToString();
						dr[0]["c11"]=r["c11"].ToString();
						dr[0]["c12"]=r["c12"].ToString();
						dr[0]["c041"]=r["c041"].ToString();
						dr[0]["c051"]=r["c051"].ToString();
						dr[0]["c051"]=r["c15"].ToString();
						dr[0]["c16"]=r["c16"].ToString();
						dr[0]["c17"]=r["c17"].ToString();
						dr[0]["c18"]=r["c18"].ToString();
					}
				}
				
				sql="SELECT d.STT,sum(decode(b.loaiba,1,0,1)) as c01,";
				sql+="sum(case when b.loaiba<>1 and c.phai=1 then 1 else 0 end) as c02,";
				sql+="sum(case when b.loaiba<>1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c03,";
				sql+="sum(case when b.loaiba<>1 and a.ttlucrv=7 then 1 else 0 end) as c04 ";
				sql+=" FROM xxx.XUATVIEN a,xxx.BENHANDT b,BTDBN c,ICD10 d";
				sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.MAICD = d.CICD10 and length(trim(d.stt))>0 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.loaiba=3";
				sql+=" group by d.stt";
				DataSet lds=m.get_data_mmyy(sql,s_tu1,s_den,false);
				if(lds!=null && lds.Tables.Count>0)
				{
					foreach(DataRow r in lds.Tables[0].Rows)
					{
						dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
						if (dr.Length>0)
						{
							dr[0]["c01"]=decimal.Parse(dr[0]["c01"].ToString())+decimal.Parse(r["c01"].ToString());
							dr[0]["c02"]=decimal.Parse(dr[0]["c02"].ToString())+decimal.Parse(r["c02"].ToString());
							dr[0]["c03"]=decimal.Parse(dr[0]["c03"].ToString())+decimal.Parse(r["c03"].ToString());
							dr[0]["c04"]=decimal.Parse(dr[0]["c04"].ToString())+decimal.Parse(r["c04"].ToString());
						}
					}
				}
				
				#region nguyennhan
				if (m.bICDNguyennhan)
				{
					sql="SELECT d.STT,sum(decode(b.loaiba,3,0,1)) as c05,";
					sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c09";
					sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d,cdnguyennhan e";
					sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.maql=e.maql and e.MAICD = d.CICD10 and d.stt is not null and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
					sql+="group by d.stt";
					foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
					{
						dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
						if (dr.Length>0)
						{
							dr[0]["c05"]=decimal.Parse(dr[0]["c05"].ToString())+decimal.Parse(r["c05"].ToString());
							dr[0]["c09"]=decimal.Parse(dr[0]["c09"].ToString())+decimal.Parse(r["c09"].ToString());
						}
					}
				}
				#endregion

				tongcong("ma>=2 and ma<=58","C01",4);
				tongcong("ma>=60 and ma<=98","C02",4);
				tongcong("ma>=100 and ma<=103","C03",4);
				tongcong("ma>=105 and ma<=115","C04",4);
				tongcong("ma>=117 and ma<=124","C05",4);
				tongcong("ma>=126 and ma<=135","C06",4);
				tongcong("ma>=137 and ma<=146","C07",4);
				tongcong("ma>=148 and ma<=150","C08",4);
				tongcong("ma>=152 and ma<=173","C09",4);
				tongcong("ma>=175 and ma<=189","C10",4);
				tongcong("ma>=191 and ma<=208","C11",4);
				tongcong("ma>=210 and ma<=211","C12",4);
				tongcong("ma>=213 and ma<=223","C13",4);
				tongcong("ma>=225 and ma<=247","C14",4);
				tongcong("ma>=249 and ma<=259","C15",4);
				tongcong("ma>=261 and ma<=269","C16",4);
				tongcong("ma>=271 and ma<=283","C17",4);
				tongcong("ma>=285 and ma<=288","C18",4);
				tongcong("ma>=290 and ma<=308","C19",4);
				tongcong("ma>=310 and ma<=323","C20",4);
				tongcong("ma>=325 and ma<=333","C21",4);
			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09,sum(a.c10) as c10,sum(a.c11) as c11,sum(a.c12) as c12";
				if (m.Mabv.Substring(0,3)=="701") sql+=",sum(a.c041) as c041,sum(a.c051) as c051,sum(a.c15) as c15,sum(a.c16) as c16,sum(a.c17) as c17,sum(a.c18) as c18";
				sql+=" from bieu_11 a,dm_11 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
						dr[0]["c11"]=Decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c11"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c12"].ToString());
						if (m.Mabv.Substring(0,3)=="701")
						{
							dr[0]["c041"]=Decimal.Parse(dr[0]["c041"].ToString())+Decimal.Parse(r["c041"].ToString());
							dr[0]["c051"]=Decimal.Parse(dr[0]["c051"].ToString())+Decimal.Parse(r["c051"].ToString());
							dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
							dr[0]["c16"]=Decimal.Parse(dr[0]["c16"].ToString())+Decimal.Parse(r["c16"].ToString());
							dr[0]["c17"]=Decimal.Parse(dr[0]["c17"].ToString())+Decimal.Parse(r["c17"].ToString());
							dr[0]["c18"]=Decimal.Parse(dr[0]["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
						}
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10+c11+c12=0");
			ds.AcceptChanges();
			return ds;
		}
		//
		public DataSet bieu_11(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			DataRow[] dr;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			if (benhan)
			{	
				sql="SELECT d.STT,sum(decode(b.loaiba,1,0,1)) as c01,sum(case when b.loaiba<>1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c02,";
				sql+="sum(case when b.loaiba<>1 and a.ttlucrv=7 then 1 else 0 end) as c03,";
				sql+="sum(decode(b.loaiba,1,1,0)) as c04,sum(case when b.loaiba=1 and c.phai=1 then 1 else 0 end) as c041,";
				sql+="sum(case when b.loaiba=1 and a.ttlucrv=7 then 1 else 0 end) as c05,";
				sql+="sum(case when b.loaiba=1 and c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c051,";
				sql+="sum(decode(b.loaiba,1,to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1,0)) as c06,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c07,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c08,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 and a.ttlucrv=7 then 1 else 0 end) as c09,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c11,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c12,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c15,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 then 1 else 0 end) as c16,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c17,";
				sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when b.loaiba<>1 and c.phai=1 then 1 else 0 end) as c19";
				sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d";
				sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.MAICD = d.CICD10 and length(trim(d.stt))>0 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.loaiba<>3";
				sql+=" group by d.stt";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
					if (dr.Length>0)
					{
						dr[0]["c01"]=r["c01"].ToString();
						dr[0]["c02"]=r["c02"].ToString();
						dr[0]["c03"]=r["c03"].ToString();
						dr[0]["c04"]=r["c04"].ToString();
						dr[0]["c05"]=r["c05"].ToString();
						dr[0]["c06"]=r["c06"].ToString();
						dr[0]["c07"]=r["c07"].ToString();
						dr[0]["c08"]=r["c08"].ToString();
						dr[0]["c09"]=r["c09"].ToString();
						dr[0]["c10"]=r["c10"].ToString();
						dr[0]["c11"]=r["c11"].ToString();
						dr[0]["c12"]=r["c12"].ToString();
						if (m.Mabv.Substring(0,3)=="701")
						{
							dr[0]["c041"]=r["c041"].ToString();
							dr[0]["c051"]=r["c051"].ToString();
							dr[0]["c15"]=r["c15"].ToString();
							dr[0]["c16"]=r["c16"].ToString();
							dr[0]["c17"]=r["c17"].ToString();
							dr[0]["c18"]=r["c18"].ToString();
							dr[0]["c19"]=r["c19"].ToString();
						}
					}
				}
				
				sql="SELECT d.STT,sum(decode(b.loaiba,1,0,1)) as c01,sum(case when b.loaiba<>1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c02,";
				sql+="sum(case when b.loaiba<>1 and a.ttlucrv=7 then 1 else 0 end) as c03,sum(case when b.loaiba<>1 and c.phai=1 then 1 else 0 end) as c19";
				sql+=" FROM xxx.XUATVIEN a,xxx.BENHANDT b,BTDBN c,ICD10 d";
				sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.MAICD = d.CICD10 and length(trim(d.stt))>0 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" and b.loaiba=3";
				sql+=" group by d.stt";
				DataSet lds=m.get_data_mmyy(sql,s_tu1,s_den,false);
				if(lds!=null && lds.Tables.Count>0)
				{
					foreach(DataRow r in lds.Tables[0].Rows)
					{
						dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
						if (dr.Length>0)
						{
							dr[0]["c01"]=decimal.Parse(dr[0]["c01"].ToString())+decimal.Parse(r["c01"].ToString());
							dr[0]["c02"]=decimal.Parse(dr[0]["c02"].ToString())+decimal.Parse(r["c02"].ToString());
							dr[0]["c03"]=decimal.Parse(dr[0]["c03"].ToString())+decimal.Parse(r["c03"].ToString());
							dr[0]["c19"]=decimal.Parse(dr[0]["c19"].ToString())+decimal.Parse(r["c19"].ToString());
						}
					}
				}
				
				#region nguyennhan
				if (m.bICDNguyennhan)
				{
					sql="SELECT d.STT,sum(decode(b.loaiba,3,0,1)) as c04,";
					sql+="sum(case when b.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c07";
					sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d,cdnguyennhan e";
					sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.maql=e.maql and e.MAICD = d.CICD10 and d.stt is not null and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
					sql+="group by d.stt";
					foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
					{
						dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
						if (dr.Length>0)
						{
							dr[0]["c04"]=decimal.Parse(dr[0]["c04"].ToString())+decimal.Parse(r["c04"].ToString());
							dr[0]["c07"]=decimal.Parse(dr[0]["c07"].ToString())+decimal.Parse(r["c07"].ToString());
						}
					}
				}
				#endregion

				tongcong("ma>=2 and ma<=58","C01",4);
				tongcong("ma>=60 and ma<=98","C02",4);
				tongcong("ma>=100 and ma<=103","C03",4);
				tongcong("ma>=105 and ma<=115","C04",4);
				tongcong("ma>=117 and ma<=124","C05",4);
				tongcong("ma>=126 and ma<=135","C06",4);
				tongcong("ma>=137 and ma<=146","C07",4);
				tongcong("ma>=148 and ma<=150","C08",4);
				tongcong("ma>=152 and ma<=173","C09",4);
				tongcong("ma>=175 and ma<=189","C10",4);
				tongcong("ma>=191 and ma<=208","C11",4);
				tongcong("ma>=210 and ma<=211","C12",4);
				tongcong("ma>=213 and ma<=223","C13",4);
				tongcong("ma>=225 and ma<=247","C14",4);
				tongcong("ma>=249 and ma<=259","C15",4);
				tongcong("ma>=261 and ma<=269","C16",4);
				tongcong("ma>=271 and ma<=283","C17",4);
				tongcong("ma>=285 and ma<=288","C18",4);
				tongcong("ma>=290 and ma<=308","C19",4);
				tongcong("ma>=310 and ma<=323","C20",4);
				tongcong("ma>=325 and ma<=333","C21",4);
			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09,sum(a.c10) as c10,sum(a.c11) as c11,sum(a.c12) as c12";
				if (m.Mabv.Substring(0,3)=="701") sql+=",sum(a.c041) as c041,sum(a.c051) as c051,sum(a.c15) as c15,sum(a.c16) as c16,sum(a.c17) as c17,sum(a.c18) as c18";
				sql+=" from bieu_11 a,dm_11 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
						dr[0]["c11"]=Decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c11"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c12"].ToString());
						if (m.Mabv.Substring(0,3)=="701")
						{
							dr[0]["c041"]=Decimal.Parse(dr[0]["c041"].ToString())+Decimal.Parse(r["c041"].ToString());
							dr[0]["c051"]=Decimal.Parse(dr[0]["c051"].ToString())+Decimal.Parse(r["c051"].ToString());
							dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
							dr[0]["c16"]=Decimal.Parse(dr[0]["c16"].ToString())+Decimal.Parse(r["c16"].ToString());
							dr[0]["c17"]=Decimal.Parse(dr[0]["c17"].ToString())+Decimal.Parse(r["c17"].ToString());
							dr[0]["c18"]=Decimal.Parse(dr[0]["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
						}
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10+c11+c12=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet bieu_11_khoa(string s_tu,string s_tu1,string s_den,string s_table,string s_makp,bool phatsinh,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				s_tu1=s_tu1+" "+m.sGiobaocao;
				s_den=s_den+" "+m.sGiobaocao;
			}
			DataRow[] dr;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
			sql="SELECT d.STT,sum(decode(e.loaiba,1,0,1)) as c01,sum(case when e.loaiba<>1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c02,";
			sql+="sum(case when e.loaiba<>1 and a.ttlucrk=7 then 1 else 0 end) as c03,";
			sql+="sum(decode(e.loaiba,1,1,0)) as c04,sum(case when e.loaiba=1 and c.phai=1 then 1 else 0 end) as c041,";
			sql+="sum(case when e.loaiba=1 and a.ttlucrk=7 then 1 else 0 end) as c05,";
			sql+="sum(case when e.loaiba=1 and c.phai=1 and a.ttlucrk=7 then 1 else 0 end) as c051,";
			sql+="sum(decode(e.loaiba,1,to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1,0)) as c06,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c07,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c08,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 and a.ttlucrk=7 then 1 else 0 end) as c09,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrk=7 then 1 else 0 end) as c10,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c11,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then to_date(a.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as c12,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c15,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 then 1 else 0 end) as c16,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrk=7 then 1 else 0 end) as c17,";
			sql+="sum(case when e.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and c.phai=1 and a.ttlucrk=7 then 1 else 0 end) as c18";
			sql+=" FROM xuatkhoa a,nhapkhoa b,BTDBN c,ICD10 d,benhandt e";
			sql+=" where a.id = b.id and b.mabn=c.mabn and a.maicd=d.cicd10 ";
			sql+=" and b.maql=e.maql and length(trim(d.stt))>0 ";
			if (time) sql+=" and a.ngay between to_date('"+s_tu1+"',"+stime+") and to_date('"+s_den+"',"+stime+")";
			else sql+=" and to_date(a.ngay,"+stime+") between to_date('"+s_tu1+"',"+stime+") and to_date('"+s_den+"',"+stime+")";
			//if (s_makp!="") sql+=" and b.makp in ("+s_makp.Substring(0,s_makp.Length-1)+")";
			if (s_makp!="" )
			{
				string s=s_makp.Replace(",","','");
				sql+=" and b.makp in ('"+s.Substring(0,s.Length-3)+"')";
			}
			sql+=" group by d.stt";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
				if (dr.Length>0)
				{
					dr[0]["c01"]=r["c01"].ToString();
					dr[0]["c02"]=r["c02"].ToString();
					dr[0]["c03"]=r["c03"].ToString();
					dr[0]["c04"]=r["c04"].ToString();
					dr[0]["c05"]=r["c05"].ToString();
					dr[0]["c06"]=r["c06"].ToString();
					dr[0]["c07"]=r["c07"].ToString();
					dr[0]["c08"]=r["c08"].ToString();
					dr[0]["c09"]=r["c09"].ToString();
					dr[0]["c10"]=r["c10"].ToString();
					dr[0]["c11"]=r["c11"].ToString();
					dr[0]["c12"]=r["c12"].ToString();
					if (m.Mabv.Substring(0,3)=="701")
					{
						dr[0]["c041"]=r["c041"].ToString();
						dr[0]["c051"]=r["c051"].ToString();
						dr[0]["c15"]=r["c15"].ToString();
						dr[0]["c16"]=r["c16"].ToString();
						dr[0]["c17"]=r["c17"].ToString();
						dr[0]["c18"]=r["c18"].ToString();
					}
				}
			}
				#region nguyennhan
			if (m.bICDNguyennhan)
			{
				sql="SELECT d.STT,sum(decode(f.loaiba,3,0,1)) as c04,";
				sql+="sum(case when f.loaiba=1 and to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c07";
				sql+=" FROM xuatkhoa a,nhapkhoa b,BTDBN c,ICD10 d,cdnguyennhan e,benhandt f";
				sql+=" where a.id=b.id and b.MABN=c.MABN and b.maql=e.maql and b.maql=f.maql and e.MAICD = d.CICD10 and d.stt is not null and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				if (s_makp!="") sql+=" and b.makp in ("+s_makp.Substring(0,s_makp.Length-1)+")";
				sql+="group by d.stt";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
					if (dr.Length>0)
					{
						dr[0]["c04"]=decimal.Parse(dr[0]["c04"].ToString())+decimal.Parse(r["c04"].ToString());
						dr[0]["c07"]=decimal.Parse(dr[0]["c07"].ToString())+decimal.Parse(r["c07"].ToString());
					}
				}
			}
				#endregion

			tongcong("ma>=2 and ma<=58","C01",4);
			tongcong("ma>=60 and ma<=98","C02",4);
			tongcong("ma>=100 and ma<=103","C03",4);
			tongcong("ma>=105 and ma<=115","C04",4);
			tongcong("ma>=117 and ma<=124","C05",4);
			tongcong("ma>=126 and ma<=135","C06",4);
			tongcong("ma>=137 and ma<=146","C07",4);
			tongcong("ma>=148 and ma<=150","C08",4);
			tongcong("ma>=152 and ma<=173","C09",4);
			tongcong("ma>=175 and ma<=189","C10",4);
			tongcong("ma>=191 and ma<=208","C11",4);
			tongcong("ma>=210 and ma<=211","C12",4);
			tongcong("ma>=213 and ma<=223","C13",4);
			tongcong("ma>=225 and ma<=247","C14",4);
			tongcong("ma>=249 and ma<=259","C15",4);
			tongcong("ma>=261 and ma<=269","C16",4);
			tongcong("ma>=271 and ma<=283","C17",4);
			tongcong("ma>=285 and ma<=288","C18",4);
			tongcong("ma>=290 and ma<=308","C19",4);
			tongcong("ma>=310 and ma<=323","C20",4);
			tongcong("ma>=325 and ma<=333","C21",4);

			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10+c11+c12=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet kh_bieu_15(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			DataRow[] dr;
			ds=m.get_data("select * from "+s_table+" order by ma"); 
//			if (benhan)
//			{	
//				sql="SELECT d.STT,count(*) as c01,sum(decode(a.ttlucrv,7,1,0)) as c02,";
//				sql+="sum(case when to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c03,";
//				sql+="sum(case when to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c04,";
//				sql+="sum(case when to_char(sysdate,'yyyy')-c.namsinh<15 and a.ttlucrv=7 then 1 else 0 end) as c05,";
//				sql+="sum(case when to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06";
//				sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d";
//				sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.MAICD = d.CICD10 and b.loaiba=1 and length(trim(d.stt))>0 and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
//				sql+="group by d.stt";
//				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
//				{
//					dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
//					if (dr.Length>0)
//					{
//						dr[0]["c01"]=r["c01"].ToString();
//						dr[0]["c02"]=r["c02"].ToString();
//						dr[0]["c03"]=r["c03"].ToString();
//						dr[0]["c04"]=r["c04"].ToString();
//						dr[0]["c05"]=r["c05"].ToString();
//						dr[0]["c06"]=r["c06"].ToString();
//					}
//				}
//				#region nguyennhan
//				if (m.bICDNguyennhan)
//				{
//					sql="SELECT d.STT,count(*) as c01,sum(case when to_char(sysdate,'yyyy')-c.namsinh<15 then 1 else 0 end) as c03";
//					sql+=" FROM XUATVIEN a,BENHANDT b,BTDBN c,ICD10 d,cdnguyennhan e";
//					sql+=" where a.MAQL = b.MAQL and a.MABN = c.MABN and a.maql=e.maql and e.MAICD = d.CICD10 and d.stt is not null and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
//					sql+="group by d.stt";
//					foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
//					{
//						dr=ds.Tables[0].Select("stt='"+r["stt"].ToString()+"'");
//						if (dr.Length>0)
//						{
//							dr[0]["c01"]=decimal.Parse(dr[0]["c01"].ToString())+decimal.Parse(r["c01"].ToString());
//							dr[0]["c03"]=decimal.Parse(dr[0]["c03"].ToString())+decimal.Parse(r["c03"].ToString());
//						}
//					}
//
//				}
//				#endregion
//				tongcong("ma>=2 and ma<=58","C01",4);
//				tongcong("ma>=60 and ma<=98","C02",4);
//				tongcong("ma>=100 and ma<=103","C03",4);
//				tongcong("ma>=105 and ma<=115","C04",4);
//				tongcong("ma>=117 and ma<=124","C05",4);
//				tongcong("ma>=126 and ma<=135","C06",4);
//				tongcong("ma>=137 and ma<=146","C07",4);
//				tongcong("ma>=148 and ma<=150","C08",4);
//				tongcong("ma>=152 and ma<=173","C09",4);
//				tongcong("ma>=175 and ma<=189","C10",4);
//				tongcong("ma>=191 and ma<=208","C11",4);
//				tongcong("ma>=210 and ma<=211","C12",4);
//				tongcong("ma>=213 and ma<=223","C13",4);
//				tongcong("ma>=225 and ma<=247","C14",4);
//				tongcong("ma>=249 and ma<=259","C15",4);
//				tongcong("ma>=261 and ma<=269","C16",4);
//				tongcong("ma>=271 and ma<=283","C17",4);
//				tongcong("ma>=285 and ma<=288","C18",4);
//				tongcong("ma>=290 and ma<=308","C19",4);
//				tongcong("ma>=310 and ma<=323","C20",4);
//				tongcong("ma>=325 and ma<=333","C21",4);
//			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,sum(a.c09) as c09,sum(a.c10) as c10,";
				sql+="sum(a.c11) as c11,sum(a.c12) as c12 from kh_bieu_15 a,dm_11 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
//						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
//						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
//						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
//						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
//						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
//						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c01"]=Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(r["c10"].ToString());
						dr[0]["c11"]=Decimal.Parse(r["c11"].ToString());
						dr[0]["c12"]=Decimal.Parse(r["c12"].ToString());
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10+c011+c12=0");
			ds.AcceptChanges();
			return ds;
		}

		public DataSet kh_bieu_145(string s_tu,string s_tu1,string s_den,string s_table,bool benhan,bool thongke,bool phatsinh)
		{
			DataRow[] dr;
			ds=m.get_data("select ma,stt,ten,0 as c25,0 as c26,c01,c02,c03,c04,c05,c06,c07,c08,c09,c10,c11,c12,0 as c21,0 as c22,0 as c23,0 as c24,c13,c14,c15,c16,c17,c18,c19,c20 from "+s_table+" order by ma"); 
			if (benhan)
			{	
				sql="SELECT e.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,BTDNN_BV d,BTDNN e";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and c.mann=d.mann and d.mannbo=e.mann";
				sql+=" and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by e.stt";
				upd_data(sql);
				//dia diem
				sql="SELECT d.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,DMDIADIEM d";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and b.diadiem=d.ma";
				sql+=" and d.stt<>0 and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by d.stt";
				upd_data(sql);
				//bo phan
				sql="SELECT d.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,DMbophan d";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and b.bophan=d.ma";
				sql+=" and d.stt<>0 and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by d.stt";
				upd_data(sql);
				//nguyen nhan
				sql="SELECT d.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";

				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";

				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";

				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,DMnguyennhan d";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and b.nguyennhan=d.ma";
				sql+=" and d.stt<>0 and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by d.stt";
				upd_data(sql);
				//ngo doc
				sql="SELECT d.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,DMngodoc d";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and b.ngodoc=d.ma";
				sql+=" and d.stt<>0 and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by d.stt";
				upd_data(sql);
				//xu tri
				sql="SELECT d.stt,sum(decode(c.phai,0,1,0)) as c01, sum(case when c.phai=0 and a.ttlucrv=7 then 1 else 0 end) as c02,";
				sql+="sum(decode(c.phai,1,1,0)) as c03, sum(case when c.phai=1 and a.ttlucrv=7 then 1 else 0 end) as c04,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c05, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c06,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 then 1 else 0 end) as c07, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh<=4 and a.ttlucrv=7 then 1 else 0 end) as c08,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c09,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c10,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14 then 1 else 0 end) as c11,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>4 and to_char(sysdate,'yyyy')-c.namsinh<=14  and a.ttlucrv=7 then 1 else 0 end) as c12,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c21,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c22,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19 then 1 else 0 end) as c23,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>14 and to_char(sysdate,'yyyy')-c.namsinh<=19  and a.ttlucrv=7 then 1 else 0 end) as c24,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c13,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c14,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60 then 1 else 0 end) as c15,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>19 and to_char(sysdate,'yyyy')-c.namsinh<=60  and a.ttlucrv=7 then 1 else 0 end) as c16,";
				sql+="sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c17, sum(case when c.phai=0 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c18,";
				sql+="sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 then 1 else 0 end) as c19, sum(case when c.phai=1 and to_char(sysdate,'yyyy')-c.namsinh>60 and a.ttlucrv=7 then 1 else 0 end) as c20";
				sql+=" FROM XUATVIEN a,TAINANTT b,BTDBN c,dmxutri d";
				sql+=" where b.mabn=c.mabn and b.maql=a.maql(+) and b.xutri=d.ma";
				sql+=" and d.stt<>0 and to_date(b.ngay,'dd/mm/yy') between to_date('"+s_tu1+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by d.stt";
				upd_data(sql);
				//
				tongcong("ma>=3 and ma<=9",1,3);
				tongcong("ma>=3 and ma<=9",2,3);
				tongcong("ma>=11 and ma<=17",10,3);
				tongcong("ma>=19 and ma<=22",18,3);
				tongcong("ma>=31 and ma<=36",30,3);
				tongcong("ma in (24,25,26,27,28,29,30,37,38,39)",23,3);
				tongcong("ma>=41 and ma<=48",40,3);
			}
			
			if (thongke)
			{
				sql="select a.ma,sum(a.c01) as c01,sum(a.c02) as c02,sum(a.c03) as c03,sum(a.c04) as c04,";
				sql+="sum(a.c05) as c05,sum(a.c06) as c06,sum(a.c07) as c07,sum(a.c08) as c08,";
				sql+="sum(a.c09) as c09,sum(a.c10) as c10,sum(a.c11) as c11,sum(a.c12) as c12,";
				sql+="sum(a.c13) as c13,sum(a.c14) as c14,sum(a.c15) as c15,sum(a.c16) as c16,";
				sql+="sum(0) as c21,sum(0) as c22,sum(0) as c23,sum(0) as c24,";
				sql+="sum(a.c17) as c17,sum(a.c18) as c18,sum(a.c19) as c19,sum(a.c20) as c20";
				sql+=" from kh_bieu_1451 a,kh_dm_1451 b where a.ma=b.ma ";
				sql+=" and to_date(a.ngay,'dd/mm/yy') between to_date('"+s_tu+"','dd/mm/yy') and to_date('"+s_den+"','dd/mm/yy')";
				sql+=" group by a.ma";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					dr=ds.Tables[0].Select("ma="+int.Parse(r["ma"].ToString()));
					if (dr.Length>0)
					{
						dr[0]["c01"]=Decimal.Parse(dr[0]["c01"].ToString())+Decimal.Parse(r["c01"].ToString());
						dr[0]["c02"]=Decimal.Parse(dr[0]["c02"].ToString())+Decimal.Parse(r["c02"].ToString());
						dr[0]["c03"]=Decimal.Parse(dr[0]["c03"].ToString())+Decimal.Parse(r["c03"].ToString());
						dr[0]["c04"]=Decimal.Parse(dr[0]["c04"].ToString())+Decimal.Parse(r["c04"].ToString());
						dr[0]["c05"]=Decimal.Parse(dr[0]["c05"].ToString())+Decimal.Parse(r["c05"].ToString());
						dr[0]["c06"]=Decimal.Parse(dr[0]["c06"].ToString())+Decimal.Parse(r["c06"].ToString());
						dr[0]["c07"]=Decimal.Parse(dr[0]["c07"].ToString())+Decimal.Parse(r["c07"].ToString());
						dr[0]["c08"]=Decimal.Parse(dr[0]["c08"].ToString())+Decimal.Parse(r["c08"].ToString());
						dr[0]["c09"]=Decimal.Parse(dr[0]["c09"].ToString())+Decimal.Parse(r["c09"].ToString());
						dr[0]["c10"]=Decimal.Parse(dr[0]["c10"].ToString())+Decimal.Parse(r["c10"].ToString());
						dr[0]["c11"]=Decimal.Parse(dr[0]["c11"].ToString())+Decimal.Parse(r["c11"].ToString());
						dr[0]["c12"]=Decimal.Parse(dr[0]["c12"].ToString())+Decimal.Parse(r["c12"].ToString());
						dr[0]["c13"]=Decimal.Parse(dr[0]["c13"].ToString())+Decimal.Parse(r["c13"].ToString());
						dr[0]["c14"]=Decimal.Parse(dr[0]["c14"].ToString())+Decimal.Parse(r["c14"].ToString());
						dr[0]["c15"]=Decimal.Parse(dr[0]["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
						dr[0]["c16"]=Decimal.Parse(dr[0]["c16"].ToString())+Decimal.Parse(r["c16"].ToString());			
						dr[0]["c21"]=Decimal.Parse(dr[0]["c21"].ToString())+Decimal.Parse(r["c21"].ToString());
						dr[0]["c22"]=Decimal.Parse(dr[0]["c22"].ToString())+Decimal.Parse(r["c22"].ToString());
						dr[0]["c23"]=Decimal.Parse(dr[0]["c23"].ToString())+Decimal.Parse(r["c23"].ToString());
						dr[0]["c24"]=Decimal.Parse(dr[0]["c24"].ToString())+Decimal.Parse(r["c24"].ToString());
						dr[0]["c17"]=Decimal.Parse(dr[0]["c17"].ToString())+Decimal.Parse(r["c17"].ToString());
						dr[0]["c18"]=Decimal.Parse(dr[0]["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
						dr[0]["c19"]=Decimal.Parse(dr[0]["c19"].ToString())+Decimal.Parse(r["c19"].ToString());
						dr[0]["c20"]=Decimal.Parse(dr[0]["c20"].ToString())+Decimal.Parse(r["c20"].ToString());
					}
				}
			}
			if (phatsinh) m.delrec(ds.Tables[0],"c01+c02+c03+c04+c05+c06+c07+c08+c09+c10+c11+c12+c13+c14+c15+c16+c17+c18+c19+c20+c21+c22+c23+c24=0");
			ds.AcceptChanges();
			return ds;
		}

		private void upd_data(string sql)
		{
			DataRow[] dr;
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				dr=ds.Tables[0].Select("ma="+int.Parse(r["stt"].ToString()));
				if (dr.Length>0)
				{
					dr[0]["c01"]=r["c01"].ToString();
					dr[0]["c02"]=r["c02"].ToString();
					dr[0]["c03"]=r["c03"].ToString();
					dr[0]["c04"]=r["c04"].ToString();
					dr[0]["c05"]=r["c05"].ToString();
					dr[0]["c06"]=r["c06"].ToString();
					dr[0]["c07"]=r["c07"].ToString();
					dr[0]["c08"]=r["c08"].ToString();
					dr[0]["c09"]=r["c09"].ToString();
					dr[0]["c10"]=r["c10"].ToString();
					dr[0]["c11"]=r["c11"].ToString();
					dr[0]["c12"]=r["c12"].ToString();
					dr[0]["c13"]=r["c13"].ToString();
					dr[0]["c14"]=r["c14"].ToString();
					dr[0]["c15"]=r["c15"].ToString();
					dr[0]["c16"]=r["c16"].ToString();
					dr[0]["c17"]=r["c17"].ToString();
					dr[0]["c18"]=r["c18"].ToString();
					dr[0]["c19"]=r["c19"].ToString();
					dr[0]["c20"]=r["c20"].ToString();
					dr[0]["c21"]=r["c21"].ToString();
					dr[0]["c22"]=r["c22"].ToString();
					dr[0]["c23"]=r["c23"].ToString();
					dr[0]["c24"]=r["c24"].ToString();
				}
			}
		}

		public DataSet upd_ththbn(string tu,string den,string makp,bool time)
		{
			int iadd=m.iNgaydieutri_ngayra_ngayvao_1;
			string stime=(time)?"'dd/mm/yyyy hh24:mi'":"'dd/mm/yyyy'";
			if (time)
			{
				tu=tu+" "+m.sGiobaocao;
				den=den+" "+m.sGiobaocao;
			}
			Int64 songay=m.songay(m.StringToDate(den.Substring(0,10)),m.StringToDate(tu.Substring(0,10)),1);
			dt=new DataTable();
			sql="select * from btdkp_bv ";
			sql+=" where kehoach>0 ";
			dt=m.get_data(sql).Tables[0];
			ds=new DataSet();
			DataRow r1,r2;
			ds=m.get_data("select * from ththbn");
			dc=new DataColumn();
			dc.ColumnName="C15";//ngaydt
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C16";//ngaydt ravien
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C17";//songay
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="viettat";//songay
			dc.DataType=Type.GetType("System.String");
			ds.Tables[0].Columns.Add(dc);
			sql="SELECT c.MAKP,sum(case when (to_date(to_char(a.ngay,"+stime+"),"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C03,";
			sql+="sum(case when a.khoachuyen='01' and to_date(to_char(a.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C04,";
			sql+="sum(case when a.khoachuyen<>'01' and to_date(to_char(a.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C05,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=1 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C06,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=2 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C07,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=3 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C08,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=4 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C09,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=5 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C10,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=6 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C11,";
			sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=7 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C12,";
			sql+="sum(case when b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+") then 1 else 0 end) C13";
			sql+=" FROM NHAPKHOA a,XUATKHOA b,BTDKP_BV c,benhandt d ";
			sql+=" WHERE a.MAKP=c.MAKP and a.maql=d.maql and a.ID=b.ID(+) and d.loaiba=1 and a.maba<20";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			sql+=" GROUP BY c.makp ORDER BY c.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r2=m.getrowbyid(dt,"makp='"+r["makp"].ToString()+"'");
				if (r2!=null)
				{
					r1 = ds.Tables[0].NewRow();
					r1["makp"] = r["makp"].ToString();
					r1["tenkp"] = r2["tenkp"].ToString();
					r1["c02"] = Decimal.Parse(r2["kehoach"].ToString());
					r1["c14"]= Decimal.Parse(r2["thucke"].ToString());
					r1["c03"] = Decimal.Parse(r["c03"].ToString());
					r1["c13"]= Decimal.Parse(r["c13"].ToString());
					r1["c17"]= songay.ToString();
					r1["c04"]=Decimal.Parse(r["c04"].ToString());
					r1["c05"]=Decimal.Parse(r["c05"].ToString());
					r1["c06"]=Decimal.Parse(r["c06"].ToString());
					r1["c07"]=Decimal.Parse(r["c07"].ToString());
					r1["c08"]=Decimal.Parse(r["c08"].ToString());
					r1["c09"]=Decimal.Parse(r["c09"].ToString());
					r1["c10"]=Decimal.Parse(r["c10"].ToString());
					r1["c11"]=Decimal.Parse(r["c11"].ToString());
					r1["c12"]=Decimal.Parse(r["c12"].ToString());
					r1["c13"]=Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())+Decimal.Parse(r["c05"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c10"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					r1["c15"]=0;
					r1["c16"]=0;
					ds.Tables[0].Rows.Add(r1);
				}
			}
			#region ngaydt linh 05/12/2007
			sql="SELECT a.makp,ceil(sum(case when b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+") ";
			sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yyyy') else to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') end-";
			sql+="case when to_date(to_char(a.ngay,"+stime+"),"+stime+")>to_date('"+tu+"',"+stime+")";
			sql+="then to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yyyy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end)))c15 " ;
			sql+="FROM NHAPKHOA a,XUATKHOA b,benhandt c where  a.maql=c.maql and a.ID = b.ID and c.loaiba=1 and a.maba<20 ";
			sql+=" and to_date(b.ngay,"+stime+") between to_date('"+tu+"',"+stime+") and to_date('"+den+"',"+stime+")";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
					r1["c16"]=Decimal.Parse(r1["c16"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
			}//Thuy 10.05.2012
			sql="SELECT a.makp,ceil(sum(case when b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+") ";
			sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yyyy') else to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') end-";
			sql+="case when to_date(to_char(a.ngay,"+stime+"),"+stime+")>to_date('"+tu+"',"+stime+")";
			sql+="then to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yyyy') end+"+
				" (case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end)))c15 " ;
			sql+="FROM NHAPKHOA a,XUATKHOA b,benhandt c where  a.maql=c.maql and a.ID = b.ID and c.loaiba=1 and a.maba<20 ";
			sql+=" and to_date(a.ngay,"+stime+")<=to_date('"+den+"',"+stime+")";
			sql+=" and to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+")";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
			}
			sql="SELECT a.makp,ceil(sum(case when b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+") ";
			sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yyyy') else to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') end-";
			sql+="case when to_date(to_char(a.ngay,"+stime+"),"+stime+")>to_date('"+tu+"',"+stime+")";
			sql+="then to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yyyy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu.Substring(0,10) + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end)))c15 " ;
			sql+="FROM NHAPKHOA a,XUATKHOA b,benhandt c where  a.maql=c.maql and a.ID = b.ID(+) and c.loaiba=1 and a.maba<20 ";
			sql+=" and to_date(to_char(a.ngay,"+stime+"),"+stime+")<=to_date('"+den+"',"+stime+")";
//			sql+=" and b.ngay is null";
			sql+=" and (b.ngay is null or  to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('"+tu+"','dd/mm/yyyy'))";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
			}
			//xuat tam
			//			if (time)
			//			{
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or b.ngayra>to_date('"+den+"','dd/mm/yyyy hh24:mi') ";
			//				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when b.ngayvao>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			else
			//			{
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+den+"','dd/mm/yy') ";
			//				sql+="then to_date('"+den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when to_date(b.ngayvao,"+stime+")>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			sql+="FROM NHAPKHOA a,XUATtam b,benhandt c where  a.maql=c.maql and a.ID = b.ID and c.loaiba=1 and a.maba<20 ";
			//			sql+=" and to_date(b.ngayra,"+stime+") between to_date('"+tu+"',"+stime+") and to_date('"+den+"',"+stime+")";
			//			if (makp!="") sql+=" and a.makp='"+makp+"'";
			//			sql+=" group by a.makp";
			//			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			//			{
			//				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
			//				if (r1!=null)
			//				{
			//					r1["c15"]=Decimal.Parse(r1["c15"].ToString())-Decimal.Parse(r["c15"].ToString());
			//					r1["c16"]=Decimal.Parse(r1["c16"].ToString())-Decimal.Parse(r["c15"].ToString());
			//				}
			//			}
			//			if (time)
			//			{
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or b.ngayra>to_date('"+den+"','dd/mm/yyyy hh24:mi') ";
			//				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when b.ngayvao>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			else
			//			{
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+den+"','dd/mm/yy') ";
			//				sql+="then to_date('"+den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when to_date(b.ngayvao,"+stime+")>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			sql+="FROM NHAPKHOA a,XUATtam b,benhandt c where  a.maql=c.maql and a.ID = b.ID and c.loaiba=1 and a.maba<20 ";
			//			sql+=" and to_date(b.ngayvao,"+stime+")<=to_date('"+den+"',"+stime+")";
			//			sql+=" and to_date(b.ngayra,"+stime+")>to_date('"+den+"',"+stime+")";
			//			if (makp!="") sql+=" and a.makp='"+makp+"'";
			//			sql+=" group by a.makp";
			//			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			//			{
			//				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
			//				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())-Decimal.Parse(r["c15"].ToString());
			//			}
			//			if (time)
			//			{
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or b.ngayra>to_date('"+den+"','dd/mm/yyyy hh24:mi') ";
			//				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when b.ngayvao>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			else
			//			{				
			//				sql="SELECT a.makp,sum(case when b.ngayra is null or to_date(b.ngayra,'dd/mm/yy')>to_date('"+den+"','dd/mm/yy') ";
			//				sql+="then to_date('"+den+"','dd/mm/yy') else to_date(b.ngayra,'dd/mm/yy') end-";
			//				sql+="case when to_date(b.ngayvao,"+stime+")>to_date('"+tu+"',"+stime+")";
			//				sql+="then to_date(b.ngayvao,'dd/mm/yy') else to_date('"+tu+"','dd/mm/yy') end+(case when to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tu + "','dd/mm/yyyy') then decode(a.khoachuyen,'01',"+iadd+",0) else 1 end))c15 " ;//+decode(a.khoachuyen,'01',"+iadd+",0)) c15 ";
			//			}
			//			sql+="FROM NHAPKHOA a,XUATtam b,benhandt c where  a.maql=c.maql and a.ID = b.ID(+) and c.loaiba=1 and a.maba<20 ";
			//			sql+=" and to_date(b.ngayvao,"+stime+")<=to_date('"+den+"',"+stime+")";
			//			sql+=" and b.ngayra is null";
			//			if (makp!="") sql+=" and a.makp='"+makp+"'";
			//			sql+=" group by a.makp";
			//			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			//			{
			//				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
			//				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())-Decimal.Parse(r["c15"].ToString());
			//			}
			#endregion
			foreach(DataRow r in ds.Tables[0].Rows)
			{
				r1=m.getrowbyid(dt,"makp='"+r["makp"].ToString()+"'");
				if (r1!=null) r["viettat"]=r1["viettat"].ToString();
				else r["viettat"]=r1["makp"].ToString();
			}
			DataSet dsr=new DataSet();
			dsr=ds.Copy();
			ds.Clear();
			ds.Merge(dsr.Tables[0].Select("true","viettat"));
			return ds;
		}

		public DataSet upd_ththbn_ngtru(string tu,string den,string makp,int loaiba,bool time)
		{
			//linh 5/12/2007
			//string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			int iadd=m.iNgaydieutri_ngayra_ngayvao_1;
			sql="select * from btdkp_bv where makp<>'01' ";
			if (loaiba==2)
			{
				sql+=" and (maba like '%20%'";
				sql+=" or maba like '%21%'";
				sql+=" or maba like '%22%'";
				sql+=" or maba like '%23%')";
				if (makp!="" )
				{
					string s=makp.Replace(",","','");
					sql+=" and makp='"+makp+"'";
				}
			}
			else if (loaiba==4) sql+=" and makp='"+LibMedi.AccessData.phongluu+"'";
			sql+=" order by loai,makp";
			string _makp="'";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				_makp+=r["makp"].ToString()+"','";
			_makp=_makp.Substring(0,_makp.Length-2);
			string stime=(time)?"'dd/mm/yyyy hh24:mi'":"'dd/mm/yyyy'";
			if (time)
			{
				tu=tu+" "+m.sGiobaocao;
				den=den+" "+m.sGiobaocao;
			}
			Int64 songay=m.songay(m.StringToDate(den.Substring(0,10)),m.StringToDate(tu.Substring(0,10)),1);
			dt=new DataTable();
			dt=m.get_data("select * from btdkp_bv").Tables[0];
			ds=new DataSet();
			DataRow r1,r2;
			ds=m.get_data("select * from ththbn");
			dc=new DataColumn();
			dc.ColumnName="C15";//ngaydt
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C16";//ngaydt ravien
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C17";//songay
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C18";//ngaydtbhyt
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			dc=new DataColumn();
			dc.ColumnName="C19";//ngaydt ra vien bhyt
			dc.DataType=Type.GetType("System.Decimal");
			ds.Tables[0].Columns.Add(dc);
			//linh 5/12/2007
//			if (time)
//			{
				sql="SELECT c.MAKP,sum(case when (to_date(to_char(a.ngay,"+stime+"),"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C03,";
				sql+="sum(case when (a.madoituong=1 and to_date(to_char(a.ngay,"+stime+"),"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C02,";
				sql+="sum(case when to_date(to_char(a.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C04,";
				sql+="sum(case when a.madoituong=1 and to_date(to_char(a.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C05,";
				sql+="Sum(case when b.ngay is null then 0 else case when to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C06,";//Thuy 28.05.2012 b? b.ttlucrv=1 And
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=2 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C07,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=3 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C08,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=4 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C09,";
				sql+="Sum(case when a.madoituong=1 and to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C10,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=6 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C11,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=7 And to_date(to_char(b.ngay,"+stime+"),"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C12,";
				sql+="sum(case when b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+") then 1 else 0 end) C13,";
				sql+="sum(case when (a.madoituong=1) and (b.ngay is null or to_date(to_char(b.ngay,"+stime+"),"+stime+")>to_date('"+den+"',"+stime+")) then 1 else 0 end) C14";
//			}
//			else
//			{
//				sql="SELECT c.MAKP,sum(case when (to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C03,";
//				sql+="sum(case when (a.madoituong=1 and to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C02,";
//				sql+="sum(case when to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C04,";
//				sql+="sum(case when a.madoituong=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C05,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C06,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=2 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C07,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=3 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C08,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=4 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C09,";
//				sql+="Sum(case when a.madoituong=1 and to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C10,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=6 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C11,";
//				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrv=7 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C12,";
//				sql+="sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") then 1 else 0 end) C13,";
//				sql+="sum(case when (a.madoituong=1) and (b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+")) then 1 else 0 end) C14";
//			}
			sql+=" FROM benhandt a,xuatvien b,btdkp_bv c ";
			sql+=" WHERE a.makp=c.makp and a.maql=b.maql(+) and a.loaiba="+loaiba;
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" GROUP BY c.makp ORDER BY c.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r2=m.getrowbyid(dt,"makp='"+r["makp"].ToString()+"'");
				if (r2!=null)
				{
					r1 = ds.Tables[0].NewRow();
					r1["makp"] = r["makp"].ToString();
					r1["tenkp"] = r2["tenkp"].ToString();
					r1["c02"] = Decimal.Parse(r["c02"].ToString());
					r1["c14"]= Decimal.Parse(r["c14"].ToString());
					r1["c03"] = Decimal.Parse(r["c03"].ToString());
					r1["c13"]= Decimal.Parse(r["c13"].ToString());
					r1["c17"]= songay.ToString();
					r1["c04"]=Decimal.Parse(r["c04"].ToString());
					r1["c05"]=Decimal.Parse(r["c05"].ToString());
					r1["c06"]=Decimal.Parse(r["c06"].ToString());
					r1["c07"]=Decimal.Parse(r["c07"].ToString());
					r1["c08"]=Decimal.Parse(r["c08"].ToString());
					r1["c09"]=Decimal.Parse(r["c09"].ToString());
					r1["c10"]=Decimal.Parse(r["c10"].ToString());
					r1["c11"]=Decimal.Parse(r["c11"].ToString());
					r1["c12"]=Decimal.Parse(r["c12"].ToString());
					r1["c13"]=Decimal.Parse(r["c03"].ToString())+Decimal.Parse(r["c04"].ToString())-(Decimal.Parse(r["c06"].ToString())+Decimal.Parse(r["c07"].ToString())+Decimal.Parse(r["c08"].ToString())+Decimal.Parse(r["c09"].ToString())+Decimal.Parse(r["c11"].ToString())+Decimal.Parse(r["c12"].ToString()));
					r1["c14"]=decimal.Parse(r["c02"].ToString())+decimal.Parse(r["c05"].ToString())-decimal.Parse(r["c10"].ToString());
					r1["c15"]=0;
					r1["c16"]=0;
					r1["c18"]=0;
					r1["c19"]=0;
					ds.Tables[0].Rows.Add(r1);
				}
			}
			/*
			#region ngaydt linh 05/12/2007
			sql="select a.makp,sum(";
			//vao<tu and (den<ra or ra is null) => den-tu
			sql+=" case when to_date(to_char(a.ngay,"+stime+"),"+stime+")< to_date('"+tu+"',"+stime+") ";
			sql+=" and (to_date('"+den+"',"+stime+")< to_date(to_char(b.ngay,"+stime+"),"+stime+") or b.ngay is null) then to_date('"+den+"',"+stime+")-to_date('"+tu+"',"+stime+")+1 else";
			//vao<tu and (ra<=den)              => ra - tu
			sql+=" case when to_date(to_char(a.ngay,"+stime+"),"+stime+")< to_date('"+tu+"',"+stime+") ";
			sql+=" and to_date(to_char(b.ngay,"+stime+"),"+stime+")<=to_date('"+den+"',"+stime+") then to_date(to_char(b.ngay,"+stime+"),"+stime+")-to_date('"+tu+"',"+stime+")+decode(b.ttlucrk,5,0,6,0,7,0,1) else ";
			//vao>=tu and (den<ra or ra is null)=> den-vao
			sql+=" case when to_date(to_char(a.ngay,"+stime+"),"+stime+")>= to_date('"+tu+"',"+stime+") ";
			sql+=" and (to_date('"+den+"',"+stime+")< to_date(to_char(b.ngay,"+stime+"),"+stime+") or b.ngay is null) then to_date('"+den+"',"+stime+")-to_date(to_char(a.ngay,"+stime+"),"+stime+")+1 else ";
			//vao>=tu and (ra<=den)             => ra - vao
			sql+=" case when to_date(to_char(b.ngay,"+stime+"),"+stime+")>= to_date('"+tu+"',"+stime+") ";
			sql+=" and to_date(to_char(b.ngay,"+stime+"),"+stime+")<=to_date('"+den+"',"+stime+") then to_date(to_char(b.ngay,"+stime+"),"+stime+")-to_date(to_char(a.ngay,"+stime+"),"+stime+")+decode(b.ttlucrk,5,0,6,0,7,0,1) end end end end) c15";
			sql+=" from nhapkhoa a,xuatkhoa b,benhandt c,btdbn d ";
			sql+=" where a.id=b.id(+) and a.maql=c.maql and a.maba<20 and c.loaiba=1 and a.mabn=d.mabn and a.makp is not null ";
			sql+=" and to_date(to_char(a.ngay,"+stime+"),"+stime+") <=to_date('"+den+"',"+stime+") and to_date(to_char(nvl(b.ngay,sysdate),"+stime+"),"+stime+") >=to_date('"+tu+"',"+stime+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
					r1["c16"]=Decimal.Parse(r1["c16"].ToString())+Decimal.Parse(r["c15"].ToString());
				}
			}
			*/
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.loaiba="+loaiba;
			sql+=" and to_date(b.ngay,"+stime+") between to_date('"+tu+"',"+stime+") and to_date('"+den+"',"+stime+")";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
					r1["c16"]=Decimal.Parse(r1["c16"].ToString())+Decimal.Parse(r["c16"].ToString());
				}
			}
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.loaiba="+loaiba;
			sql+=" and to_date(a.ngay,"+stime+")<=to_date('"+den+"',"+stime+")";
			sql+=" and to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+")";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
			}
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+"+iadd+") c15 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.loaiba="+loaiba;
			sql+=" and to_date(a.ngay,"+stime+")<=to_date('"+den+"',"+stime+")";
			sql+=" and b.ngay is null ";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null) r1["c15"]=Decimal.Parse(r1["c15"].ToString())+Decimal.Parse(r["c15"].ToString());
			}
			//bhyt
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when b.ngay>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end+1) c19 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end+1) c19 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.madoituong=1 and a.loaiba="+loaiba;
			sql+=" and to_date(b.ngay,"+stime+") between to_date('"+tu+"',"+stime+") and to_date('"+den+"',"+stime+")";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c18"]=Decimal.Parse(r1["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
					r1["c19"]=Decimal.Parse(r1["c19"].ToString())+Decimal.Parse(r["c18"].ToString());
				}
			}
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when b.ngay>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end +1) c19 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end +1) c19 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.madoituong=1 and a.loaiba="+loaiba;
			sql+=" and to_date(a.ngay,"+stime+")<=to_date('"+den+"',"+stime+")";
			sql+=" and to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+")";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c18"]=Decimal.Parse(r1["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
					//r1["c19"]=Decimal.Parse(r1["c19"].ToString())+Decimal.Parse(r["c19"].ToString());
				}
			}
			if (time)
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or b.ngay>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when a.ngay>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when b.ngay>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end +1) c19 ";
			}
			else
			{
				sql="SELECT a.makp,sum(case when b.ngay is null or to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") ";
				sql+="then to_date('"+den.Substring(0,10)+"','dd/mm/yy') else to_date(b.ngay,'dd/mm/yy') end-";
				sql+="case when to_date(a.ngay,"+stime+")>to_date('"+tu+"',"+stime+")";
				sql+="then to_date(a.ngay,'dd/mm/yy') else to_date('"+tu.Substring(0,10)+"','dd/mm/yy') end+1) c18,";
				sql+="sum(case when b.ngay is null then 0 else case when to_date(b.ngay,"+stime+")>to_date('"+den+"',"+stime+") then 0 else to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy') end end +1) c19 ";
			}
			sql+="FROM benhandt a,xuatvien b where  a.maql=b.maql(+) and a.madoituong=1 and a.loaiba="+loaiba;
			sql+=" and to_date(a.ngay,"+stime+")<=to_date('"+den+"',"+stime+")";
			sql+=" and b.ngay is null ";
			if (_makp!="") sql+=" and a.makp in ("+_makp+")";
			sql+=" group by a.makp";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"makp='"+r["makp"].ToString()+"'");
				if (r1!=null)
				{
					r1["c18"]=Decimal.Parse(r1["c18"].ToString())+Decimal.Parse(r["c18"].ToString());
					//r1["c19"]=Decimal.Parse(r1["c19"].ToString())+Decimal.Parse(r["c19"].ToString());
				}
			}
			return ds;
		}

		public DataSet get_bctiepbenh(string m_mmyyyy,string m_makp)
		{
			ds=new DataSet();
			ds.ReadXml("..\\..\\..\\xml\\m_bctiepbenh.xml");
			sql="select to_char(b.ngay,'dd') dd,b.makp,a.namsinh,a.phai,c.mannbo mann,b.bnmoi,b.madoituong,0 tt,";
			sql+="case when substr(b.tuoivao,4,1)<>'0' then substr(b.tuoivao,2,3)||'0' else b.tuoivao end tuoivao";
			sql+=" from btdbn a,xxx.tiepdon b,btdnn_bv c,btdkp_bv d";
			sql+=" where a.mabn=b.mabn and a.mann=c.mann and b.makp=d.makp";
			sql+=" and to_char(b.ngay,'mmyyyy')='"+m_mmyyyy+"'";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" order by to_char(b.ngay,'dd')";
			DataTable dt=m.get_data_nam(m_mmyyyy.Substring(0,2)+m_mmyyyy.Substring(4,2)+"+",sql).Tables[0];
			for(int r=0;r<ds.Tables[0].Rows.Count;r++)
			{
				for(int c=4;c<ds.Tables[0].Columns.Count;c++)
				{
					sql="dd='"+ds.Tables[0].Columns[c].ToString().Substring(1)+"'";
					if (ds.Tables[0].Rows[r]["dk"].ToString()!="")
						sql+=" and "+ds.Tables[0].Rows[r]["dk"].ToString().Trim();
					ds.Tables[0].Rows[r][c]=long.Parse(ds.Tables[0].Rows[r][c].ToString())+dt.Select(sql).Length;
				}
			}
			return ds;
		}


		public DataSet get_slkhambenh(string m_tu,string m_den,string m_makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			DataTable dt=m.get_data("select * from icd10 order by cicd10").Tables[0];
			ds=new DataSet();
			DataSet dsxml=new DataSet();
			dsxml.ReadXml("..\\..\\..\\xml\\m_slkhambenh.xml");
			DataSet dsdk=new DataSet();
			dsdk.ReadXml("..\\..\\..\\xml\\m_dkslkhambenh.xml");
			sql="select b.maicd,a.phai,c.mann,c.mannbo,count(*) so from btdbn a,xxx.benhandt b,btdnn_bv c";
			sql+=" where a.mabn=b.mabn and a.mann=c.mann and b.loaiba=3";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" group by b.maicd,a.phai,c.mann,c.mannbo";
			sql+=" order by b.maicd,a.phai,c.mann,c.mannbo";
			DataRow r1,r2,r3;
			DataRow [] dr1;
			ds=m.get_data_mmyy(sql,m_tu,m_den,false);
			foreach(DataRow r in ds.Tables[0].Rows)
			{
				r1=m.getrowbyid(dsxml.Tables[0],"maicd='"+r["maicd"].ToString()+"'");
				if (r1==null)
				{
					r3=m.getrowbyid(dt,"cicd10='"+r["maicd"].ToString()+"'");
					if (r3!=null)
					{
						r2=dsxml.Tables[0].NewRow();
						r2["maicd"]=r["maicd"].ToString();
						r2["chandoan"]=r3["vviet"].ToString();
						for(int i=1;i<=18;i++) r2["c"+i.ToString().PadLeft(2,'0')]=0;
						//
						foreach(DataRow r4 in dsdk.Tables[0].Rows)
						{
							sql="maicd='"+r["maicd"].ToString()+"'";
							sql+=" and "+r4["dk"].ToString();
							dr1=ds.Tables[0].Select(sql);
							for(int i=0;i<dr1.Length;i++)
								r2[r4["cot"].ToString()]=decimal.Parse(r2[r4["cot"].ToString()].ToString())+decimal.Parse(dr1[i]["so"].ToString());
						}
						//
						dsxml.Tables[0].Rows.Add(r2);
					}
				}
			}
			return dsxml;
		}

		public DataSet get_btpkham(string m_tu,string m_den,string m_makp,bool time,bool taikham)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			DataTable dt=m.get_data("select * from doituong").Tables[0];
			ds=new DataSet();
			DataSet dsxml=new DataSet();
			if(m.Mabv_so==701424)//benh vien mat
				dsxml.ReadXml("..\\..\\..\\xml\\m_btpkham.xml");
			else
			{
				sql=" select 0 as loai, rownum as id, 'trim (maicd) in ('''||cicd10 ||''')' as dk, vviet||' ['||cicd10||']' as noidung, 0 as D01, 0 as D011, 0 as D02, 0 as D03, 0 as D04, 0 as D041, 0 as D05, 0 as D051, 0 as D06, 0 as D061, 0 as D07, 0 as D08, 0 as D09, 0 as D10, 0 as D11, 0 as D12, 0 as D13, 0 as D14, 0 as D15, 0 as D16, 0 as D17 from icd10 order by id_chapter, id_nhom, cicd10 ";
				dsxml=m.get_data(sql);
			}
			if (taikham)
			{
				sql="select b.maicd,b.madoituong,a.phai,a.namsinh,a.matt,count(*) so from btdbn a,xxx.benhandt b";
				sql+=" where a.mabn=b.mabn and b.loaiba=3 and b.mangtr<>0";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by b.maicd,b.madoituong,a.phai,a.namsinh,a.matt";
				ds=m.get_data_mmyy(sql,m_tu,m_den,false);
			}
			else
			{
				sql="select b.maicd,b.madoituong,a.phai,a.namsinh,a.matt,count(*) so from btdbn a,xxx.benhandt b";
				sql+=" where a.mabn=b.mabn and b.loaiba=3";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by b.maicd,b.madoituong,a.phai,a.namsinh,a.matt";
				ds=m.get_data_mmyy(sql,m_tu,m_den,false);
				sql="select '' as maicd,b.madoituong,a.phai,a.namsinh,a.matt,count(*) so from btdbn a,xxx.tiepdon b";
				sql+=" where a.mabn=b.mabn and b.noitiepdon=0 and b.done is null";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by '',b.madoituong,a.phai,a.namsinh,a.matt";
				ds.Merge(m.get_data_mmyy(sql,m_tu,m_den,false));
			}
			int tuoi;
			string matt=m.Mabv.Substring(0,3);
			DataRow r2;
			foreach(DataRow r in dsxml.Tables[0].Select("loai=0"))
			{
				foreach(DataRow r1 in ds.Tables[0].Select(r["DK"].ToString()))
				{
					r2=m.getrowbyid(dt,"madoituong="+int.Parse(r1["madoituong"].ToString()));
					if (r2!=null)
					{
						tuoi=int.Parse(m_tu.Substring(6,4))-((r1["namsinh"].ToString()!="")?int.Parse(r1["namsinh"].ToString()):int.Parse(m_tu.Substring(6,4)));
						if (r1["matt"].ToString()==matt) r["d01"]=decimal.Parse(r["d01"].ToString())+decimal.Parse(r1["so"].ToString());
						else r["d011"]=decimal.Parse(r["d011"].ToString())+decimal.Parse(r1["so"].ToString());
						if (r1["madoituong"].ToString()=="1") r["d02"]=decimal.Parse(r["d02"].ToString())+decimal.Parse(r1["so"].ToString());
						else if (r2["mien"].ToString()=="1")
						{
							if (r1["matt"].ToString()==matt) r["d04"]=decimal.Parse(r["d04"].ToString())+decimal.Parse(r1["so"].ToString());
							else r["d041"]=decimal.Parse(r["d041"].ToString())+decimal.Parse(r1["so"].ToString());
							if (tuoi<6)
							{
								if (r1["matt"].ToString()==matt) r["d08"]=decimal.Parse(r["d08"].ToString())+decimal.Parse(r1["so"].ToString());
								else r["d09"]=decimal.Parse(r["d09"].ToString())+decimal.Parse(r1["so"].ToString());
							}
						}
						else r["d03"]=decimal.Parse(r["d03"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<15)
						{
							if (r1["matt"].ToString()==matt) r["d05"]=decimal.Parse(r["d05"].ToString())+decimal.Parse(r1["so"].ToString());
							else r["d051"]=decimal.Parse(r["d051"].ToString())+decimal.Parse(r1["so"].ToString());						 
						}
						if (tuoi<6)
						{
							if (r1["matt"].ToString()==matt) r["d06"]=decimal.Parse(r["d06"].ToString())+decimal.Parse(r1["so"].ToString());
							else r["d061"]=decimal.Parse(r["d061"].ToString())+decimal.Parse(r1["so"].ToString());
						}
						if (r1["phai"].ToString()=="1") r["d07"]=decimal.Parse(r["d07"].ToString())+decimal.Parse(r1["so"].ToString());
					}
				}
			}
			return dsxml;
		}

        public DataSet get_btpkham_nguoi(string m_tu,string m_den,string m_makp,bool time,bool taikham)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			int nam=int.Parse(m_tu.Substring(6,4));
			string matt=m.Mabv.Substring(0,3);
			ds=new DataSet();
			if (taikham)
			{
				sql="select a.mabn,sum(1) as c01,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt='"+matt+"' then 1 else 0 end) as c02,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt<>'"+matt+"' then 1 else 0 end) as c03,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt='"+matt+"' then 1 else 0 end) as c04,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt<>'"+matt+"' then 1 else 0 end) as c05,";
				sql+="sum(case when a.phai=1 then 1 else 0 end) as c06 ";
				sql+="from btdbn a,xxx.benhandt b";
				sql+=" where a.mabn=b.mabn and b.loaiba=3 and b.mangtr<>0";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by a.mabn";
				ds=m.get_data_mmyy(sql,m_tu,m_den,false);
			}
			else
			{
				sql="select a.mabn,sum(1) as c01,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt='"+matt+"' then 1 else 0 end) as c02,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt<>'"+matt+"' then 1 else 0 end) as c03,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt='"+matt+"' then 1 else 0 end) as c04,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt<>'"+matt+"' then 1 else 0 end) as c05,";
				sql+="sum(case when a.phai=1 then 1 else 0 end) as c06 ";
				sql+="from btdbn a,xxx.benhandt b";
				sql+=" where a.mabn=b.mabn and b.loaiba=3";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by a.mabn";
				ds=m.get_data_mmyy(sql,m_tu,m_den,false);
				sql=" select a.mabn,sum(1) as c01,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt='"+matt+"' then 1 else 0 end) as c02,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<15 and a.matt<>'"+matt+"' then 1 else 0 end) as c03,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt='"+matt+"' then 1 else 0 end) as c04,";
				sql+="sum(case when "+nam+"-to_number(a.namsinh)<6 and a.matt<>'"+matt+"' then 1 else 0 end) as c05,";
				sql+="sum(case when a.phai=1 then 1 else 0 end) as c06 ";
				sql+="from btdbn a,xxx.tiepdon b";
				sql+=" where a.mabn=b.mabn and b.noitiepdon=0 and b.done is null";
				if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
				if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
				sql+=" group by a.mabn";
				ds.Merge(m.get_data_mmyy(sql,m_tu,m_den,false));

			}
			DataSet ret=new DataSet();
			ret=ds.Copy();
			ret.Clear();
			DataRow r2;
			string _ma="";
			foreach(DataRow r in ds.Tables[0].Select("true","mabn"))
			{
				//r1=m.getrowbyid(ret.Tables[0],"mabn='"+r["mabn"].ToString()+"'");
				//if (r1==null)
				if (_ma=="" || _ma!=r["mabn"].ToString())
				{
					r2=ret.Tables[0].NewRow();
					r2["mabn"]=r["mabn"].ToString();
					r2["c01"]=r["c01"].ToString();
					r2["c02"]=r["c02"].ToString();
					r2["c03"]=r["c03"].ToString();
					r2["c04"]=r["c04"].ToString();
					r2["c05"]=r["c05"].ToString();
					r2["c06"]=r["c06"].ToString();
					ret.Tables[0].Rows.Add(r2);
					_ma=r["mabn"].ToString();
				}
			}
			return ret;
		}

		public DataSet get_btngtru(string m_tu,string m_den,string m_makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			ds=new DataSet();
			DataSet dsxml=new DataSet();
			if(m.Mabv_so==701424) dsxml.ReadXml("..\\..\\..\\xml\\m_btpkham.xml");//bv mat
			else
			{
				sql=" select 0 as loai, rownum as id, 'trim (maicd) in ('''||cicd10 ||''')' as dk, vviet||' ['||cicd10||']' as noidung, 0 as D01, 0 as D011, 0 as D02, 0 as D03, 0 as D04, 0 as D041, 0 as D05, 0 as D051, 0 as D06, 0 as D061, 0 as D07, 0 as D08, 0 as D09, 0 as D10, 0 as D11, 0 as D12, 0 as D13, 0 as D14, 0 as D15, 0 as D16, 0 as D17 from icd10 order by id_chapter, id_nhom, cicd10 ";
				dsxml=m.get_data(sql);
			}
			sql="select b.maicd,a.namsinh,b.loaiba,nvl(c.ttlucrv,1) as ttlucrv,count(*) as so from ";
			sql+=" btdbn a,benhandt b,xuatvien c";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+)";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=2 and b.mangtr<>0";
			sql+=" group by b.maicd,a.namsinh,b.loaiba,nvl(c.ttlucrv,1)";
			ds=m.get_data(sql);
			sql="select b.maicd,a.namsinh,b.loaiba,nvl(c.ttlucrv,1) as ttlucrv,count(*) as so from ";
			sql+=" btdbn a,xxx.benhandt b,xxx.xuatvien c";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+)";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=3 and b.mangtr<>0";
			sql+=" group by b.maicd,a.namsinh,b.loaiba,nvl(c.ttlucrv,1)";
			ds.Merge(m.get_data_mmyy(sql,m_tu,m_den,false));
			int tuoi;
			foreach(DataRow r in dsxml.Tables[0].Rows)
			{
				foreach(DataRow r1 in ds.Tables[0].Select(r["DK"].ToString()))
				{
					tuoi=int.Parse(m_tu.Substring(6,4))-int.Parse(r1["namsinh"].ToString());
					if (r1["loaiba"].ToString()=="2")
					{
						r["d01"]=decimal.Parse(r["d01"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<15) r["d02"]=decimal.Parse(r["d02"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<7) r["d03"]=decimal.Parse(r["d03"].ToString())+decimal.Parse(r1["so"].ToString());
					}
					if (r1["loaiba"].ToString()=="3")
					{
						r["d04"]=decimal.Parse(r["d04"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<15) r["d05"]=decimal.Parse(r["d05"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<7) r["d06"]=decimal.Parse(r["d06"].ToString())+decimal.Parse(r1["so"].ToString());
					}
					//else if (r1["loaiba"].ToString()=="2")
					//{
						r["d07"]=decimal.Parse(r["d07"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<15) r["d08"]=decimal.Parse(r["d08"].ToString())+decimal.Parse(r1["so"].ToString());
						if (tuoi<7) r["d09"]=decimal.Parse(r["d09"].ToString())+decimal.Parse(r1["so"].ToString());
					//}
				}
				
			}
			return dsxml;
		}

		public DataSet get_thngtru(string m_tu,string m_den,string m_makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'",mabv=m.Mabv.Substring(0,3);
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			DataRow r1,r2;
			DataRow [] dr;
			int nam=int.Parse(m_den.Substring(6,4));
			sql="select e.tenkp,";
			sql+="sum(case when b.loaiba=2 then 1 else 0 end) as d01,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d02,";
			sql+="sum(case when b.loaiba=2 and a.matt<>'"+mabv+"' then 1 else 0 end) as d03,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d04,";
			sql+="sum(case when b.loaiba=2 and a.matt='"+mabv+"' then 1 else 0 end) as d05,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d06,";
			sql+="sum(0) as d07,sum(0) as d08,sum(0) as d09,sum(0) as d10,sum(0) as d11,sum(0) as d12,sum(0) as d13,sum(0) as d14,";
			sql+="sum(case when b.loaiba=3 then 1 else 0 end) as d15,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d16,";
			sql+="sum(case when b.loaiba=3 and a.matt<>'"+mabv+"' then 1 else 0 end) as d17,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d18,";
			sql+="sum(case when b.loaiba=3 and a.matt='"+mabv+"' then 1 else 0 end) as d19,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d20,";
			sql+="sum(case when b.loaiba=3 and d.mien=0 and b.madoituong<>1 then 1 else 0 end) as d21,";
			sql+="sum(case when b.loaiba=3 and d.mien=1 and b.madoituong<>1 then 1 else 0 end) as d22,";
			sql+="sum(case when b.loaiba=3 and b.madoituong=1 then 1 else 0 end) as d23,";
			sql+="sum(1) as d24,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d25,";
			sql+="sum(case when a.matt<>'"+mabv+"' then 1 else 0 end) as d26,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d27,";
			sql+="sum(case when a.matt='"+mabv+"' then 1 else 0 end) as d28,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d29 ";
			sql+=" from btdbn a,benhandt b,xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+) and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=2 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			ds=m.get_data(sql);
		
			sql="select e.tenkp,";
			sql+="sum(case when b.loaiba=2 then 1 else 0 end) as d01,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d02,";
			sql+="sum(case when b.loaiba=2 and a.matt<>'"+mabv+"' then 1 else 0 end) as d03,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d04,";
			sql+="sum(case when b.loaiba=2 and a.matt='"+mabv+"' then 1 else 0 end) as d05,";
			sql+="sum(case when b.loaiba=2 and "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d06,";
			sql+="sum(case when b.loaiba=3 then 1 else 0 end) as d15,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d16,";
			sql+="sum(case when b.loaiba=3 and a.matt<>'"+mabv+"' then 1 else 0 end) as d17,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d18,";
			sql+="sum(case when b.loaiba=3 and a.matt='"+mabv+"' then 1 else 0 end) as d19,";
			sql+="sum(case when b.loaiba=3 and "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d20,";
			sql+="sum(case when b.loaiba=3 and d.mien=0 and b.madoituong<>1 then 1 else 0 end) as d21,";
			sql+="sum(case when b.loaiba=3 and d.mien=1 and b.madoituong<>1 then 1 else 0 end) as d22,";
			sql+="sum(case when b.loaiba=3 and b.madoituong=1 then 1 else 0 end) as d23,";
			sql+="sum(1) as d24,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d25,";
			sql+="sum(case when a.matt<>'"+mabv+"' then 1 else 0 end) as d26,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d27,";
			sql+="sum(case when a.matt='"+mabv+"' then 1 else 0 end) as d28,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d29 ";
			sql+=" from btdbn a,xxx.benhandt b,xxx.xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+) and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=3 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			foreach(DataRow r in m.get_data_mmyy(sql,m_tu,m_den,false).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"tenkp='"+r["tenkp"].ToString()+"'");
				if (r1==null)
				{
					r2=ds.Tables[0].NewRow();
					r2["tenkp"]=r["tenkp"].ToString();
					r2["d01"]=r["d01"].ToString();
					r2["d02"]=r["d02"].ToString();
					r2["d03"]=r["d03"].ToString();
					r2["d04"]=r["d04"].ToString();
					r2["d05"]=r["d05"].ToString();
					r2["d06"]=r["d06"].ToString();

					r2["d07"]=0;
					r2["d08"]=0;
					r2["d09"]=0;
					r2["d10"]=0;
					r2["d11"]=0;
					r2["d12"]=0;
					r2["d13"]=0;
					r2["d14"]=0;

					r2["d15"]=r["d15"].ToString();
					r2["d16"]=r["d16"].ToString();
					r2["d17"]=r["d17"].ToString();
					r2["d18"]=r["d18"].ToString();
					r2["d19"]=r["d19"].ToString();
					r2["d20"]=r["d20"].ToString();
					r2["d21"]=r["d21"].ToString();
					r2["d22"]=r["d22"].ToString();
					r2["d23"]=r["d23"].ToString();
					r2["d24"]=r["d24"].ToString();
					r2["d25"]=r["d25"].ToString();
					r2["d26"]=r["d26"].ToString();
					r2["d27"]=r["d27"].ToString();
					r2["d28"]=r["d28"].ToString();
					r2["d29"]=r["d29"].ToString();
					ds.Tables[0].Rows.Add(r2);
				}
				else
				{
					dr=ds.Tables[0].Select("tenkp='"+r["tenkp"].ToString()+"'");
					if (dr.Length>0)
					{
						dr[0]["d01"]=decimal.Parse(dr[0]["d01"].ToString())+decimal.Parse(r["d01"].ToString());
						dr[0]["d02"]=decimal.Parse(dr[0]["d02"].ToString())+decimal.Parse(r["d02"].ToString());
						dr[0]["d03"]=decimal.Parse(dr[0]["d03"].ToString())+decimal.Parse(r["d03"].ToString());
						dr[0]["d04"]=decimal.Parse(dr[0]["d04"].ToString())+decimal.Parse(r["d04"].ToString());
						dr[0]["d05"]=decimal.Parse(dr[0]["d05"].ToString())+decimal.Parse(r["d05"].ToString());
						dr[0]["d06"]=decimal.Parse(dr[0]["d06"].ToString())+decimal.Parse(r["d06"].ToString());
						dr[0]["d15"]=decimal.Parse(dr[0]["d15"].ToString())+decimal.Parse(r["d15"].ToString());
						dr[0]["d16"]=decimal.Parse(dr[0]["d16"].ToString())+decimal.Parse(r["d16"].ToString());
						dr[0]["d17"]=decimal.Parse(dr[0]["d17"].ToString())+decimal.Parse(r["d17"].ToString());
						dr[0]["d18"]=decimal.Parse(dr[0]["d18"].ToString())+decimal.Parse(r["d18"].ToString());
						dr[0]["d19"]=decimal.Parse(dr[0]["d19"].ToString())+decimal.Parse(r["d19"].ToString());
						dr[0]["d20"]=decimal.Parse(dr[0]["d20"].ToString())+decimal.Parse(r["d20"].ToString());
						dr[0]["d21"]=decimal.Parse(dr[0]["d21"].ToString())+decimal.Parse(r["d21"].ToString());
						dr[0]["d22"]=decimal.Parse(dr[0]["d22"].ToString())+decimal.Parse(r["d22"].ToString());
						dr[0]["d23"]=decimal.Parse(dr[0]["d23"].ToString())+decimal.Parse(r["d23"].ToString());
						dr[0]["d24"]=decimal.Parse(dr[0]["d24"].ToString())+decimal.Parse(r["d24"].ToString());
						dr[0]["d25"]=decimal.Parse(dr[0]["d25"].ToString())+decimal.Parse(r["d25"].ToString());
						dr[0]["d26"]=decimal.Parse(dr[0]["d26"].ToString())+decimal.Parse(r["d26"].ToString());
						dr[0]["d27"]=decimal.Parse(dr[0]["d27"].ToString())+decimal.Parse(r["d27"].ToString());
						dr[0]["d28"]=decimal.Parse(dr[0]["d28"].ToString())+decimal.Parse(r["d28"].ToString());
						dr[0]["d29"]=decimal.Parse(dr[0]["d29"].ToString())+decimal.Parse(r["d29"].ToString());
					}
				}
			}
			sql="select e.tenkp,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null then 1 else 0 end) as d07,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d08,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d09,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and "+nam+"-to_char(a.namsinh)<7 then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d10,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and a.matt<>'"+mabv+"' then 1 else 0 end) as d11,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and "+nam+"-to_char(a.namsinh)<7 and a.matt<>'"+mabv+"' then 1 else 0 end) as d12,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and a.matt='"+mabv+"' then 1 else 0 end) as d13,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null and "+nam+"-to_char(a.namsinh)<7 and a.matt='"+mabv+"' then 1 else 0 end) as d14";
			sql+=" from btdbn a,benhandt b,xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and c.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(c.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=2 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				dr=ds.Tables[0].Select("tenkp='"+r["tenkp"].ToString()+"'");
				if (dr.Length>0)
				{
					dr[0]["d07"]=decimal.Parse(dr[0]["d07"].ToString())+decimal.Parse(r["d07"].ToString());
					dr[0]["d08"]=decimal.Parse(dr[0]["d08"].ToString())+decimal.Parse(r["d08"].ToString());
					dr[0]["d09"]=decimal.Parse(dr[0]["d09"].ToString())+decimal.Parse(r["d09"].ToString());
					dr[0]["d10"]=decimal.Parse(dr[0]["d10"].ToString())+decimal.Parse(r["d10"].ToString());
					dr[0]["d11"]=decimal.Parse(dr[0]["d11"].ToString())+decimal.Parse(r["d11"].ToString());
					dr[0]["d12"]=decimal.Parse(dr[0]["d12"].ToString())+decimal.Parse(r["d12"].ToString());
					dr[0]["d13"]=decimal.Parse(dr[0]["d13"].ToString())+decimal.Parse(r["d13"].ToString());
					dr[0]["d14"]=decimal.Parse(dr[0]["d14"].ToString())+decimal.Parse(r["d14"].ToString());
				}
			}
			return ds;
		}

		public DataSet get_ckngtru(string m_tu,string m_den,string m_makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			DataRow [] dr;
			DataRow r1,r2;
			DataSet ds=new DataSet();
			sql="select e.tenkp,";
			sql+="sum(1) as d01,";
			sql+="sum(case when b.loaiba=2 then 1 else 0 end) as d02,";
			sql+="sum(0)as d03";
			sql+=" from btdbn a,benhandt b,xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+) and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=2 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			ds=m.get_data(sql);
			sql="select e.tenkp,";
			sql+="sum(1) as d01,";
			sql+="sum(case when b.loaiba=2 then 1 else 0 end) as d02,";
			sql+="sum(0)as d03";
			sql+=" from btdbn a,xxx.benhandt b,xxx.xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql(+) and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and b.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(b.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=3 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			foreach(DataRow r in m.get_data_mmyy(sql,m_tu,m_den,false).Tables[0].Rows)
			{
				r1=m.getrowbyid(ds.Tables[0],"tenkp='"+r["tenkp"].ToString()+"'");
				if (r1==null)
				{
					r2=ds.Tables[0].NewRow();
					r2["tenkp"]=r["tenkp"].ToString();
					r2["d01"]=r["d01"].ToString();
					r2["d02"]=r["d02"].ToString();
					r2["d03"]=r["d03"].ToString();
					ds.Tables[0].Rows.Add(r2);
				}
				else
				{
					dr=ds.Tables[0].Select("tenkp='"+r["tenkp"].ToString()+"'");
					if (dr.Length>0)
					{
						dr[0]["d01"]=decimal.Parse(dr[0]["d01"].ToString())+decimal.Parse(r["d01"].ToString());
						dr[0]["d02"]=decimal.Parse(dr[0]["d02"].ToString())+decimal.Parse(r["d02"].ToString());
						dr[0]["d03"]=decimal.Parse(dr[0]["d03"].ToString())+decimal.Parse(r["d03"].ToString());
					}
				}
			}
			//ngay dt
			sql="select e.tenkp,";
			sql+="sum(case when b.loaiba=2 and c.ngay is not null then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d03";
			sql+=" from btdbn a,benhandt b,xuatvien c,doituong d,btdkp_bv e";
			sql+=" where a.mabn=b.mabn and b.maql=c.maql and b.madoituong=d.madoituong and b.makp=e.makp";
			if (time) sql+=" and c.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(c.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and b.makp in ("+m_makp+")";
			sql+=" and b.loaiba=2 and b.mangtr<>0";
			sql+=" group by e.tenkp ";
			sql+=" order by e.tenkp ";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				dr=ds.Tables[0].Select("tenkp='"+r["tenkp"].ToString()+"'");
				if (dr.Length>0)
					dr[0]["d03"]=decimal.Parse(dr[0]["d03"].ToString())+decimal.Parse(r["d03"].ToString());
			}
			return ds;
		}

		public DataSet get_bcdieutritb(string m_tu,string m_den,string m_makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'",matt=m.Mabv.Substring(0,3);
			int nam=int.Parse(m_den.Substring(6,4));
			if (time)
			{
				m_tu=m_tu+" "+m.sGiobaocao;
				m_den=m_den+" "+m.sGiobaocao;
			}
			DataSet ds=new DataSet();
			DataSet dsxml=new DataSet();
			dsxml.ReadXml("..\\..\\..\\xml\\m_btpkham.xml");
			sql="select c.maicd,";
			sql+="sum(1) as d01,";
			sql+="sum(to_date(c.ngay,'dd/mm/yy')-to_date(d.ngay,'dd/mm/yy')+decode(d.khoachuyen,'01',1,0)) as d02,";
			sql+="sum(case when a.phai=1 then 1 else 0 end) as d03,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)>60 then 1 else 0 end) as d04,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)>60 and a.phai=1 then 1 else 0 end) as d05,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<16 then 1 else 0 end) as d06,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d07,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<5 then 1 else 0 end) as d08,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<16 then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d09,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<7 then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d10,";
			sql+="sum(case when "+nam+"-to_char(a.namsinh)<5 then to_date(c.ngay,'dd/mm/yy')-to_date(b.ngay,'dd/mm/yy')+1 else 0 end) as d11,";
			sql+="sum(case when a.matt<>'"+matt+"' then 1 else 0 end) as d12,";
			sql+="sum(case when a.matt<>'"+matt+"' and "+nam+"-to_char(a.namsinh)<16 then 1 else 0 end) as d13,";
			sql+="sum(case when a.matt<>'"+matt+"' and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d14,";
			sql+="sum(case when a.matt='"+matt+"' then 1 else 0 end) as d15,";
			sql+="sum(case when a.matt='"+matt+"' and "+nam+"-to_char(a.namsinh)<16 then 1 else 0 end) as d16,";
			sql+="sum(case when a.matt='"+matt+"' and "+nam+"-to_char(a.namsinh)<7 then 1 else 0 end) as d17";
			sql+=" from btdbn a,benhandt b,xuatkhoa c,nhapkhoa d";
			sql+=" where a.mabn=b.mabn and b.maql=d.maql and d.id=c.id and b.loaiba=1 ";//and c.ttlucrk<>5";
			if (time) sql+=" and c.ngay between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			else sql+=" and to_date(c.ngay,"+stime+") between to_date('"+m_tu+"',"+stime+") and to_date('"+m_den+"',"+stime+")";
			if (m_makp!="") sql+=" and d.makp='"+m_makp+"'";
			//sql+=" and c.maicd='H26.2'";
			sql+=" group by c.maicd";
			ds=m.get_data(sql);				
			decimal s1=0,s2=0,s3=0,s4=0,s5=0,s6=0,s7=0,s8=0,s9=0,s10=0,s11=0,s12=0,s13=0,s14=0,s15=0,s16=0,s17=0;
			decimal z1=0,z2=0,z3=0,z4=0,z5=0,z6=0,z7=0,z8=0,z9=0,z10=0,z11=0,z12=0,z13=0,z14=0,z15=0,z16=0,z17=0;
			foreach(DataRow r1 in ds.Tables[0].Rows)
			{
				z1+=decimal.Parse(r1["d01"].ToString());
				z2+=decimal.Parse(r1["d02"].ToString());
				z3+=decimal.Parse(r1["d03"].ToString());
				z4+=decimal.Parse(r1["d04"].ToString());
				z5+=decimal.Parse(r1["d05"].ToString());
				z6+=decimal.Parse(r1["d06"].ToString());
				z7+=decimal.Parse(r1["d07"].ToString());
				z8+=decimal.Parse(r1["d08"].ToString());
				z9+=decimal.Parse(r1["d09"].ToString());
				z10+=decimal.Parse(r1["d10"].ToString());
				z11+=decimal.Parse(r1["d11"].ToString());
				z12+=decimal.Parse(r1["d12"].ToString());
				z13+=decimal.Parse(r1["d13"].ToString());
				z14+=decimal.Parse(r1["d14"].ToString());
				z15+=decimal.Parse(r1["d15"].ToString());
				z16+=decimal.Parse(r1["d16"].ToString());
				z17+=decimal.Parse(r1["d17"].ToString());
			}
			foreach(DataRow r in dsxml.Tables[0].Select("id<>27","id"))
			{
				foreach(DataRow r1 in ds.Tables[0].Select(r["DK"].ToString()))
				{
					s1+=decimal.Parse(r1["d01"].ToString());
					s2+=decimal.Parse(r1["d02"].ToString());
					s3+=decimal.Parse(r1["d03"].ToString());
					s4+=decimal.Parse(r1["d04"].ToString());
					s5+=decimal.Parse(r1["d05"].ToString());
					s6+=decimal.Parse(r1["d06"].ToString());
					s7+=decimal.Parse(r1["d07"].ToString());
					s8+=decimal.Parse(r1["d08"].ToString());
					s9+=decimal.Parse(r1["d09"].ToString());
					s10+=decimal.Parse(r1["d10"].ToString());
					s11+=decimal.Parse(r1["d11"].ToString());
					s12+=decimal.Parse(r1["d12"].ToString());
					s13+=decimal.Parse(r1["d13"].ToString());
					s14+=decimal.Parse(r1["d14"].ToString());
					s15+=decimal.Parse(r1["d15"].ToString());
					s16+=decimal.Parse(r1["d16"].ToString());
					s17+=decimal.Parse(r1["d17"].ToString());
					r["d01"]=decimal.Parse(r["d01"].ToString())+decimal.Parse(r1["d01"].ToString());
					r["d02"]=decimal.Parse(r["d02"].ToString())+decimal.Parse(r1["d02"].ToString());
					r["d03"]=decimal.Parse(r["d03"].ToString())+decimal.Parse(r1["d03"].ToString());
					r["d04"]=decimal.Parse(r["d04"].ToString())+decimal.Parse(r1["d04"].ToString());
					r["d05"]=decimal.Parse(r["d05"].ToString())+decimal.Parse(r1["d05"].ToString());
					r["d06"]=decimal.Parse(r["d06"].ToString())+decimal.Parse(r1["d06"].ToString());
					r["d07"]=decimal.Parse(r["d07"].ToString())+decimal.Parse(r1["d07"].ToString());
					r["d08"]=decimal.Parse(r["d08"].ToString())+decimal.Parse(r1["d08"].ToString());
					r["d09"]=decimal.Parse(r["d09"].ToString())+decimal.Parse(r1["d09"].ToString());
					r["d10"]=decimal.Parse(r["d10"].ToString())+decimal.Parse(r1["d10"].ToString());
					r["d11"]=decimal.Parse(r["d11"].ToString())+decimal.Parse(r1["d11"].ToString());
					r["d12"]=decimal.Parse(r["d12"].ToString())+decimal.Parse(r1["d12"].ToString());
					r["d13"]=decimal.Parse(r["d13"].ToString())+decimal.Parse(r1["d13"].ToString());
					r["d14"]=decimal.Parse(r["d14"].ToString())+decimal.Parse(r1["d14"].ToString());
					r["d15"]=decimal.Parse(r["d15"].ToString())+decimal.Parse(r1["d15"].ToString());
					r["d16"]=decimal.Parse(r["d16"].ToString())+decimal.Parse(r1["d16"].ToString());
					r["d17"]=decimal.Parse(r["d17"].ToString())+decimal.Parse(r1["d17"].ToString());
				}
			}
			foreach(DataRow r in dsxml.Tables[0].Select("id=27","id"))
			{
				r["d01"]=z1-s1;
				r["d02"]=z2-s2;
				r["d03"]=z3-s3;
				r["d04"]=z4-s4;
				r["d05"]=z5-s5;
				r["d06"]=z6-s6;
				r["d07"]=z7-s7;
				r["d08"]=z8-s8;
				r["d09"]=z9-s9;
				r["d10"]=z10-s10;
				r["d11"]=z11-s11;
				r["d12"]=z12-s12;
				r["d13"]=z13-s13;
				r["d14"]=z14-s14;
				r["d15"]=z15-s15;
				r["d16"]=z16-s16;
				r["d17"]=z17-s17;
			}
			return dsxml;
		}

		public DataSet get_bcdieutri(DataSet ds,string tu,string den,string makp,bool time)
		{
			string stime=(time)?"'dd/mm/yy hh24:mi'":"'dd/mm/yy'";
			if (time)
			{
				tu=tu+" "+m.sGiobaocao;
				den=den+" "+m.sGiobaocao;
			}
			foreach(DataRow r in ds.Tables[0].Rows)
			{
				r["c01"]=0;r["c02"]=0;r["c03"]=0;r["c04"]=0;
			}
			int namsinh=int.Parse(tu.Substring(6,4));
			decimal c01=0,c02=0,c03=0,c04=0;
			long sogiuong=0,songay=m.songay(m.StringToDate(den.Substring(0,10)),m.StringToDate(tu.Substring(0,10)),1);
			sql="select kehoach from btdkp_bv";
			if (makp!="") sql+=" where makp='"+makp+"'";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows) sogiuong+=long.Parse(r["kehoach"].ToString());
			if (time)
			{
				sql="SELECT sum(case when (a.ngay<to_date('"+tu+"',"+stime+") and (b.ngay is null or b.ngay>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C03,";
				sql+="sum(case when (e.phai=1) and (a.ngay<to_date('"+tu+"',"+stime+") and (b.ngay is null or b.ngay>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C031,";
				sql+="sum(case when ("+namsinh+"-to_char(e.namsinh)<15) and (a.ngay<to_date('"+tu+"',"+stime+") and (b.ngay is null or b.ngay>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C032,";
				sql+="sum(case when ("+namsinh+"-to_char(e.namsinh)<7) and (a.ngay<to_date('"+tu+"',"+stime+") and (b.ngay is null or b.ngay>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C033,";
				sql+="sum(case when a.khoachuyen='01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C04,";
				sql+="sum(case when e.phai=1 and a.khoachuyen='01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C041,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and a.khoachuyen='01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C042,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and a.khoachuyen='01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C043,";
				sql+="sum(case when a.khoachuyen<>'01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C05,";
				sql+="sum(case when e.phai=1 and a.khoachuyen<>'01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C051,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and a.khoachuyen<>'01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C052,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and a.khoachuyen<>'01' and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C053,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk in (1,2,3,4) And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C06,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk in (1,2,3,4) And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C061,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk in (1,2,3,4) And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C062,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk in (1,2,3,4) And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C063,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=5 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C10,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=5 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C101,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=5 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C102,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=5 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C103,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=6 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C11,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=6 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C111,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=6 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C112,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=6 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C113,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=7 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C12,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=7 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C121,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=7 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C122,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=7 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C123,";
				sql+="sum(case when d.nhantu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C14,";
				sql+="sum(case when e.phai=1 and d.nhantu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C141,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.nhantu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C142,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.nhantu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C143,";
				sql+="sum(case when d.madoituong=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C15,";
				sql+="sum(case when e.phai=1 and d.madoituong=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C151,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.madoituong=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C152,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.madoituong=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C153,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C16,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C161,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C162,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C163,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C17,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C171,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C172,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ketqua=1 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(a.ngay,'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C173,";
				sql+="Sum(case when b.ngay is null then 0 else case when d.madoituong=3 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C18,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and d.madoituong=3 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C181,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and d.madoituong=3 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C182,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and d.madoituong=3 And b.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C183,";
				sql+="sum(case when d.dentu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C19,";
				sql+="sum(case when e.phai=1 and d.dentu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C191,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.dentu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C192,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.dentu=1 and a.ngay Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C193";
			}
			else
			{
				sql="SELECT sum(case when (to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C03,";
				sql+="sum(case when (e.phai=1) and (to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C031,";
				sql+="sum(case when ("+namsinh+"-to_char(e.namsinh)<15) and (to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C032,";
				sql+="sum(case when ("+namsinh+"-to_char(e.namsinh)<7) and (to_date(a.ngay,"+stime+")<to_date('"+tu+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+tu+"',"+stime+"))) then 1 else 0 end) C033,";
				sql+="sum(case when a.khoachuyen='01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C04,";
				sql+="sum(case when e.phai=1 and a.khoachuyen='01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C041,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and a.khoachuyen='01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C042,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and a.khoachuyen='01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C043,";
				sql+="sum(case when a.khoachuyen<>'01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C05,";
				sql+="sum(case when e.phai=1 and a.khoachuyen<>'01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C051,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and a.khoachuyen<>'01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C052,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and a.khoachuyen<>'01' and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C053,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk in (1,2,3,4) And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C06,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk in (1,2,3,4) And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C061,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk in (1,2,3,4) And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C062,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk in (1,2,3,4) And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C063,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=5 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C10,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=5 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C101,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=5 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C102,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=5 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C103,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=6 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C11,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=6 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C111,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=6 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C112,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=6 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C113,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ttlucrk=7 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C12,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ttlucrk=7 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C121,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ttlucrk=7 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C122,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ttlucrk=7 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C123,";
				sql+="sum(case when d.nhantu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C14,";
				sql+="sum(case when e.phai=1 and d.nhantu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C141,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.nhantu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C142,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.nhantu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C143,";
				sql+="sum(case when d.madoituong=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C15,";
				sql+="sum(case when e.phai=1 and d.madoituong=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C151,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.madoituong=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C152,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.madoituong=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C153,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C16,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C161,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C162,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C163,";
				sql+="Sum(case when b.ngay is null then 0 else case when b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(to_date(a.ngay,"+stime+"),'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C17,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(to_date(a.ngay,"+stime+"),'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C171,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(to_date(a.ngay,"+stime+"),'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C172,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and b.ketqua=1 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then to_date(b.ngay,'dd/mm/yy')-to_date(to_date(a.ngay,"+stime+"),'dd/mm/yy')+decode(a.khoachuyen,'01',1,0) else 0 end end) C173,";
				sql+="Sum(case when b.ngay is null then 0 else case when d.madoituong=3 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C18,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=1 and d.madoituong=3 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C181,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and d.madoituong=3 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C182,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and d.madoituong=3 And to_date(b.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end end) C183,";
				sql+="sum(case when d.dentu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C19,";
				sql+="sum(case when e.phai=1 and d.dentu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C191,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and d.dentu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C192,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and d.dentu=1 and to_date(a.ngay,"+stime+") Between to_date('"+tu+"',"+stime+") And to_date('"+den+"',"+stime+") then 1 else 0 end) C193";
			}
			sql+=" FROM NHAPKHOA a,XUATKHOA b,BTDKP_BV c,benhandt d,btdbn e ";
			sql+=" WHERE a.MAKP=c.MAKP and a.maql=d.maql and a.mabn=e.mabn and a.ID=b.ID(+) and d.loaiba=1 and a.maba<20";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
			{
				upd_dieutri(ds,1,decimal.Parse(r["c03"].ToString()),decimal.Parse(r["c031"].ToString()),decimal.Parse(r["c032"].ToString()),decimal.Parse(r["c033"].ToString()));
				upd_dieutri(ds,2,decimal.Parse(r["c04"].ToString()),decimal.Parse(r["c041"].ToString()),decimal.Parse(r["c042"].ToString()),decimal.Parse(r["c043"].ToString()));
				upd_dieutri(ds,3,decimal.Parse(r["c03"].ToString())+decimal.Parse(r["c04"].ToString()),decimal.Parse(r["c031"].ToString())+decimal.Parse(r["c041"].ToString()),decimal.Parse(r["c032"].ToString())+decimal.Parse(r["c042"].ToString()),decimal.Parse(r["c033"].ToString())+decimal.Parse(r["c043"].ToString()));
				upd_dieutri(ds,4,decimal.Parse(r["c03"].ToString())+decimal.Parse(r["c04"].ToString())+decimal.Parse(r["c05"].ToString())-(decimal.Parse(r["c06"].ToString())+decimal.Parse(r["c10"].ToString())+decimal.Parse(r["c11"].ToString())+decimal.Parse(r["c12"].ToString())),decimal.Parse(r["c031"].ToString())+decimal.Parse(r["c041"].ToString())+decimal.Parse(r["c051"].ToString())-(decimal.Parse(r["c061"].ToString())+decimal.Parse(r["c101"].ToString())+decimal.Parse(r["c111"].ToString())+decimal.Parse(r["c121"].ToString())),decimal.Parse(r["c032"].ToString())+decimal.Parse(r["c042"].ToString())+decimal.Parse(r["c052"].ToString())-(decimal.Parse(r["c062"].ToString())+decimal.Parse(r["c102"].ToString())+decimal.Parse(r["c112"].ToString())+decimal.Parse(r["c122"].ToString())),decimal.Parse(r["c033"].ToString())+decimal.Parse(r["c043"].ToString())+decimal.Parse(r["c053"].ToString())-(decimal.Parse(r["c063"].ToString())+decimal.Parse(r["c103"].ToString())+decimal.Parse(r["c113"].ToString())+decimal.Parse(r["c123"].ToString())));
				upd_dieutri(ds,5,decimal.Parse(r["c06"].ToString()),decimal.Parse(r["c061"].ToString()),decimal.Parse(r["c062"].ToString()),decimal.Parse(r["c063"].ToString()));
				upd_dieutri(ds,6,decimal.Parse(r["c11"].ToString()),decimal.Parse(r["c111"].ToString()),decimal.Parse(r["c112"].ToString()),decimal.Parse(r["c113"].ToString()));
				upd_dieutri(ds,7,decimal.Parse(r["c19"].ToString()),decimal.Parse(r["c191"].ToString()),decimal.Parse(r["c192"].ToString()),decimal.Parse(r["c193"].ToString()));
				upd_dieutri(ds,8,decimal.Parse(r["c05"].ToString()),decimal.Parse(r["c051"].ToString()),decimal.Parse(r["c052"].ToString()),decimal.Parse(r["c053"].ToString()));
				upd_dieutri(ds,9,decimal.Parse(r["c10"].ToString()),decimal.Parse(r["c101"].ToString()),decimal.Parse(r["c102"].ToString()),decimal.Parse(r["c103"].ToString()));
				upd_dieutri(ds,10,decimal.Parse(r["c14"].ToString()),decimal.Parse(r["c141"].ToString()),decimal.Parse(r["c142"].ToString()),decimal.Parse(r["c143"].ToString()));
				upd_dieutri(ds,11,decimal.Parse(r["c15"].ToString()),decimal.Parse(r["c151"].ToString()),decimal.Parse(r["c152"].ToString()),decimal.Parse(r["c153"].ToString()));
				upd_dieutri(ds,12,decimal.Parse(r["c12"].ToString()),decimal.Parse(r["c121"].ToString()),decimal.Parse(r["c122"].ToString()),decimal.Parse(r["c123"].ToString()));
				upd_dieutri(ds,15,decimal.Parse(r["c17"].ToString()),decimal.Parse(r["c171"].ToString()),decimal.Parse(r["c172"].ToString()),decimal.Parse(r["c173"].ToString()));
				upd_dieutri(ds,17,decimal.Parse(r["c16"].ToString()),decimal.Parse(r["c161"].ToString()),decimal.Parse(r["c162"].ToString()),decimal.Parse(r["c163"].ToString()));
				upd_dieutri(ds,18,decimal.Parse(r["c18"].ToString()),decimal.Parse(r["c181"].ToString()),decimal.Parse(r["c182"].ToString()),decimal.Parse(r["c183"].ToString()));
			}
			
			long so=m.songay(m.StringToDate(den),m.StringToDate(tu),1);
			string sngay="";
			decimal c1=0,c2=0,c3=0,c4=0;
			for(int i=0;i<so;i++)
			{
				sngay=m.DateToString("dd/MM/yyyy",m.StringToDate(den).AddDays(-1*i));
				sql="SELECT ";
				sql+="sum(case when (to_date(a.ngay,"+stime+")<to_date('"+sngay+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+sngay+"',"+stime+"))) then 1 else 0 end) C1,";
				sql+="sum(case when to_date(a.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end) C2,";
				sql+="Sum(case when b.ngay is null then 0 else case when to_date(b.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end end) C3,";
				sql+="sum(case when e.phai=0 and (to_date(a.ngay,"+stime+")<to_date('"+sngay+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+sngay+"',"+stime+"))) then 1 else 0 end) C11,";
				sql+="sum(case when e.phai=0 and to_date(a.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end) C21,";
				sql+="Sum(case when b.ngay is null then 0 else case when e.phai=0 and to_date(b.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end end) C31,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and (to_date(a.ngay,"+stime+")<to_date('"+sngay+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+sngay+"',"+stime+"))) then 1 else 0 end) C12,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 and to_date(a.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end) C22,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<15 and to_date(b.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end end) C32,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and (to_date(a.ngay,"+stime+")<to_date('"+sngay+"',"+stime+") and (b.ngay is null or to_date(b.ngay,"+stime+")>=to_date('"+sngay+"',"+stime+"))) then 1 else 0 end) C13,";
				sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 and to_date(a.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end) C23,";
				sql+="Sum(case when b.ngay is null then 0 else case when "+namsinh+"-to_char(e.namsinh)<7 and to_date(b.ngay,"+stime+") Between to_date('"+sngay+"',"+stime+") And to_date('"+sngay+"',"+stime+") then 1 else 0 end end) C33";
				sql+=" FROM NHAPKHOA a,XUATKHOA b,BTDKP_BV c,benhandt d,btdbn e ";
				sql+=" WHERE a.mabn=e.mabn and a.MAKP=c.MAKP and a.maql=d.maql and a.ID=b.ID(+) and d.loaiba=1 and a.maba<20";
				if (makp!="") sql+=" and a.makp='"+makp+"'";
				foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				{
					c1=decimal.Parse(r["c1"].ToString())+decimal.Parse(r["c2"].ToString())-decimal.Parse(r["c3"].ToString());
					c1=((c1==0 && (decimal.Parse(r["c2"].ToString())>0 || decimal.Parse(r["c3"].ToString())>0))?1:c1);
					c2=decimal.Parse(r["c11"].ToString())+decimal.Parse(r["c21"].ToString())-decimal.Parse(r["c31"].ToString());
					c2=((c2==0 && (decimal.Parse(r["c21"].ToString())>0 || decimal.Parse(r["c31"].ToString())>0))?1:c2);
					c3=decimal.Parse(r["c12"].ToString())+decimal.Parse(r["c22"].ToString())-decimal.Parse(r["c32"].ToString());
					c3=((c3==0 && (decimal.Parse(r["c22"].ToString())>0 || decimal.Parse(r["c32"].ToString())>0))?1:c3);
					c4=decimal.Parse(r["c13"].ToString())+decimal.Parse(r["c23"].ToString())-decimal.Parse(r["c33"].ToString());
					c4=((c4==0 && (decimal.Parse(r["c23"].ToString())>0 || decimal.Parse(r["c33"].ToString())>0))?1:c4);
					upd_dieutri(ds,14,c1,c2,c3,c4);
				}
			}
			
			sql="SELECT ";
			sql+="sum(1) c15,";
			sql+="sum(case when e.phai=0 then 1 else 0 end) c151,";
			sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<15 then 1 else 0 end) c152,";
			sql+="sum(case when "+namsinh+"-to_char(e.namsinh)<7 then 1 else 0 end) c153";
			sql+=" FROM xuatvien a,cdkemtheo b,BTDKP_BV c,benhandt d,btdbn e ";
			sql+=" WHERE a.maql=b.maql and a.MAKP=c.MAKP and a.maql=d.maql and a.mabn=e.mabn and d.loaiba=1";
			sql+=" and to_date(a.ngay,"+stime+") between to_date('"+tu+"',"+stime+") and to_date('"+den+"',"+stime+")";
			sql+=" and substr(b.maicd,1,3) in ('B20','B21','B22','B23')";
			if (makp!="") sql+=" and a.makp='"+makp+"'";
			foreach(DataRow r in m.get_data(sql).Tables[0].Rows)
				upd_dieutri(ds,20,(r["c15"].ToString()=="")?0:decimal.Parse(r["c15"].ToString()),(r["c151"].ToString()=="")?0:decimal.Parse(r["c151"].ToString()),(r["c152"].ToString()=="")?0:decimal.Parse(r["c152"].ToString()),(r["c153"].ToString()=="")?0:decimal.Parse(r["c153"].ToString()));				

			DataRow r1=m.getrowbyid(ds.Tables[0],"stt=14");
			if (r1!=null)
			{
				c01=decimal.Parse(r1["c01"].ToString());
				c02=decimal.Parse(r1["c02"].ToString());
				c03=decimal.Parse(r1["c03"].ToString());
				c04=decimal.Parse(r1["c04"].ToString());
			}
			foreach(DataRow r in ds.Tables[0].Select("stt=16"))
			{
				r["c01"]=(sogiuong==0)?0:c01/sogiuong;
				r["c02"]=(sogiuong==0)?0:c02/sogiuong;
				r["c03"]=(sogiuong==0)?0:c03/sogiuong;
				r["c04"]=(sogiuong==0)?0:c04/sogiuong;
			}
			return ds;
		}

		private void upd_dieutri(DataSet ds,int stt,decimal c1,decimal c2,decimal c3,decimal c4)
		{
			DataRow r=m.getrowbyid(ds.Tables[0],"stt="+stt);
			if (r!=null)
			{
				r["c01"]=decimal.Parse(r["c01"].ToString())+c1;
				r["c02"]=decimal.Parse(r["c02"].ToString())+c2;
				r["c03"]=decimal.Parse(r["c03"].ToString())+c3;
				r["c04"]=decimal.Parse(r["c04"].ToString())+c4;
			}
		}

	}
}