using System;
using System.Data;
using ConfigConnect;

namespace ThongSoData.Medisoft
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class DoiTuong
	{
		Config _ccf;
		AccessDatabase _cacc;
		/// <summary>
		/// Class thông số về đối tượng trên medisoft.
		/// </summary>
		public DoiTuong()
		{
			_ccf=new Config();
			_cacc=new AccessDatabase(_ccf);
			//
			// TODO: Add constructor logic here
			//
		}
		/// <summary>
		/// Class thông số về đối tượng trên medisoft.
		/// </summary>
		/// <param name="cfg">Class cấu hình connect database.</param>
		public DoiTuong(Config cfg)
		{
			_ccf=cfg;
			_cacc=new AccessDatabase(_ccf);
			//
			// TODO: Add constructor logic here
			//
		}
		DataSet get_data(string sql)
		{
			return _cacc.f_get_data(sql);
		}
		#region cao vũ thêm E341 29/12/2014
		/// <summary>
		/// BHYT tính theo thông tư 1314 chi trả cho bệnh nhân nội trú trái tuyến.
		/// </summary>
		public int piBhyt1314_TraiTuyen_NoiTru
		{
			//E34-1
			get
			{
				try
				{
					DataSet ds=get_data("select ten from thongso where id=50341");
					return int.Parse(ds.Tables[0].Rows[0][0].ToString());
				}
				catch{return 0;}
			}
		}
		#endregion
		#region cao vũ thêm E341 29/12/2014
		/// <summary>
		/// BHYT tính theo thông tư 1314 chi trả cho bệnh nhân ngoại trú trái tuyến.
		/// </summary>
		public int piBhyt1314_TraiTuyen_NgoaiTru
		{
			//E34-2
			get
			{
				try
				{
					DataSet ds=get_data("select ten from thongso where id=50342");
					return int.Parse(ds.Tables[0].Rows[0][0].ToString());
				}
				catch{return 0;}
			}
		}
		#endregion
		#region cao vũ thêm E34 26/12/2014
		/// <summary>
		/// Áp dụng chi phí tính theo thông tư 1314 của bhyt
		/// </summary>
		public bool pbBhytApDungTT1314
		{
			//E34
			get
			{
				try
				{
					DataSet ds=get_data("select ten from thongso where id=5034");
					return int.Parse(ds.Tables[0].Rows[0][0].ToString())==1;
				}
				catch{return false;}
			}
		}
		#endregion
		#region cao vũ thêm F70 20/11/2014
		public bool pbHoanTraThuocKhiBNChuaXuatVien
		{
			//F70
			get
			{
				try
				{
					DataSet ds=get_data("select ten from thongso where id=6070");
					return int.Parse(ds.Tables[0].Rows[0][0].ToString())==1;
				}
				catch{return false;}
			}
		}
		#endregion
		public int piTunguyen
		{
			get
			{
				try
				{
					DataSet ds=get_data("select ten from thongso where id=149");
					return int.Parse(ds.Tables[0].Rows[0][0].ToString());
				}
				catch{return 2;}
			}
		}
	}
}
