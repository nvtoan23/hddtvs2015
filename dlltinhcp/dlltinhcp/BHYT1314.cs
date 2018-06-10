using ConfigConnect;
using System;
using System.Xml;
using System.Data;
namespace dlltinhcp
{


    public class BHYT1314
    {
        private AccessDatabase _acc = new AccessDatabase();
        public bool bBHYT1314ApDung = true;
        public bool pbBHYT_Traituyen_tinh_Tyle_khac = false;
        public decimal pdDinhMucKTC = 36000000M;
        public decimal pdTienMienBHYT = 172500M;
        public int piTyLeTraiTuyen = 50;
        public string psNgayApDungTinhCP = "01/01/2015";

        public bool bDoituong_BHYT(int imadoituong)
        {
            if (imadoituong == 1)
            {
                return true;
            }
            string sql = "select madoituong, doituong from doituong where madoituong=" + imadoituong + " and bhyt>0";
            try
            {
                return (this._acc.f_get_data(sql).Tables[0].Rows.Count > 0);
            }
            catch
            {
                return false;
            }
        }

        public int f_get_MaQuyenLoi(string SoThe)
        {
            string str = "+CC1+TE1+";
            string str2 = "+CK2+HN2+DT2+DK2+XD2+BT2+CB2+KC2+TS2+CK2+HN4+";
            string str3 = "+TC3+HT3+CN3+TC7+HT5+CN6+";
            string str4 = "+DN4+HX4+CH4+NN4+TK4+HC4+XK4+TB4+NO4+CT4+XB4+TN4+CS4+XN4+MS4+HD4+TQ4+TA4+TY4+HG4+LS4+HS4+SV4+GB4+GD4+DN7+HX7+CH7+NN7+TK7+HC7+XK7+TB7+NO7+XB7+TN7+XN7+MS7+HD7+TQ7+TA7+TY7+HG7+LS7+HS7+GD7+XV7+TL7+";
            string str5 = "+CA5+QN5+CY5+CA3+";
            int num = 1;
            string str6 = "+" + SoThe.Substring(0, 3) + "+";
            if (str.IndexOf(str6) > -1)
            {
                return 1;
            }
            if (str2.IndexOf(str6) > -1)
            {
                return 2;
            }
            if (str3.IndexOf(str6) > -1)
            {
                return 3;
            }
            if (str4.IndexOf(str6) > -1)
            {
                return 4;
            }
            if (str5.IndexOf(str6) > -1)
            {
                num = 5;
            }
            return num;
        }

        public int f_get_MaQuyenLoiCu(string SoThe)
        {
            string str = "+CC1+TE1+";
            string str2 = "+CK2+HN4+";
            string str3 = "+TC7+HT5+CN6+";
            string str4 = "+DN7+HX7+CH7+NN7+TK7+HC7+XK7+TB7+NO7+XB7+TN7+XN7+MS7+HD7+TQ7+TA7+TY7+HG7+LS7+HS7+GD7+XV7+TL7+";
            string str5 = "+CA3+";
            int num = -1;
            string str6 = "+" + SoThe.Substring(0, 3) + "+";
            if (str.IndexOf(str6) > -1)
            {
                return 1;
            }
            if (str2.IndexOf(str6) > -1)
            {
                return 2;
            }
            if (str3.IndexOf(str6) > -1)
            {
                return 3;
            }
            if (str4.IndexOf(str6) > -1)
            {
                return 4;
            }
            if (str5.IndexOf(str6) > -1)
            {
                num = 5;
            }
            return num;
        }

        public int f_get_TyLeThe(string SoThe, string NgayRaVien)
        {
            int num2 = 100;
            try
            {
                num2 = int.Parse(NgayRaVien.Substring(8, 2) + NgayRaVien.Substring(3, 2) + NgayRaVien.Substring(0, 2));
            }
            catch
            {
            }
            if ((num2 < 0x24a55) && (this.f_get_MaQuyenLoiCu(SoThe) > -1))
            {
                return 0;
            }
            switch (this.f_get_MaQuyenLoi(SoThe))
            {
                case 1:
                case 2:
                case 5:
                    return 100;

                case 3:
                    return 0x5f;

                case 4:
                    return 80;
            }
            return 100;
        }

        public int f_get_TyLeThe(string SoThe, string NgayRaVien, bool TraiTuyen)
        {
            int num = 100;
            int num2 = 100;
            try
            {
                num2 = int.Parse(NgayRaVien.Substring(8, 2) + NgayRaVien.Substring(3, 2) + NgayRaVien.Substring(0, 2));
            }
            catch
            {
            }
            if ((num2 < 0x24a55) && (this.f_get_MaQuyenLoiCu(SoThe) > -1))
            {
                return 0;
            }
            switch (this.f_get_MaQuyenLoi(SoThe))
            {
                case 1:
                case 2:
                case 5:
                    num = 100;
                    break;

                case 3:
                    num = 0x5f;
                    break;

                case 4:
                    num = 80;
                    break;
            }
            if (TraiTuyen)
            {
            }
            return num;
        }

        private void f_getBHYT_SoTienBHTra(string SoThe, decimal SoTienBHYT, string NgayRaVien, bool TraiTuyen, ref decimal BHYTTra, ref int TyLeChiTra)
        {
            int num = 30;
            bool flag = false;
            int num2 = 100;
            int num3 = 100;
            int num4 = 0x24a55;
            int num5 = int.Parse(NgayRaVien.Substring(8, 2) + NgayRaVien.Substring(3, 2) + NgayRaVien.Substring(0, 2));
            num3 = this.f_get_TyLeThe(SoThe, NgayRaVien);
            if (!flag)
            {
                num5 = 0x24a55;
            }
            if (TraiTuyen)
            {
                num2 = num;
                if (num5 < num4)
                {
                    num3 = 100;
                }
            }
            else
            {
                num2 = 100;
            }
            BHYTTra = (((SoTienBHYT * num2) / 100M) * num3) / 100M;
            TyLeChiTra = num3;
        }

        public void f_set_TinhChiTraKyThuatCao(string SoThe, string NgayRaVien, decimal SoTienDinhMucKTC, ref decimal BHYTTra, ref decimal BNTra)
        {
            if (int.Parse(NgayRaVien.Substring(8, 2) + NgayRaVien.Substring(3, 2) + NgayRaVien.Substring(0, 2)) >= 0x24a55)
            {
                switch (this.f_get_MaQuyenLoi(SoThe))
                {
                    case 1:
                    case 5:
                        BHYTTra += BNTra;
                        BNTra = 0M;
                        return;
                }
                if (BHYTTra > SoTienDinhMucKTC)
                {
                    BNTra += BHYTTra - SoTienDinhMucKTC;
                    BHYTTra = SoTienDinhMucKTC;
                }
            }
        }

        public void f_set_TinhLaiChiPhiTheoTT1314(ref DataSet dsxml, DataTable dtdmvp, string fieldBHYTTra, string fieldBNTra, decimal TongTien)
        {
            this.f_set_TinhLaiChiPhiTheoTT1314(ref dsxml, dtdmvp, fieldBHYTTra, fieldBNTra, TongTien, "");
        }

        public void f_set_TinhLaiChiPhiTheoTT1314(ref DataView dsxml, DataTable dtdmvp, string fieldBHYTTra, string fieldBNTra, decimal TongTien)
        {
            this.f_set_TinhLaiChiPhiTheoTT1314(ref dsxml, dtdmvp, fieldBHYTTra, fieldBNTra, TongTien, "");
        }

        public void f_set_TinhLaiChiPhiTheoTT1314(ref DataSet dsxml, DataTable dtdmvp, string fieldBHYTTra, string fieldBNTra, decimal TongTien, string SoThe)
        {
            dsxml.WriteXml("dxmlin.xml", XmlWriteMode.WriteSchema);
            decimal num = TongTien;
            for (int i = 0; i < dsxml.Tables[0].Rows.Count; i++)
            {
                DataRow row = dsxml.Tables[0].Rows[i];
                if (this.bDoituong_BHYT(int.Parse(row["madoituong"].ToString())))
                {
                    string soThe = "";
                    try
                    {
                        soThe = row["sothe"].ToString();
                    }
                    catch
                    {
                        soThe = SoThe;
                    }
                    int num3 = this.f_get_TyLeThe(soThe, this.psNgayApDungTinhCP);
                    int piTyLeTraiTuyen = 100;
                    decimal bHYTTra = 0M;
                    decimal num6 = 0M;
                    int num7 = 0;
                    try
                    {
                        num7 = int.Parse(row["vat"].ToString());
                    }
                    catch
                    {
                    }
                    decimal num8 = 0M;
                    try
                    {
                        num8 = decimal.Parse(row["dongia"].ToString());
                    }
                    catch
                    {
                    }
                    decimal num9 = 0M;
                    try
                    {
                        num9 = decimal.Parse(row["soluong"].ToString());
                    }
                    catch
                    {
                    }
                    if (row["traituyen"].ToString() != "0")
                    {
                        piTyLeTraiTuyen = this.piTyLeTraiTuyen;
                        if (this.pbBHYT_Traituyen_tinh_Tyle_khac)
                        {
                            try
                            {
                                DataRow[] rowArray = dtdmvp.Select("id=" + row["ma"].ToString());
                                if ((rowArray.Length > 0) && (decimal.Parse(rowArray[0]["bhyt"].ToString()) > decimal.Parse(rowArray[0]["bhyt_tt"].ToString())))
                                {
                                    piTyLeTraiTuyen = int.Parse(rowArray[0]["bhyt_tt"].ToString());
                                }
                            }
                            catch
                            {
                            }
                        }
                        if (decimal.Parse(row["sotien"].ToString()) == 0M)
                        {
                            num9 = 0M;
                        }
                        bHYTTra = (((((num3 * piTyLeTraiTuyen) * num9) * num8) * (1 + (num7 / 100))) / 100M) / 100M;
                        try
                        {
                            if (row["kythuat"].ToString() == "0")
                            {
                                decimal bNTra = 0M;
                                this.f_set_TinhChiTraKyThuatCao(soThe, this.psNgayApDungTinhCP, this.pdDinhMucKTC, ref bHYTTra, ref bNTra);
                            }
                        }
                        catch
                        {
                            bHYTTra = (((num3 * num9) * num8) * (1 + (num7 / 100))) / 100M;
                        }
                    }
                    else
                    {
                        try
                        {
                            if (row["kythuat"].ToString() == "0")
                            {
                                decimal num11 = 0M;
                                this.f_set_TinhChiTraKyThuatCao(soThe, this.psNgayApDungTinhCP, this.pdDinhMucKTC, ref bHYTTra, ref num11);
                            }
                            else if (decimal.Parse(row["sotien"].ToString()) == 0M)
                            {
                                bHYTTra = 0M;
                            }
                            else
                            {
                                bHYTTra = (((num3 * num9) * num8) * (1 + (num7 / 100))) / 100M;
                            }
                        }
                        catch
                        {
                            bHYTTra = (((num3 * num9) * num8) * (1 + (num7 / 100))) / 100M;
                        }
                        if (num <= this.pdTienMienBHYT)
                        {
                            bHYTTra = num9 * num8;
                        }
                    }
                    num6 = (num8 * num9) - bHYTTra;
                    row[fieldBNTra] = num6;
                    row[fieldBHYTTra] = bHYTTra;
                    try
                    {
                        row["tltraituyen"] = piTyLeTraiTuyen;
                    }
                    catch
                    {
                    }
                    try
                    {
                        row["tlchitra"] = num3;
                    }
                    catch
                    {
                    }
                }
            }
            dsxml.AcceptChanges();
        }

        public void f_set_TinhLaiChiPhiTheoTT1314(ref DataView dsxml, DataTable dtdmvp, string fieldBHYTTra, string fieldBNTra, decimal TongTien, string SoThe)
        {
            DataSet set = new DataSet();
            set.Tables.Add(dsxml.Table.Copy());
            this.f_set_TinhLaiChiPhiTheoTT1314(ref set, dtdmvp, fieldBHYTTra, fieldBNTra, TongTien, SoThe);
            dsxml.Table = set.Tables[0].Copy();
        }
    }
}

