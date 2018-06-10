namespace dllbhyt
{
    using System;
    using System.Data;

    public class CThongSoXML
    {
        private string _sTenfileXMLThongSo = @"..\..\..\xml\Option_maubaocao_bhyt.xml";

        private void f_chung_setOptionBHYT(int id, int loai, string ten, string giatri)
        {
            DataSet set = new DataSet();
            try
            {
                try
                {
                    set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                }
                catch
                {
                    set.Tables.Add();
                    set.Tables[0].Columns.Add("id");
                    set.Tables[0].Columns.Add("loai");
                    set.Tables[0].Columns.Add("ten");
                    set.Tables[0].Columns.Add("giatri");
                }
                DataRow[] rowArray = set.Tables[0].Select("id=" + id);
                rowArray[0]["loai"] = loai.ToString();
                rowArray[0]["ten"] = ten;
                rowArray[0]["giatri"] = giatri;
            }
            catch
            {
                DataRow row = set.Tables[0].NewRow();
                row["id"] = id;
                row["loai"] = loai;
                row["ten"] = ten;
                row["giatri"] = giatri;
                set.Tables[0].Rows.Add(row);
            }
            set.WriteXml(this._sTenfileXMLThongSo, XmlWriteMode.WriteSchema);
        }

        public bool pChung_bvK1
        {
            get
            {
                DataSet set = new DataSet();
                try
                {
                    set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                    return (set.Tables[0].Select("id=" + ((idthongso) 0x15).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enC_bvK1.GetHashCode(), 0, "Danh dau benh vien k1.", "0");
                    return false;
                }
            }
        }

        public string pChung_FontChu_excel
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return set.Tables[0].Select("id=" + ((idthongso) 2).GetHashCode())[0]["giatri"].ToString();
                }
                catch
                {
                    string giatri = "Arial";
                    this.f_chung_setOptionBHYT(idthongso.enC_FontChu_excel.GetHashCode(), 0, "Mau excel 41 cot_font chu.", giatri);
                    return giatri;
                }
            }
        }

        public bool pChung_MauExcel21_tachtuyen
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 20).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enC_Ex21_tachtuyen.GetHashCode(), 0x15, "Mau excel 21:tach theo tuyen.", "0");
                    return false;
                }
            }
        }

        public bool pChung_MauExcel41cot_hoten
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 3).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enC_Ex41cot_hoten.GetHashCode(), 0, "Mau excel 41 cot_ho ten benh nhan chu thuong.", "0");
                    return false;
                }
            }
        }

        public bool pChung_MauExcel41cot_SapXepNgay
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 8).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enC_Ex41cot_SapXepNgay.GetHashCode(), 0, "Mau excel 41_sort theo ngay tu nho den lon trong bao cao excel 24cot.", "0");
                    return false;
                }
            }
        }

        public bool pChung_MauExcel41cot_TheBHYTBo5KiTuCuoi
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 10))[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enC_Ex41cot_TheBHYTBo5KiTuCuoi.GetHashCode(), 0, "Mau excel 41_Khong hien thi 5 ki tu cuoi cua the BHYT.", "0");
                    return false;
                }
            }
        }

        public string pNgoaiTru_MauExcel25CT_themdscot
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return set.Tables[0].Select("id=" + ((idthongso) 0x13).GetHashCode())[0]["giatri"].ToString();
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNg_Ex25CT_themdscot.GetHashCode(), 0x19, "Mau excel 25CT th\x00c3\x00aam c\x00e1\x00bb™t.", "");
                    return "";
                }
            }
        }

        public string pNgoaiTru_MauExcel41cot_DanhSachCotHienThi
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return set.Tables[0].Select("id=" + ((idthongso) 11).GetHashCode())[0]["giatri"].ToString();
                }
                catch
                {
                    string giatri = "stt#hoten#namsinh#gioitinh#mathe#ma_dkbd#makhoa#mabenh#ngay_vao#ngay_ra#ngaydtr#t_tongchi#t_xn#t_cdha#t_thuoc#t_mau#t_pttt#t_vtytth#t_vtyttt#t_dvktc#t_ktg#t_kham#t_vchuyen#t_bnct#t_bhtt#t_ngoaids#lydo_vv#benhkhac#noikcb#nam_qt#thang_qt#gt_tu#gt_den#diachi#giamdinh#t_xuattoan#lydo_xt#t_datuyen#t_vuottran#loaikcb#noi_ttoan#tt_tngt#mabn#sobienlai#quyenso";
                    this.f_chung_setOptionBHYT(idthongso.enNg_Ex41cot_DanhSachCot.GetHashCode(), 0x4f, "Mau excel 41_danh sach cot se hien thi.", giatri);
                    return giatri;
                }
            }
        }

        public string pNgoaiTru_MauExcel79HD_themdscot
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return set.Tables[0].Select("id=" + ((idthongso) 0x12).GetHashCode())[0]["giatri"].ToString();
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNo_Ex80HD_themcotngaykham.GetHashCode(), 0x4f, "Mau excel 79HD th\x00c3\x00aam c\x00e1\x00bb™t.", "");
                    return "";
                }
            }
        }

        public bool pNoitru_Mau80HD_LocMakpTheoXuatVien
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 7).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNo_80HD_LocMakpTheoXuatVien.GetHashCode(), 0, "Mau80_loc ma khoa phong theo khoa xuat vien cuoi cung.", "0");
                    return false;
                }
            }
        }

        public string pNoiTru_MauExcel41cot_DanhSachCotHienThi
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return set.Tables[0].Select("id=" + ((idthongso) 9).GetHashCode())[0]["giatri"].ToString();
                }
                catch
                {
                    string giatri = this.pNoiTru_MauExcel41cot_DanhSachCotHienThi_default;
                    this.f_chung_setOptionBHYT(idthongso.enNo_Ex41cot_DanhSachCot.GetHashCode(), 80, "Mau excel 41_danh sach cot se hien thi.", giatri);
                    return giatri;
                }
            }
        }

        public string pNoiTru_MauExcel41cot_DanhSachCotHienThi_default
        {
            get
            {
                return "stt#hoten#namsinh#gioitinh#mathe#ma_dkbd#makhoa#mabenh#ngay_vao#ngay_ra#ngaydtr#t_tongchi#t_xn#t_cdha#t_thuoc#t_mau#t_pttt#t_vtytth#t_vtyttt#t_dvktc#t_ktg#t_kham#t_vchuyen#t_bnct#t_bhtt#t_ngoaids#lydo_vv#benhkhac#noikcb#nam_qt#thang_qt#gt_tu#gt_den#diachi#giamdinh#t_xuattoan#lydo_xt#t_datuyen#t_vuottran#loaikcb#noi_ttoan#tt_tngt#mabn#sobienlai#quyenso";
            }
        }

        public bool pNoitru_MauExcel41cot_groupnhom
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 6).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNo_Ex41cot_groupnhom.GetHashCode(), 0, "Mau excel 41 cot_khong group theo nhom.", "0");
                    return false;
                }
            }
        }

        public bool pNoitru_MauExcel41cot_noikcb
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 1).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNo_Ex41cot_noikcb.GetHashCode(), 80, "Noitru_mau excel 41 cot_cot 'noikcb'= cot 'ma_dkbd.", "0");
                    return false;
                }
            }
        }

        public bool pNoiTru_MauExcel80HD_themcotngaykham
        {
            get
            {
                DataSet set = new DataSet();
                set.ReadXml(this._sTenfileXMLThongSo, XmlReadMode.ReadSchema);
                try
                {
                    return (set.Tables[0].Select("id=" + ((idthongso) 13).GetHashCode())[0]["giatri"].ToString() == "1");
                }
                catch
                {
                    this.f_chung_setOptionBHYT(idthongso.enNo_Ex80HD_themcotngaykham.GetHashCode(), 80, "Mau excel 80HD th\x00c3\x00aam c\x00e1\x00bb™t ng\x00c3\x00a0y kh\x00c3\x00a1m.", "0");
                    return false;
                }
            }
        }

        private enum idthongso
        {
            enC_bvK1 = 0x15,
            enC_Ex21_tachtuyen = 20,
            enC_Ex41cot_hoten = 3,
            enC_Ex41cot_SapXepNgay = 8,
            enC_Ex41cot_TheBHYTBo5KiTuCuoi = 10,
            enC_FontChu_excel = 2,
            enNg_Ex25CT_themdscot = 0x13,
            enNg_Ex41cot_DanhSachCot = 11,
            enNg_Ex79HD_themdscot = 0x12,
            enNo_80HD_LocMakpTheoXuatVien = 7,
            enNo_Ex41cot_DanhSachCot = 9,
            enNo_Ex41cot_groupnhom = 6,
            enNo_Ex41cot_noikcb = 1,
            enNo_Ex80HD_themcotngaykham = 13
        }
    }
}

