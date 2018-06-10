namespace dllbhyt
{
    using dichso;
    using dlltinhcp;
    using Excel;
    using System;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Forms;

    public class LibMaubaocao
    {
        private bool _bQLTraiTuyen;
        private decimal _deTyLeTraiTuyen;
        private DataSet _dsgiavp;
        private int _iGiaBhyt;
        private int _iSoThe15KiTu_ChieuDai;
        private int _iSoThe15KiTu_vitri;
        private LibClass _lib;
        private string[] _s_arrCap;
        private string _s_xmlpath;
        private string _sIDProcessExcelCurrent;
        private string _sSoThe15KiTu_kitu80;
        private string _sSoThe15KiTu_kitu95;
        private string _sViTriTheMoi;
        private string _sViTriTheTrongTinh;
        private CThongSoXML _tsxml;
        private Excel.Range orange;
        private _Worksheet osheet;
        private _Workbook owb;
        private Excel.Application oxl;
        private string tenfile;

        public LibMaubaocao()
        {
            this._s_arrCap = new string[] { "", "\r\n  ", "\r\n    ", "\r\n      ", "\r\n        " };
            this._s_xmlpath = @"..\xml\";
            this.tenfile = "";
            this._sViTriTheMoi = "";
            this._sViTriTheTrongTinh = "";
            this._iGiaBhyt = 0;
            this._sIDProcessExcelCurrent = "";
            this._iSoThe15KiTu_vitri = 2;
            this._iSoThe15KiTu_ChieuDai = 1;
            this._bQLTraiTuyen = false;
            this._sSoThe15KiTu_kitu80 = "";
            this._sSoThe15KiTu_kitu95 = "";
            this._deTyLeTraiTuyen = 0M;
            this._dsgiavp = new DataSet();
            this._tsxml = new CThongSoXML();
            this._lib = new LibClass();
            this.f_LibMaubaocao_load();
        }

        public LibMaubaocao(LibClass libbc)
        {
            this._s_arrCap = new string[] { "", "\r\n  ", "\r\n    ", "\r\n      ", "\r\n        " };
            this._s_xmlpath = @"..\xml\";
            this.tenfile = "";
            this._sViTriTheMoi = "";
            this._sViTriTheTrongTinh = "";
            this._iGiaBhyt = 0;
            this._sIDProcessExcelCurrent = "";
            this._iSoThe15KiTu_vitri = 2;
            this._iSoThe15KiTu_ChieuDai = 1;
            this._bQLTraiTuyen = false;
            this._sSoThe15KiTu_kitu80 = "";
            this._sSoThe15KiTu_kitu95 = "";
            this._deTyLeTraiTuyen = 0M;
            this._dsgiavp = new DataSet();
            this._tsxml = new CThongSoXML();
            this._lib = libbc;
            this.f_LibMaubaocao_load();
        }

        private void exp_excel_1399(DataSet vds, string tungay, string denngay)
        {
            try
            {
                int num = 9;
                int num2 = vds.Tables[0].Rows.Count + 9;
                int num3 = vds.Tables[0].Columns.Count - 2;
                int num4 = num2 + 2;
                this.tenfile = this._lib.Export_Excel(vds, "baocaomau21xml");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int i = 0; i < num; i++)
                {
                    this.osheet.get_Range(this._lib.getIndex(i) + "1", this._lib.getIndex(i) + "1").EntireRow.Insert(Missing.Value);
                }
                string[] strArray = new string[] { "STT", "M\x00e3 số theo danh mục do BYT ban h\x00e0nh", "T\x00ean VTYT theo danh mục do BYT ban h\x00e0nh", "T\x00ean thương mại", "Quy c\x00e1ch", "\x00d0ơn vị t\x00ednh", "Gi\x00e1 mua v\x00e0o (đồng)", "Ngoại tr\x00fa", "Nội tr\x00fa", "Gi\x00e1 thanh to\x00e1n BHYT (đồng)", "Th\x00e0nh tiền (đồng)" };
                int num6 = 7;
                for (int j = 0; j < strArray.Length; j++)
                {
                    this.osheet.Cells[num, j + 1] = strArray[j];
                    this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                    this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num + 1)), this._lib.getIndex(j) + ((num + 1)));
                    this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    if (j != num6)
                    {
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + num, this._lib.getIndex(j) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.WrapText = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        this.osheet.Cells[num - 1, j + 1] = "Số lượng";
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num - 1)), this._lib.getIndex(j + 1) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.Font.Bold = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j++;
                        this.osheet.Cells[num, j + 1] = strArray[j];
                        this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                }
                this.orange = this.osheet.get_Range("A" + ((num - 1)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(6) + ((num + 2)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.NumberFormat = "#,##0.00";
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, 3] = "TỔNG CỘNG";
                this.orange = this.osheet.get_Range("A" + (((vds.Tables[0].Rows.Count + num) + 2)), this._lib.getIndex(6) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.MergeCells = true;
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, vds.Tables[0].Columns.Count] = "=sum(" + string.Concat(new object[] { this._lib.getIndex(vds.Tables[0].Columns.Count - 1), num + 2, ":", this._lib.getIndex(vds.Tables[0].Columns.Count - 1), (vds.Tables[0].Rows.Count + num) + 1 });
                this.oxl.ActiveWindow.DisplayZeros = false;
                this.osheet.Cells[1, 1] = "T\x00ean cơ sở y tế: " + this._lib.Tenbv;
                this.osheet.Cells[2, 1] = "M\x00e3 cơ sở y tế: " + this._lib.Mabv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[4, 3] = "THỐNG K\x00ca VẬT TƯ Y TẾ THANH TO\x00c1N BHYT";
                this.osheet.Cells[6, 3] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "4", this._lib.getIndex(num3 + 1) + "6");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.oxl.Visible = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void exp_excel_1399_20(DataSet vds, string tungay, string denngay)
        {
            try
            {
                int num = 9;
                int num2 = vds.Tables[0].Rows.Count + 9;
                int num3 = vds.Tables[0].Columns.Count - 2;
                int num4 = num2 + 2;
                this.tenfile = this._lib.Export_Excel(vds, "baocaomau21xml");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int i = 0; i < num; i++)
                {
                    this.osheet.get_Range(this._lib.getIndex(i) + "1", this._lib.getIndex(i) + "1").EntireRow.Insert(Missing.Value);
                }
                string[] strArray = new string[] { "STT", "M\x00e3 số theo danh mục BYT", "T\x00ean hoạt chất", "T\x00ean thuốc th\x00e0nh phẩm", "Đường d\x00f9ng, dạng b\x00e0o chế", "H\x00e0m lượng/ Nồng độ", "Số đăng k\x00fd hoặc GPNK", "\x00d0ơn vị t\x00ednh", "Ngoại tr\x00fa", "Nội tr\x00fa", "Đơn gi\x00e1 (đồng)", "Th\x00e0nh tiền (đồng)" };
                int num6 = 8;
                for (int j = 0; j < strArray.Length; j++)
                {
                    this.osheet.Cells[num, j + 1] = strArray[j];
                    this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                    this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num + 1)), this._lib.getIndex(j) + ((num + 1)));
                    this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    if (j != num6)
                    {
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + num, this._lib.getIndex(j) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.WrapText = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        this.osheet.Cells[num - 1, j + 1] = "Số lượng";
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num - 1)), this._lib.getIndex(j + 1) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.Font.Bold = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j++;
                        this.osheet.Cells[num, j + 1] = strArray[j];
                        this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                }
                this.orange = this.osheet.get_Range("A" + ((num - 1)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(6) + ((num + 2)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.NumberFormat = "#,##0.00";
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, 3] = "TỔNG CỘNG";
                this.orange = this.osheet.get_Range("A" + (((vds.Tables[0].Rows.Count + num) + 2)), this._lib.getIndex(6) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.MergeCells = true;
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, vds.Tables[0].Columns.Count] = "=sum(" + string.Concat(new object[] { this._lib.getIndex(vds.Tables[0].Columns.Count - 1), num + 2, ":", this._lib.getIndex(vds.Tables[0].Columns.Count - 1), (vds.Tables[0].Rows.Count + num) + 1 });
                this.oxl.ActiveWindow.DisplayZeros = false;
                this.osheet.Cells[1, 1] = "T\x00ean cơ sở y tế: " + this._lib.Tenbv;
                this.osheet.Cells[2, 1] = "M\x00e3 cơ sở y tế: " + this._lib.Mabv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[4, 3] = "THỐNG K\x00ca THUỐC THANH TO\x00c1N BHYT";
                this.osheet.Cells[6, 3] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "4", this._lib.getIndex(num3 + 1) + "6");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.oxl.Visible = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void exp_excel_1399_21(DataSet vds, string tungay, string denngay)
        {
            try
            {
                int num = 9;
                int num2 = vds.Tables[0].Rows.Count + 9;
                int num3 = vds.Tables[0].Columns.Count - 2;
                int num4 = num2 + 2;
                this.tenfile = this._lib.Export_Excel(vds, "baocaomau21xml");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int i = 0; i < num; i++)
                {
                    this.osheet.get_Range(this._lib.getIndex(i) + "1", this._lib.getIndex(i) + "1").EntireRow.Insert(Missing.Value);
                }
                string[] strArray = new string[] { "STT", "M\x00e3 số theo danh mục BYT", "T\x00ean dịch vụ y tế", "Ngoại tr\x00fa", "Nội tr\x00fa", "Đơn gi\x00e1 (đồng)", "Th\x00e0nh tiền (đồng)" };
                int num6 = 3;
                for (int j = 0; j < strArray.Length; j++)
                {
                    this.osheet.Cells[num, j + 1] = strArray[j];
                    this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                    this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num + 1)), this._lib.getIndex(j) + ((num + 1)));
                    this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    if (j != num6)
                    {
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + num, this._lib.getIndex(j) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.WrapText = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        this.osheet.Cells[num - 1, j + 1] = "Số lượng";
                        this.orange = this.osheet.get_Range(this._lib.getIndex(j) + ((num - 1)), this._lib.getIndex(j + 1) + ((num - 1)));
                        this.orange.MergeCells = true;
                        this.orange.Font.Bold = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        j++;
                        this.osheet.Cells[num, j + 1] = strArray[j];
                        this.osheet.Cells[num + 1, j + 1] = "[" + ((j + 1)) + "]";
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                }
                for (int k = 0; k < vds.Tables[0].Rows.Count; k++)
                {
                    if (vds.Tables[0].Rows[k]["stt"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range(this._lib.getIndex(0) + ((((num + 1) + k) + 1)), this._lib.getIndex(2) + ((((num + 1) + k) + 1)));
                        this.orange.MergeCells = true;
                        this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        this.orange = this.osheet.get_Range(this._lib.getIndex(0) + ((((num + 1) + k) + 1)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + ((((num + 1) + k) + 1)));
                        this.orange.Font.Bold = true;
                        this.orange.Interior.ColorIndex = 15;
                    }
                }
                this.orange = this.osheet.get_Range("A" + ((num - 1)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(6) + ((num + 2)), this._lib.getIndex(vds.Tables[0].Columns.Count - 1) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.NumberFormat = "#,##0.00";
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, 3] = "TỔNG CỘNG";
                this.orange = this.osheet.get_Range("A" + (((vds.Tables[0].Rows.Count + num) + 2)), this._lib.getIndex(6) + (((vds.Tables[0].Rows.Count + num) + 2)));
                this.orange.MergeCells = true;
                this.osheet.Cells[(vds.Tables[0].Rows.Count + num) + 2, vds.Tables[0].Columns.Count] = "=sum(" + string.Concat(new object[] { this._lib.getIndex(vds.Tables[0].Columns.Count - 1), num + 2, ":", this._lib.getIndex(vds.Tables[0].Columns.Count - 1), (vds.Tables[0].Rows.Count + num) + 1 });
                this.oxl.ActiveWindow.DisplayZeros = false;
                this.osheet.Cells[1, 1] = "T\x00ean cơ sở y tế: " + this._lib.Tenbv;
                this.osheet.Cells[2, 1] = "M\x00e3 cơ sở y tế: " + this._lib.Mabv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[4, 3] = "THỐNG K\x00ca THUỐC THANH TO\x00c1N BHYT";
                this.osheet.Cells[6, 3] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "4", this._lib.getIndex(num3 + 1) + "6");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.oxl.Visible = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public string exp_excel_file(DataSet vds, string tenfile)
        {
            try
            {
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Add(Missing.Value);
                this.osheet = (Worksheet) this.owb.Worksheets.get_Item(1);
                this.oxl.ActiveWindow.DisplayGridlines = true;
                this.owb.SaveAs(tenfile, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.owb.Close(true, Missing.Value, Missing.Value);
                this.oxl.Quit();
                this.releaseObject(this.osheet);
                this.releaseObject(this.owb);
                this.releaseObject(this.oxl);
                return tenfile;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                return "";
            }
        }

        private string f_cdata_tag(string str)
        {
            str = str.Replace("&", "").Replace("<", "").Replace(">", "").Replace("'", "").Replace("*", "");
            return ("<![CDATA[" + str + "]]>");
        }

        public void f_Chung_MauExcel_20_xanhpon(DataSet dsxml, string tieudebaocao, string thoigian)
        {
            DataSet set = dsxml.Copy();
            dsxml.Clear();
            string str = "";
            int num = 0;
            decimal num2 = 0M;
            string path = Directory.GetCurrentDirectory() + "//Excel//excelmau20user.xls";
            string str3 = "Arial";
            StringBuilder builder = new StringBuilder();
            StreamWriter writer = new StreamWriter(path, false, Encoding.Unicode);
            string str4 = str + "<table><tr>";
            string str5 = str4 + "<td colspan=3 style=\"font-family:" + str3 + ";align=left\">" + this._lib.Syte + "</td></tr>";
            string str6 = str5 + "<tr><td colspan=3 style=\"font-family:" + str3 + ";align=left\">" + this._lib.Tenbv + "</td></tr>";
            string str7 = str6 + "<tr><th colspan=5 rowspan=2 style=\"font-family:" + str3 + ";align=center;font-size:16pt\">" + tieudebaocao + "</th></tr><tr></tr>";
            str = str7 + "<tr><th colspan=5 style=\"font-family:" + str3 + ";align=center;font-size:10pt;\">" + thoigian + "</th></tr><tr></tr></table>";
            writer.Write(str);
            str = "<table border=1 style=\"font-family:" + str3 + ";font-size:11pt\"><tr><th>STT</th><th>T\x00caN THU?C</th><th>H\x00c0M LU?NG</th><th>S? ĐK/GPNK</th><th>H\x00c3NG S?N XU?T</th><th>NU?C S?N XU?T</th><th>ĐVT</th><th>S? LU?NG</th><th>ĐON GI\x00c1</th><th>TH\x00c0NH TI?N</th><th>BV \x00c1P TH?U</th></tr></table>";
            writer.Write(str);
            str = "<table border=1 style=\"font-family:" + str3 + ";font-size:10pt\">";
            string[] strArray = new string[] { "mabv1 ='" + this._lib.Mabv + "'#TUYẾN 1", "mabv1 <>'" + this._lib.Mabv + "' and mabv1 like '" + this._lib.MABHXH.Substring(0, 2) + "%'#TUYẾN 2", "mabv1 <>'" + this._lib.Mabv + "' and mabv1 not like '" + this._lib.MABHXH.Substring(0, 2) + "%'#TUYẾN 3" };
            for (int i = 0; i < strArray.Length; i++)
            {
                string filterExpression = this._tsxml.pChung_MauExcel21_tachtuyen ? strArray[i].Split(new char[] { '#' })[0] : " 1=1 ";
                string str9 = "";
                string str10 = "";
                string str11 = "";
                decimal num4 = 0M;
                str9 = "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=11>" + strArray[i].Split(new char[] { '#' })[1] + "</td></tr>";
                DataRow[] rowArray = set.Tables[0].Select(filterExpression);
                if (rowArray.Length > 0)
                {
                    for (int j = 0; j < rowArray.Length; j++)
                    {
                        try
                        {
                            str11 = str11 + "<tr style=\"font-family:" + str3 + "\">";
                            str11 = string.Concat(new object[] { str11, "<td>", ++num, "</td>" });
                            str11 = str11 + "<td>" + rowArray[j]["ten"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["HAMLUONG"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["SODK"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["HANGSX"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["NUOCSX"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["DVT"].ToString() + "</td>";
                            str11 = str11 + "<td>" + rowArray[j]["SOLUONG"].ToString() + "</td>";
                            str11 = str11 + "<td>" + decimal.Parse(rowArray[j]["dongia"].ToString()).ToString("###,###,###") + "</td>";
                            str11 = str11 + "<td>" + decimal.Parse(rowArray[j]["sotien"].ToString()).ToString("###,###,###") + "</td>";
                            str11 = str11 + "<td></td>";
                            str11 = str11 + "</tr>";
                            num2 += decimal.Parse(rowArray[j]["sotien"].ToString());
                            num4 += decimal.Parse(rowArray[j]["sotien"].ToString());
                        }
                        catch
                        {
                        }
                    }
                }
                str10 = "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=9>cộng " + strArray[i].Split(new char[] { '#' })[1] + "</td><td >" + num4.ToString("###,###,###") + "</td><td ></td></tr>";
                if (!this._tsxml.pChung_MauExcel21_tachtuyen)
                {
                    str = str + str11;
                    break;
                }
                if (num4 == 0M)
                {
                    str10 = str9 = "";
                }
                str = str + str9 + str11 + str10;
            }
            str = str + "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=9>Tổng cộng</td><td >" + num2.ToString("###,###,###") + "</td><td ></td></tr></table>";
            writer.Write(str);
            writer.Close();
            Process.Start(path);
        }

        public void f_Chung_MauExcel_21_xanhpon(DataSet dsxml, string tieudebaocao, string thoigian)
        {
            DataSet set = dsxml.Copy();
            dsxml.Clear();
            string str = "";
            int num = 0;
            decimal num2 = 0M;
            decimal num3 = 0M;
            string path = Environment.CurrentDirectory + "//Excel//excelmau21user.xls";
            string str3 = "Arial";
            StringBuilder builder = new StringBuilder();
            StreamWriter writer = new StreamWriter(path, false, Encoding.Unicode);
            string[] strArray = new string[] { "A#1#X\x00e9t nghiệm", "B#2#CĐHA,TDCN", "C#4#M\x00e1u", "D#5#PT,TT", "E#6#V\x00e2?t tu y t\x00ea\x00b4", "F#11#Ti\x00ea`n giuo`ng", "G#1,2,4,5,6,11#Nh\x00f3m kh\x00e1c" };
            string str4 = str + "<table><tr>";
            string str5 = str4 + "<td colspan=3 style=\"font-family:" + str3 + ";align=left\">" + this._lib.Syte + "</td></tr>";
            string str6 = str5 + "<tr><td colspan=3 style=\"font-family:" + str3 + ";align=left\">" + this._lib.Tenbv + "</td></tr>";
            string str7 = str6 + "<tr><th colspan=5 rowspan=2 style=\"font-family:" + str3 + ";align=center;font-size:16pt\">" + tieudebaocao + "</th></tr><tr></tr>";
            str = str7 + "<tr><th colspan=5 style=\"font-family:" + str3 + ";align=center;font-size:10pt;\">" + thoigian + "</th></tr><tr></tr></table>";
            writer.Write(str);
            str = "<table border=1 style=\"font-family:" + str3 + ";font-size:11pt\"><tr><th>STT</th><th>T\x00caN D?CH V? K? THU?T</th><th>S? LU?NG</th><th>ĐON GI\x00c1</th><th>TH\x00c0NH TI?N</th></tr><tr><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th></tr></table>";
            writer.Write(str);
            str = "<table border=1 style=\"font-family:" + str3 + ";font-size:10pt\">";
            string[] strArray2 = new string[] { "tuyen=1#TUYẾN 1", "tuyen=2#TUYẾN 2", "tuyen=3#TUYẾN 3" };
            for (int i = 0; i < strArray2.Length; i++)
            {
                string str8 = this._tsxml.pChung_MauExcel21_tachtuyen ? strArray2[i].Split(new char[] { '#' })[0] : " 1=1 ";
                string str9 = "";
                string str10 = "";
                string str11 = "";
                decimal num5 = 0M;
                str9 = "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=5>" + strArray2[i].Split(new char[] { '#' })[1] + "</td></tr>";
                for (int j = 0; j < strArray.Length; j++)
                {
                    string str12 = "";
                    if (j == (strArray.Length - 1))
                    {
                        str12 = " and nhombhyt not in (" + strArray[j].Split(new char[] { '#' })[1] + ")";
                    }
                    else
                    {
                        str12 = " and nhombhyt=" + strArray[j].Split(new char[] { '#' })[1];
                    }
                    DataRow[] rowArray = set.Tables[0].Select(str8 + str12);
                    if (rowArray.Length > 0)
                    {
                        str11 = str11 + "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td >" + strArray[j].Split(new char[] { '#' })[0] + "</td><td colspan=4>" + strArray[j].Split(new char[] { '#' })[2] + "</td></tr>";
                        num3 = 0M;
                        for (int k = 0; k < rowArray.Length; k++)
                        {
                            try
                            {
                                str11 = str11 + "<tr style=\"font-family:" + str3 + "\">";
                                str11 = string.Concat(new object[] { str11, "<td>", ++num, "</td>" });
                                str11 = str11 + "<td>" + rowArray[k]["ten"].ToString() + "</td>";
                                str11 = str11 + "<td>" + rowArray[k]["soluong"].ToString() + "</td>";
                                str11 = str11 + "<td>" + decimal.Parse(rowArray[k]["dongia"].ToString()).ToString("###,###,###") + "</td>";
                                str11 = str11 + "<td>" + decimal.Parse(rowArray[k]["sotien"].ToString()).ToString("###,###,###") + "</td>";
                                str11 = str11 + "</tr>";
                                num2 += decimal.Parse(rowArray[k]["sotien"].ToString());
                                num3 += decimal.Parse(rowArray[k]["sotien"].ToString());
                                num5 += decimal.Parse(rowArray[k]["sotien"].ToString());
                            }
                            catch
                            {
                            }
                        }
                        str11 = str11 + "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=4>cộng " + strArray[j].Split(new char[] { '#' })[2] + "</td><td >" + num3.ToString("###,###,###") + "</td></tr>";
                    }
                }
                str10 = "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=4>cộng " + strArray2[i].Split(new char[] { '#' })[1] + "</td><td >" + num5.ToString("###,###,###") + "</td></tr>";
                if (!this._tsxml.pChung_MauExcel21_tachtuyen)
                {
                    str = str + str11;
                    break;
                }
                if (num5 == 0M)
                {
                    str10 = str9 = "";
                }
                str = str + str9 + str11 + str10;
            }
            str = str + "<tr style=\"font-family:" + str3 + ";background-color:#C0C0C0;font-weight: bold\"><td colspan=4>Tổng cộng</td><td >" + num2.ToString("###,###,###") + "</td></tr></table>";
            writer.Write(str);
            writer.Close();
            Process.Start(path);
        }

        private string f_convert_to_unicode(string strText)
        {
            try
            {
                return Convert.ToBase64String(Encoding.UTF8.GetBytes(strText));
            }
            catch
            {
                return "";
            }
        }

        public void f_CreateMaLK_BHYT(int loaibn, string mmyy)
        {
            try
            {
                if (mmyy == "")
                {
                    mmyy = DateTime.Now.ToString("MMyy");
                }
                DateTime time = new DateTime(int.Parse(mmyy.Substring(2)) + 0x7d0, int.Parse(mmyy.Substring(0, 2)), 1);
                DateTime time2 = time;
                if (loaibn == 1)
                {
                    mmyy = "";
                }
                else
                {
                    time = time.AddMonths(1);
                    time2 = time.AddMonths(-2);
                }
                string user = "";
                string sql = "";
                decimal num = 0M;
                while (time2 <= time)
                {
                    if (this._lib.bMmyy(time2.ToString("MMyy")))
                    {
                        if (loaibn == 1)
                        {
                            user = this._lib.user;
                        }
                        else
                        {
                            user = this._lib.user + time2.ToString("MMyy");
                        }
                        sql = "create table " + user + ".bc_bhytmalk(maql numeric(22),malk numeric(15) default 0,constraint pk_" + user + "_bc_bhytmalk primary key(maql))";
                        this._lib.execute_data(sql);
                        try
                        {
                            num = decimal.Parse(this._lib.get_data("select count(maql) from " + user + ".bc_bhytmalk").Tables[0].Rows[0][0].ToString());
                        }
                        catch
                        {
                            num = 0M;
                        }
                        this._lib.execute_data(string.Concat(new object[] { "insert into ", user, ".bc_bhytmalk  select maql,substr(to_char(maql),1,6)||lpad(to_char(stt),9,'0') as malk from ( select rownum+", num, " as stt,  maql from ", user, ".benhandt where maql not in(select maql from ", user, ".bc_bhytmalk) order by maql)" }));
                        if (loaibn == 1)
                        {
                            return;
                        }
                    }
                    time2 = time2.AddMonths(1);
                }
            }
            catch
            {
            }
        }

        public void f_CreateView_BHYT(int loaibn, string mmyy)
        {
            try
            {
                if (mmyy == "")
                {
                    mmyy = DateTime.Now.ToString("MMyy");
                }
                DateTime time = new DateTime(int.Parse(mmyy.Substring(2)) + 0x7d0, int.Parse(mmyy.Substring(0, 2)), 1);
                DateTime time2 = time;
                if (loaibn == 1)
                {
                    mmyy = "";
                }
                else
                {
                    time = time.AddMonths(1);
                    time2 = time.AddMonths(-2);
                }
                string str = "";
                string str2 = "";
                string str3 = "";
                string str4 = "";
                while (time2 <= time)
                {
                    if (this._lib.bMmyy(time2.ToString("MMyy")))
                    {
                        str2 = "#" + this.f_get_sql_sotheBHYT(loaibn, time2.ToString("MMyy"));
                        str = str + str2;
                        str4 = "#" + this.f_get_sql_benhandt(loaibn, time2.ToString("MMyy"));
                        str3 = str3 + str4;
                    }
                    time2 = time2.AddMonths(1);
                    if (loaibn == 1)
                    {
                        break;
                    }
                }
                str = str.Trim(new char[] { '#' }).Replace("#", " union all ");
                str3 = str3.Trim(new char[] { '#' }).Replace("#", " union all ");
                if (str != "")
                {
                    str2 = "create or replace view " + this._lib.user + mmyy + ".vi_thebhyt" + ((loaibn == 2) ? "ngtr" : "") + " (maql,sothe,mabv,traituyen,tungay,denngay,malk,makhuvuc) as  select distinct * from (" + str + ") ";
                    this._lib.execute_data(str2.Trim(new char[] { '_' }));
                    str4 = "create or replace view " + this._lib.user + mmyy + ".vi_benhandt" + ((loaibn == 2) ? "ngtr" : "") + " (maql,mabn,ngay,maicd,chandoan,nhantu,malk,mabs,makp,loaiba,mavaovien,ngayrv) as  select distinct maql,mabn,ngay,maicd,chandoan,nhantu,malk,mabs,makp,loaiba,mavaovien,ngayrv from (" + str3 + ") ";
                    this._lib.execute_data(str4.Trim(new char[] { '_' }));
                }
            }
            catch
            {
            }
        }

        private void f_exp_excel_mau808_run(bool print, DataSet ds11, int loaibc, string tungay, string denngay, string fontchu)
        {
            DataRow row = ds11.Tables[0].NewRow();
            int ordinal = ds11.Tables[0].Columns["t_tongchi"].Ordinal;
            for (int i = 0; i < ds11.Tables[0].Columns.Count; i++)
            {
                row[i] = (i < ordinal) ? Convert.ToChar((int) (i + 0x41)).ToString() : (((i + 1) - ordinal));
            }
            ds11.Tables[0].Rows.InsertAt(row, 0);
            int num3 = 0;
            int num4 = 3;
            int num5 = 5;
            int num6 = ds11.Tables[0].Rows.Count + 5;
            int num7 = ds11.Tables[0].Columns.Count - 1;
            num3 = num6;
            this.tenfile = this._lib.Export_Excel(ds11, "bccpkcb");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num4; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num4 + 7) + num5.ToString(), this._lib.getIndex(ds11.Tables[0].Columns["t_bhtt"].Ordinal) + num6.ToString()).NumberFormat = "#,##0";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(num7) + (num3 - 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num6, this._lib.getIndex(num7 + 3) + num6);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Bold = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num5, this._lib.getIndex(num7 + 3) + num5);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.Font.Bold = true;
                this.orange.RowHeight = 15;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(num7 + 2) + num6.ToString());
                this.orange.Font.Name = fontchu;
                this.orange.Font.Size = 12;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 3] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT " + ((loaibc == 0) ? "NGOẠI TR\x00da" : "NỮI TR\x00da");
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(num7) + "1");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.osheet.Cells[2, 3] = "Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "2", this._lib.getIndex(num7) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                int num9 = num5;
                for (int k = 1; k < ds11.Tables[0].Rows.Count; k++)
                {
                    num9++;
                    this.orange = this.osheet.get_Range("A" + num9.ToString(), this._lib.getIndex(num7 - 1) + num9.ToString());
                    if (((ds11.Tables[0].Rows[k]["stt"].ToString() == "A") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((ds11.Tables[0].Rows[k]["stt"].ToString() == "C") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "I") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "II"))
                    {
                        this.orange.Font.ColorIndex = 10;
                        this.orange.Font.Bold = true;
                    }
                    if (ds11.Tables[0].Rows[k]["t_tongchi"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range("B" + num9, this._lib.getIndex(num7) + num9);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "") || !char.IsDigit(ds11.Tables[0].Rows[k]["stt"].ToString(), 0))
                    {
                        this.orange = this.osheet.get_Range("B" + num9, this._lib.getIndex(ds11.Tables[0].Columns["ngay_ra"].Ordinal) + num9);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                }
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception.Message, this._lib.Msg);
            }
        }

        public string f_get_cdkemtheo(System.Data.DataTable dtkemtheo, string maql)
        {
            string str = ",";
            DataRow[] rowArray = dtkemtheo.Select("maql=" + maql);
            for (int i = 0; i < rowArray.Length; i++)
            {
                if (str.IndexOf("," + rowArray[i]["maicd"].ToString() + ",") <= -1)
                {
                    str = str + rowArray[i]["maicd"].ToString() + ",";
                }
            }
            return str.Trim(new char[] { ',' });
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

        private int f_get_madk_thebhyt(string mabv_bn)
        {
            try
            {
                return this.f_get_madk_thebhyt(mabv_bn, "");
            }
            catch
            {
                return 3;
            }
        }

        public int f_get_madk_thebhyt(string mabv_bn, string sothebhyt)
        {
            try
            {
                string str = "";
                string[] strArray = this._sViTriTheTrongTinh.Split(new char[] { ',' });
                for (int i = 0; i < strArray.Length; i++)
                {
                    str = str + sothebhyt.Substring(int.Parse(strArray[i]), 1);
                }
                if (mabv_bn == this._lib.Mabv)
                {
                    return 1;
                }
                if (((str == "") && (mabv_bn.Substring(0, 2) == this._lib.MABHXH)) || ((str == this._lib.MABHXH) && (str != "")))
                {
                    return 2;
                }
                return 3;
            }
            catch
            {
                return 3;
            }
        }

        private int f_get_madk_thebhyt1(string sothe, string manoidkbd)
        {
            string str = this._lib.thetrongtinh_vitri_old;
            string str2 = this._lib.thetrongtinh();
            string mabv = this._lib.Mabv;
            string str4 = "";
            try
            {
                if (sothe.ToString().Length == 0x12)
                {
                    str4 = sothe.Substring(13, 2);
                }
                else if ((sothe.Length == 13) || (sothe.Length == 0x10))
                {
                    str4 = sothe.Substring(int.Parse(str.Substring(0, str.IndexOf(","))), int.Parse(str.Substring(str.IndexOf(",") + 1)));
                }
                else if (sothe.Length == 20)
                {
                    str4 = sothe.Substring(int.Parse(str.Substring(0, str.IndexOf(","))), int.Parse(str.Substring(str.IndexOf(",") + 1)));
                }
                else
                {
                    str4 = str2;
                }
                if ((mabv == manoidkbd) && (str4 == this._lib.MABHXH))
                {
                    return 1;
                }
                if ((str4 == str2) && (mabv != manoidkbd))
                {
                    return 2;
                }
                return 3;
            }
            catch
            {
                return 1;
            }
        }

        public string f_get_MaNoiDKKB_BD(string sothe)
        {
            string str = "";
            if ((sothe.Length == 13) || (sothe.Length == 0x10))
            {
                return (sothe.Substring(int.Parse(this._sViTriTheMoi.Substring(0, this._sViTriTheMoi.IndexOf(","))), 2) + sothe.Substring(int.Parse(this._sViTriTheMoi.Substring(0, this._sViTriTheMoi.IndexOf(","))) + 3, 3));
            }
            try
            {
                str = sothe.Substring(sothe.Length - 5, 5);
            }
            catch
            {
            }
            return str;
        }

        public string f_get_MaTheBHYT(string sothe)
        {
            string str = "";
            if (sothe.Length == 0x12)
            {
                return sothe.Substring(0, 13);
            }
            if ((sothe.Length == 13) || (sothe.Length == 0x10))
            {
                return sothe.Substring(0, int.Parse(this._sViTriTheMoi.Substring(0, this._sViTriTheMoi.IndexOf(","))));
            }
            if (sothe.Length == 20)
            {
                str = sothe.Substring(0, 15);
            }
            return str;
        }

        private string f_get_sql_benhandt(int loaibn, string mmyy)
        {
            string str = "";
            string user = this._lib.user;
            if (loaibn == 1)
            {
                return ("select a.ngay,a.maql,a.mabn,a.maicd,a.chandoan,a.nhantu,case when a2.malk is null then 0 else a2.malk end as malk,a.mabs,case when a.mangtr is null or a.mangtr=0 then a.loaiba else 2 end as loaiba,a.mavaovien,a3.ngay as ngayrv,a.makp from " + user + ".benhandt  a left join " + user + ".bc_bhytmalk a2 on a2.maql=a.maql left join " + user + ".xuatvien a3 on a.maql=a3.maql");
            }
            if (loaibn == 2)
            {
                return ("select distinct  a.ngay,a.maql,a.mabn,a.maicd,a.chandoan,a.nhantu,a.malk,mabs,loaiba,mavaovien,ngayrv,a.makp from (select a.ngay,a.maql,a.mabn,a.maicd,a.chandoan,a.nhantu,case when a2.malk is null then 0 else a2.malk end as malk,a.mabs,a.mavaovien,case when a.mangtr is null or a.mangtr=0 then a.loaiba else 2 end as loaiba,a3.ngay as ngayrv,a.makp from " + user + mmyy + ".benhandt  a  left join " + user + mmyy + ".bc_bhytmalk a2 on a2.maql=a.maql left join " + user + mmyy + ".xuatvien a3 on a.maql=a3.maql union all select a.ngay,a.maql,a.mabn,a.maicd,a.chandoan,a.nhantu,case when a2.malk is null then 0 else a2.malk end as malk,a.mabs,a.mavaovien,case when a.mangtr is null or a.mangtr=0 then a.loaiba else 2 end as loaiba,a3.ngay as ngayrv,a.makp from " + user + ".benhandt a  left join " + user + ".bc_bhytmalk a2 on a2.maql=a.maql left join " + user + ".xuatvien a3 on a.maql=a3.maql) a");
            }
            if (loaibn == 3)
            {
                str = "select a.ngay,a.maql,a.mabn,a.maicd,a.chandoan,a.nhantu,case when a2.malk is null then 0 else a2.malk end as malk,a.mabs,a.mavaovien,case when a.mangtr is null or a.mangtr=0 then a.loaiba else 2 end as loaiba,a3.ngay as ngayrv,a.makp from " + user + mmyy + ".benhandt  a left join " + user + mmyy + ".bc_bhytmalk a2 on a2.maql=a.maql left join " + user + mmyy + ".xuatvien a3 on a.maql=a3.maql";
            }
            return str;
        }

        public string f_get_sql_BenhAnDT(int loaibn, string mmyy)
        {
            string user = this._lib.user;
            if (loaibn == 1)
            {
                mmyy = "";
            }
            return ("select * from " + user + mmyy + ".vi_benhandt" + ((loaibn == 2) ? "ngtr" : ""));
        }

        private string f_get_sql_sotheBHYT(int loaibn, string mmyy)
        {
            string str = "";
            string user = this._lib.user;
            if (loaibn == 1)
            {
                return ("select a.maql,a.sothe,mabv,traituyen,a.tungay,a.denngay,a2.malk,a.makhuvuc from " + user + ".bhyt a  inner join (select maql,max(denngay) denngay  from " + user + ".bhyt where denngay is not null  group by maql) b  on a.maql=b.maql and a.denngay=b.denngay and a.sudung=1  left join " + user + ".bc_bhytmalk a2 on a2.maql=a.maql");
            }
            if (loaibn == 2)
            {
                return ("select a.maql,a.sothe,mabv,traituyen,a.tungay,a.denngay,a.malk,a.makhuvuc  from (select a.maql,sothe,tungay,denngay,traituyen,mabv,a2.malk,a.makhuvuc from " + user + mmyy + ".bhyt a left join " + user + mmyy + ".bc_bhytmalk a2 on a2.maql=a.maql where sudung=1  union all select a.maql,sothe,tungay,denngay,traituyen,mabv,malk,a.makhuvuc from " + user + ".bhyt a left join " + user + ".bc_bhytmalk a2 on a2.maql=a.maql where sudung=1)a   inner join (select maql,max(denngay) denngay  from (select maql,denngay from " + user + mmyy + ".bhyt where sudung=1  union all select maql,denngay from " + user + ".bhyt where sudung=1 ) where denngay is not null group by maql) b  on a.maql=b.maql and a.denngay=b.denngay");
            }
            if (loaibn == 3)
            {
                str = " select a.maql,a.sothe,mabv,traituyen,a.tungay,a.denngay,a2.malk,a.makhuvuc from " + user + mmyy + ".bhyt a  inner join (select maql,max(denngay) denngay  from " + user + mmyy + ".bhyt where denngay is not null and sudung=1  group by maql) b  on a.maql=b.maql and a.denngay=b.denngay and a.sudung=1 left join " + user + mmyy + ".bc_bhytmalk a2 on a2.maql=a.maql";
            }
            return str;
        }

        public string f_get_sql_theBHYT(int loaibn, string mmyy)
        {
            string user = this._lib.user;
            if (loaibn == 1)
            {
                mmyy = "";
            }
            return ("select maql,sothe,mabv,traituyen,tungay,denngay,malk,makhuvuc from " + user + mmyy + ".vi_thebhyt" + ((loaibn == 2) ? "ngtr" : ""));
        }

        public string f_get_tenFileReport(int idmaubaocao)
        {
            try
            {
                DataSet set = new DataSet();
                set.Tables.Add(this.f_getdata_maubaocao(0, false));
                set.Merge(this.f_getdata_maubaocao(1, false));
                return set.Tables[0].Select("id=" + idmaubaocao)[0]["report"].ToString();
            }
            catch
            {
                return "";
            }
        }

        public decimal f_get_TyLeTheBHYT(string sothe, int traituyen)
        {
            return this.f_get_TyLeTheBHYT(sothe, traituyen, -1);
        }

        public decimal f_get_TyLeTheBHYT(string sothe, int traituyen, int mavp)
        {
            decimal num = 100M;
            string str = "+" + sothe.Substring(this._iSoThe15KiTu_vitri, this._iSoThe15KiTu_ChieuDai) + "+";
            try
            {
                if ((sothe.Length == 15) || (sothe.Length == 20))
                {
                    if (this._sSoThe15KiTu_kitu80.IndexOf(str) > -1)
                    {
                        num = 80M;
                    }
                    else if (this._sSoThe15KiTu_kitu95.IndexOf(str) > -1)
                    {
                        num = 95M;
                    }
                }
                if ((traituyen == 1) && this._bQLTraiTuyen)
                {
                    num = this._deTyLeTraiTuyen;
                    if (mavp > 0)
                    {
                        DataRow[] rowArray = this._dsgiavp.Tables[0].Select("mavp=" + mavp);
                        if (((rowArray.Length > 0) && (decimal.Parse(rowArray[0]["bhyt_tt"].ToString()) > 0M)) && (decimal.Parse(rowArray[0]["bhyt_tt"].ToString()) < this._deTyLeTraiTuyen))
                        {
                            num = decimal.Parse(rowArray[0]["bhyt_tt"].ToString());
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return num;
        }

        public System.Data.DataTable f_getdata_cdkemtheo(int loaibn, string mmyy, string maql)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            if (loaibn == 1)
            {
                return this._lib.get_data("select maql,maicd,chandoan from " + this._lib.user + ".cdkemtheo where 1=1 " + (((maql != "") != null) ? (" and maql=" + maql) : "")).Tables[0];
            }
            if (loaibn == 2)
            {
                try
                {
                    table = this._lib.get_data("select maql,maicd,chandoan from " + this._lib.user + ".cdkemtheo where 1=1 " + (((maql != "") != null) ? (" and maql=" + maql) : "") + " union all select maql,maicd,chandoan from " + this._lib.user + mmyy + ".cdkemtheo where 1=1 " + (((maql != "") != null) ? (" and maql=" + maql) : "")).Tables[0];
                }
                catch
                {
                }
                return table;
            }
            if (loaibn == 3)
            {
                table = this._lib.get_data(" select maql,maicd,chandoan from " + this._lib.user + mmyy + ".cdkemtheo where 1=1 " + (((maql != "") != null) ? (" and maql=" + maql) : "")).Tables[0];
            }
            return table;
        }

        public System.Data.DataTable f_getdata_maubaocao(int loaibc, bool laytatca)
        {
            DataSet set = new DataSet();
            System.Data.DataTable table = new System.Data.DataTable();
            try
            {
                set.ReadXml(@"..\..\..\xml\maubaocao_bhyt.xml", XmlReadMode.ReadSchema);
                string columnName = set.Tables[0].Columns["stt"].ColumnName;
                columnName = set.Tables[0].Columns["ten"].ColumnName;
                columnName = set.Tables[0].Columns["sudung"].ColumnName;
                columnName = set.Tables[0].Columns["report"].ColumnName;
                columnName = set.Tables[0].Columns["loai"].ColumnName;
                columnName = set.Tables[0].Columns["id"].ColumnName;
            }
            catch
            {
                set = new DataSet();
                set = this.f_getdata_maubaocao_macdinh(loaibc);
            }
            table = set.Tables[0].Clone();
            string str2 = "";
            if (!laytatca)
            {
                str2 = " and sudung=1";
            }
            foreach (DataRow row in set.Tables[0].Select("loai=" + loaibc + str2, "stt"))
            {
                table.Rows.Add(row.ItemArray);
            }
            return table;
        }

        public DataSet f_getdata_maubaocao_macdinh(int loaibc)
        {
            DataSet set = new DataSet();
            set.Tables.Add("maubaocao");
            set.Tables[0].Columns.Add("id");
            set.Tables[0].Columns.Add("loai");
            set.Tables[0].Columns.Add("stt");
            set.Tables[0].Columns.Add("ten");
            set.Tables[0].Columns.Add("report");
            set.Tables[0].Columns.Add("sudung");
            string[] strArray = new string[] { 
                "1;0;0;Mẫu excel theo CV 808;;1", "2;0;0;Mẫu excel 41 cột;;1", "3;0;0;Mẫu BHYT 79a-HD;;1", "4;1;0;Mẫu BHYT 80a-HD;;1", "5;1;0;Mẫu excel theo CV 808;;1", "6;1;0;Mẫu excel 41 cột;;1", "7;0;0;Mẫu BHYT 25a tổng hợp;;1", "8;1;0;Mẫu BHYT 26a tổng hợp;;1", "9;0;0;Mẫu BHYT 25a chi tiết;;1", "10;1;0;Mẫu BHYT 26a chi tiết;;1", "11;0;0;Mẫu BHYT 01(ngoại) theo TT 2348;;1", "12;1;0;Mẫu BHYT 01(nội) theo TT 2348;;1", "14;0;0;Mẫu BHYT 01(ngoại) theo TT 9324;;1", "15;1;0;Mẫu BHYT 01(nội) theo TT 9324;;1", "16;0;0;Mẫu BHYT 01(ngoại) theo TT 324-2016;;1", "17;1;0;Mẫu BHYT 01(nội) theo TT 324-2016;;1", 
                "18;0;0;Mẫu BHYT ngoại theo TT 324 mới;;1", "19;1;0;Mẫu BHYT nội theo TT 324 mới;;1", "20;0;0;Mẫu BHYT 79a-HD(mẫu 2);;1", "21;1;0;Mẫu BHYT 80a-HD(mẫu 2);;1", "22;1;0;Mẫu BHYT CV4210(nội tr\x00fa);;1", "23;0;0;Mẫu BHYT CV4210(ngoại tr\x00fa);;1", "13;1;0;Mẫu BC 38 cột;;1"
             };
            for (int i = 0; i < strArray.Length; i++)
            {
                string[] strArray2 = strArray[i].Split(new char[] { ';' });
                if (strArray2[1] == loaibc.ToString())
                {
                    DataRow row = set.Tables[0].NewRow();
                    row[0] = strArray2[0];
                    row[1] = strArray2[1];
                    row[2] = strArray2[2];
                    row[3] = strArray2[3];
                    row[4] = strArray2[4];
                    row[5] = strArray2[5];
                    set.Tables[0].Rows.Add(row);
                }
            }
            return set;
        }

        private string f_getstr_TagXML(string fieldName, string values)
        {
            return ("<" + fieldName.ToUpper() + ">" + values + "</" + fieldName.ToUpper() + ">");
        }

        public void f_ins_items_thuoc_mau19_6(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add();
            dset.Tables[0].Columns.Add("mavp", typeof(decimal));
            dset.Tables[0].Columns.Add("stt");
            dset.Tables[0].Columns.Add("ma_vtyt");
            dset.Tables[0].Columns.Add("ten_vtyt");
            dset.Tables[0].Columns.Add("ten_thuongmai");
            dset.Tables[0].Columns.Add("quy_cach");
            dset.Tables[0].Columns.Add("don_vi");
            dset.Tables[0].Columns.Add("gia_mua", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sl_noitru", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sl_ngoaitru", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("gia_thanhtoan", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            foreach (DataRow row in dsdulieu.Tables[0].Select("", "maql,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and gia_mua=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(dset.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3 = dset.Tables[0].NewRow();
                    row3["stt"] = dset.Tables[0].Rows.Count + 1;
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    row3["ten_thuongmai"] = row["ten"].ToString();
                    row3["don_vi"] = row["dvt"].ToString();
                    row3["quy_cach"] = row["donvi"].ToString();
                    try
                    {
                        row3["ma_vtyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString() + "")[0]["mabyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_vtyt"].ToString() == "")
                    {
                        row3["ma_vtyt"] = row["mavp1"].ToString();
                    }
                    try
                    {
                        row3["ten_vtyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                    }
                    catch
                    {
                    }
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noitru"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noitru"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoaitru"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoaitru"] = 0;
                        }
                    }
                    try
                    {
                        row3["gia_mua"] = Math.Round(Convert.ToDecimal(row["dongia"].ToString()));
                    }
                    catch
                    {
                        row3["gia_mua"] = 0;
                    }
                    try
                    {
                        row3["gia_thanhtoan"] = Math.Round(Convert.ToDecimal(row["dongia"].ToString()));
                    }
                    catch
                    {
                        row3["gia_thanhtoan"] = 0;
                    }
                    row3["thanh_tien"] = Math.Round(Convert.ToDecimal(row["sotien"].ToString()));
                    dset.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noitru"] = Math.Round((decimal) (Convert.ToDecimal(row2["sl_noitru"].ToString()) + Convert.ToDecimal(row["soluong"].ToString())));
                        row2["thanh_tien"] = Math.Round((decimal) (Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString())));
                    }
                    catch
                    {
                        row2["sl_noitru"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoaitru"] = Math.Round((decimal) (Convert.ToDecimal(row2["sl_ngoaitru"].ToString()) + Convert.ToDecimal(row["soluong"].ToString())));
                        row2["thanh_tien"] = Math.Round((decimal) (Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString())));
                    }
                    catch
                    {
                        row2["sl_ngoaitru"] = 0;
                    }
                }
            }
            dset.Tables[0].Columns.Remove("mavp");
            dset.AcceptChanges();
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb19_4");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }

        public void f_ins_items_thuoc_mau19_tt1399(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet vds = new DataSet();
            vds.Tables.Add();
            vds.Tables[0].Columns.Add("mavp", typeof(decimal));
            vds.Tables[0].Columns.Add("stt");
            vds.Tables[0].Columns.Add("ma_bhyt");
            vds.Tables[0].Columns.Add("ten_bhyt");
            vds.Tables[0].Columns.Add("ten_thuoc");
            vds.Tables[0].Columns.Add("quy_cach");
            vds.Tables[0].Columns.Add("don_vi_tinh");
            vds.Tables[0].Columns.Add("don_gia_mua", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("sl_ngoai", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("sl_noi", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("don_gia", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            foreach (DataRow row in dsdulieu.Tables[0].Select("nhombhyt in(6,14)", "maql,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and don_gia=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(vds.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3 = vds.Tables[0].NewRow();
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    row3["ten_thuoc"] = row["ten"].ToString();
                    row3["don_vi_tinh"] = row["dvt"].ToString();
                    try
                    {
                        row3["quy_cach"] = row["donvi"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row3["ma_bhyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString() + "")[0]["mabyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_bhyt"].ToString() == "")
                    {
                        row3["ma_bhyt"] = row["mavp1"].ToString();
                    }
                    try
                    {
                        row3["ten_bhyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                    }
                    catch
                    {
                    }
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noi"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noi"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoai"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoai"] = 0;
                        }
                    }
                    try
                    {
                        row3["don_gia_mua"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia_mua"] = 0;
                    }
                    try
                    {
                        row3["don_gia"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia"] = 0;
                    }
                    row3["thanh_tien"] = row["sotien"].ToString();
                    vds.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noi"] = Convert.ToDecimal(row2["sl_noi"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_noi"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoai"] = Convert.ToDecimal(row2["sl_ngoai"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_ngoai"] = 0;
                    }
                }
            }
            vds.Tables[0].Columns.Remove("mavp");
            vds.AcceptChanges();
            this.exp_excel_1399(vds, tungay, denngay);
        }

        public void f_ins_items_thuoc_mau20_4(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add();
            dset.Tables[0].Columns.Add("mavp", typeof(decimal));
            dset.Tables[0].Columns.Add("stt");
            dset.Tables[0].Columns.Add("ma_bhyt");
            dset.Tables[0].Columns.Add("ten_hoat_chat");
            dset.Tables[0].Columns.Add("ten_thuoc");
            dset.Tables[0].Columns.Add("duong_dung");
            dset.Tables[0].Columns.Add("ham_luong");
            dset.Tables[0].Columns.Add("sodk");
            dset.Tables[0].Columns.Add("don_vi_tinh");
            dset.Tables[0].Columns.Add("sl_ngoai", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sl_noi", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("don_gia", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            foreach (DataRow row in dsdulieu.Tables[0].Select("nhombhyt=3", "nhombhyt,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and don_gia=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(dset.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3 = dset.Tables[0].NewRow();
                    row3["stt"] = dset.Tables[0].Rows.Count + 1;
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    try
                    {
                        row3["ma_bhyt"] = row["masobyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_bhyt"].ToString() == "")
                    {
                        row3["ma_bhyt"] = row["mavp1"].ToString();
                    }
                    row3["ten_hoat_chat"] = row["tenhc"].ToString();
                    try
                    {
                        row3["ten_thuoc"] = row["tenbyt"].ToString();
                    }
                    catch
                    {
                        row3["ten_thuoc"] = row["ten"].ToString();
                    }
                    row3["duong_dung"] = row["duongdung"].ToString();
                    row3["ham_luong"] = row["hamluong"].ToString();
                    row3["don_vi_tinh"] = row["dvt"].ToString();
                    row3["sodk"] = row["sodk"].ToString();
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noi"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noi"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoai"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoai"] = 0;
                        }
                    }
                    try
                    {
                        row3["don_gia"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia"] = 0;
                    }
                    row3["thanh_tien"] = row["sotien"].ToString();
                    dset.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noi"] = Convert.ToDecimal(row2["sl_noi"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_noi"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoai"] = Convert.ToDecimal(row2["sl_ngoai"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_ngoai"] = 0;
                    }
                }
            }
            dset.Tables[0].Columns["ma_bhyt"].ColumnName = "ma_thuoc";
            dset.Tables[0].Columns["ten_hoat_chat"].ColumnName = "ten_hoatchat";
            dset.Tables[0].Columns["sodk"].ColumnName = "so_dky";
            dset.Tables[0].Columns["don_vi_tinh"].ColumnName = "don_vi";
            dset.Tables[0].Columns["sl_ngoai"].ColumnName = "sl_noitru";
            dset.Tables[0].Columns["sl_noi"].ColumnName = "sl_ngoaitru";
            dset.Tables[0].Columns.Remove("mavp");
            dset.AcceptChanges();
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb20_6");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }

        public void f_ins_items_thuoc_mau20_tt1399(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet vds = new DataSet();
            vds.Tables.Add();
            vds.Tables[0].Columns.Add("mavp", typeof(decimal));
            vds.Tables[0].Columns.Add("stt");
            vds.Tables[0].Columns.Add("ma_bhyt");
            vds.Tables[0].Columns.Add("ten_hoat_chat");
            vds.Tables[0].Columns.Add("ten_thuoc");
            vds.Tables[0].Columns.Add("duong_dung");
            vds.Tables[0].Columns.Add("ham_luong");
            vds.Tables[0].Columns.Add("sodk");
            vds.Tables[0].Columns.Add("don_vi_tinh");
            vds.Tables[0].Columns.Add("sl_ngoai", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("sl_noi", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("don_gia", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            foreach (DataRow row in dsdulieu.Tables[0].Select("nhombhyt=3", "nhombhyt,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and don_gia=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(vds.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3 = vds.Tables[0].NewRow();
                    row3["stt"] = vds.Tables[0].Rows.Count + 1;
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    try
                    {
                        row3["ma_bhyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString())[0]["mabyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_bhyt"].ToString() == "")
                    {
                        row3["ma_bhyt"] = row["mavp1"].ToString();
                    }
                    row3["ten_hoat_chat"] = row["tenhc"].ToString();
                    try
                    {
                        row3["ten_thuoc"] = set2.Tables[0].Select("id=" + row["mavp"].ToString())[0]["tenbyt"].ToString();
                    }
                    catch
                    {
                        row3["ten_thuoc"] = row["ten"].ToString();
                    }
                    row3["duong_dung"] = row["duongdung"].ToString();
                    row3["ham_luong"] = row["hamluong"].ToString();
                    row3["don_vi_tinh"] = row["dvt"].ToString();
                    row3["sodk"] = row["sodk"].ToString();
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noi"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noi"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoai"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoai"] = 0;
                        }
                    }
                    try
                    {
                        row3["don_gia"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia"] = 0;
                    }
                    row3["thanh_tien"] = row["sotien"].ToString();
                    vds.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noi"] = Convert.ToDecimal(row2["sl_noi"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_noi"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoai"] = Convert.ToDecimal(row2["sl_ngoai"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_ngoai"] = 0;
                    }
                }
            }
            vds.Tables[0].Columns.Remove("mavp");
            vds.AcceptChanges();
            this.exp_excel_1399_20(vds, tungay, denngay);
        }

        public void f_ins_items_thuoc_mau21_5(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add();
            dset.Tables[0].Columns.Add("mavp", typeof(decimal));
            dset.Tables[0].Columns.Add("stt");
            dset.Tables[0].Columns.Add("ma_bhyt");
            dset.Tables[0].Columns.Add("ten");
            dset.Tables[0].Columns.Add("sl_ngoai", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sl_noi", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("don_gia", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            foreach (DataRow row in dsdulieu.Tables[0].Select("", "nhombhyt,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and don_gia=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(dset.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3 = dset.Tables[0].NewRow();
                    row3["stt"] = dset.Tables[0].Rows.Count + 1;
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    try
                    {
                        row3["ma_bhyt"] = row["masobyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_bhyt"].ToString() == "")
                    {
                        row3["ma_bhyt"] = row["mavp1"].ToString();
                    }
                    try
                    {
                        row3["ten"] = row["tenbyt"].ToString();
                    }
                    catch
                    {
                        row3["ten"] = row["ten"].ToString();
                    }
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noi"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noi"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoai"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoai"] = 0;
                        }
                    }
                    try
                    {
                        row3["don_gia"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia"] = 0;
                    }
                    row3["thanh_tien"] = row["sotien"].ToString();
                    dset.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noi"] = Convert.ToDecimal(row2["sl_noi"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_noi"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoai"] = Convert.ToDecimal(row2["sl_ngoai"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_ngoai"] = 0;
                    }
                }
            }
            dset.Tables[0].Columns["ma_bhyt"].ColumnName = "ma_dvkt";
            dset.Tables[0].Columns["ten"].ColumnName = "ten_dvkt";
            dset.Tables[0].Columns["sl_ngoai"].ColumnName = "sl_ngoaitru";
            dset.Tables[0].Columns["sl_noi"].ColumnName = "sl_noitru";
            dset.Tables[0].Columns.Remove("mavp");
            dset.AcceptChanges();
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb21_5");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }

        public void f_ins_items_thuoc_mau21_tt1399(DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet vds = new DataSet();
            vds.Tables.Add();
            vds.Tables[0].Columns.Add("mavp", typeof(decimal));
            vds.Tables[0].Columns.Add("stt");
            vds.Tables[0].Columns.Add("ma_bhyt");
            vds.Tables[0].Columns.Add("ten");
            vds.Tables[0].Columns.Add("sl_ngoai", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("sl_noi", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("don_gia", typeof(decimal)).DefaultValue = 0;
            vds.Tables[0].Columns.Add("thanh_tien", typeof(decimal)).DefaultValue = 0;
            string exp = "";
            string str2 = "";
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            foreach (DataRow row in dsdulieu.Tables[0].Select("", "nhombhyt,mavp,dongia"))
            {
                exp = string.Concat(new object[] { "mavp=", row["mavp"].ToString(), " and don_gia=", Convert.ToDecimal(row["dongia"].ToString()) });
                DataRow row2 = this._lib.getrowbyid(vds.Tables[0], exp);
                if (row2 == null)
                {
                    DataRow row3;
                    if (str2 != row["nhombhyt"].ToString())
                    {
                        str2 = row["nhombhyt"].ToString();
                        row3 = vds.Tables[0].NewRow();
                        row3["stt"] = "";
                        row3["mavp"] = 0;
                        row3["don_gia"] = 0;
                        row3["ten"] = row["tennhombhyt"].ToString();
                        try
                        {
                            row3["sl_ngoai"] = dsdulieu.Tables[0].Compute("sum(soluong)", "loaiba in(3) and nhombhyt=" + str2).ToString();
                        }
                        catch
                        {
                        }
                        try
                        {
                            row3["sl_noi"] = dsdulieu.Tables[0].Compute("sum(soluong)", "loaiba in(1,4) and nhombhyt=" + str2).ToString();
                        }
                        catch
                        {
                        }
                        try
                        {
                            row3["thanh_tien"] = dsdulieu.Tables[0].Compute("sum(sotien)", " nhombhyt=" + str2).ToString();
                        }
                        catch
                        {
                        }
                        vds.Tables[0].Rows.Add(row3);
                    }
                    row3 = vds.Tables[0].NewRow();
                    row3["stt"] = vds.Tables[0].Rows.Count + 1;
                    row3["mavp"] = Convert.ToDecimal(row["mavp"].ToString());
                    try
                    {
                        row3["ma_bhyt"] = set2.Tables[0].Select("id=" + row["mavp"].ToString())[0]["mabyt"].ToString();
                    }
                    catch
                    {
                    }
                    if (row3["ma_bhyt"].ToString() == "")
                    {
                        row3["ma_bhyt"] = row["mavp1"].ToString();
                    }
                    try
                    {
                        row3["ten"] = set2.Tables[0].Select("id=" + row["mavp"].ToString())[0]["tenbyt"].ToString();
                    }
                    catch
                    {
                        row3["ten"] = row["ten"].ToString();
                    }
                    if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                    {
                        try
                        {
                            row3["sl_noi"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_noi"] = 0;
                        }
                    }
                    else
                    {
                        try
                        {
                            row3["sl_ngoai"] = Convert.ToDecimal(row["soluong"].ToString());
                        }
                        catch
                        {
                            row3["sl_ngoai"] = 0;
                        }
                    }
                    try
                    {
                        row3["don_gia"] = Convert.ToDecimal(row["dongia"].ToString());
                    }
                    catch
                    {
                        row3["don_gia"] = 0;
                    }
                    row3["thanh_tien"] = row["sotien"].ToString();
                    vds.Tables[0].Rows.Add(row3);
                }
                else if ((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4"))
                {
                    try
                    {
                        row2["sl_noi"] = Convert.ToDecimal(row2["sl_noi"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_noi"] = 0;
                    }
                }
                else
                {
                    try
                    {
                        row2["sl_ngoai"] = Convert.ToDecimal(row2["sl_ngoai"].ToString()) + Convert.ToDecimal(row["soluong"].ToString());
                        row2["thanh_tien"] = Convert.ToDecimal(row2["thanh_tien"].ToString()) + Convert.ToDecimal(row["sotien"].ToString());
                    }
                    catch
                    {
                        row2["sl_ngoai"] = 0;
                    }
                }
            }
            vds.Tables[0].Columns.Remove("mavp");
            vds.AcceptChanges();
            this.exp_excel_1399_21(vds, tungay, denngay);
        }

        private void f_LibMaubaocao_load()
        {
            if (!Directory.Exists("Excel"))
            {
                Directory.CreateDirectory("Excel");
            }
            this._sViTriTheMoi = this._lib.vitrithe_moi();
            this._iGiaBhyt = this._lib.bGiabhyt;
            this._sViTriTheTrongTinh = this._lib.thetrongtinh_vitri_old;
            this._iSoThe15KiTu_vitri = int.Parse(this._lib.themoi15_vitri().Split(new char[] { ',' })[0]) - 1;
            this._iSoThe15KiTu_ChieuDai = int.Parse(this._lib.themoi15_vitri().Split(new char[] { ',' })[1]);
            this._sSoThe15KiTu_kitu80 = this._lib.sothemoi15_80().Trim(new char[] { '+' });
            this._sSoThe15KiTu_kitu80 = "+" + this._sSoThe15KiTu_kitu80 + "+";
            this._sSoThe15KiTu_kitu95 = this._lib.sothemoi15_95().Trim(new char[] { '+' });
            this._sSoThe15KiTu_kitu95 = "+" + this._sSoThe15KiTu_kitu95 + "+";
            this._bQLTraiTuyen = this._lib.BHYT_traituyen() == 1M;
            this._deTyLeTraiTuyen = this._lib.tile_traituyen();
            this._dsgiavp = this._lib.get_data("select id mavp,ten tenvp,bhyt_tt from v_giavp");
        }

        private DataSet f_Ngoaitru_excel_mau41_getdata(DataSet dsdulieu, int namqt, int thangqt)
        {
            try
            {
                dsdulieu.Tables[0].Columns.Add("yyyymmdd");
            }
            catch
            {
            }
            try
            {
                dsdulieu.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            if (dsdulieu.Tables[0].Rows.Count > 0)
            {
                bool flag = this._tsxml.pChung_MauExcel41cot_hoten;
                foreach (DataRow row in dsdulieu.Tables[0].Rows)
                {
                    try
                    {
                        row["manoidk"] = ((row["manoidk"].ToString().Substring(0, 1) == "0") ? "'" : "") + row["manoidk"].ToString();
                    }
                    catch
                    {
                    }
                    if (flag)
                    {
                        row["hoten"] = row["hoten"].ToString().ToLower();
                    }
                    try
                    {
                        row["yyyymmdd"] = row["ngayra"].ToString().Substring(6, 4) + row["ngayra"].ToString().Substring(3, 2) + row["ngayra"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                    }
                }
            }
            DataSet set = new DataSet();
            set.Tables.Add("Table");
            set.Tables[0].Columns.Add(new DataColumn("stt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("hoten", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("namsinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gioitinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mathe", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ma_dkbd", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("makhoa", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mabenh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_vao", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_ra", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngaydtr", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_tongchi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_xn", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_cdha", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_thuoc", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_mau", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_pttt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtytth", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtyttt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_dvktc", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_ktg", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_kham", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vchuyen", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bnct", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bhtt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_ngoaids", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("lydo_vv", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("benhkhac", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noikcb", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("nam_qt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("thang_qt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_tu", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_den", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("diachi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("giamdinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_xuattoan", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("lydo_xt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_datuyen", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vuottran", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("loaikcb", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noi_ttoan", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("tt_tngt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mabn", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("sobienlai", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("quyenso", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TYLEBHYT", typeof(string)));
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "t_tongchi", "ngaydtr", "t_xn", "t_cdha", "t_thuoc", "t_mau", "t_pttt", "t_vtytth", "t_vtyttt", "t_dvktc", "t_ktg", "t_kham", "t_vchuyen", "t_bnct", "t_bhtt", "t_ngoaids" };
            string str = "";
            DataRow row2 = set.Tables[0].NewRow();
            DataRow row3 = set.Tables[0].NewRow();
            int num = 0;
            bool flag2 = this._tsxml.pNoitru_MauExcel41cot_noikcb;
            bool flag3 = !this._tsxml.pNoitru_MauExcel41cot_groupnhom;
            bool flag4 = this._tsxml.pChung_MauExcel41cot_SapXepNgay;
            bool flag5 = this._tsxml.pChung_MauExcel41cot_TheBHYTBo5KiTuCuoi;
            string filterExpression = "";
            string sort = "";
            for (int i = 1; i <= 3; i++)
            {
                switch (i)
                {
                    case 1:
                        str = "ABỆNH nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str = "BBỆNH nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str = "CBỆNH nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row4 = set.Tables[0].NewRow();
                row4["stt"] = str.Substring(0, 1);
                row4["hoten"] = str.Substring(1);
                if (flag3)
                {
                    set.Tables[0].Rows.Add(row4);
                }
                DataRow row5 = set.Tables[0].NewRow();
                for (int j = 0; j <= 1; j++)
                {
                    DataRow row6 = set.Tables[0].NewRow();
                    row6["stt"] = (j == 0) ? "I" : "II";
                    row6["hoten"] = (j == 0) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN";
                    if (flag3)
                    {
                        set.Tables[0].Rows.Add(row6);
                    }
                    DataRow row7 = set.Tables[0].NewRow();
                    if (flag3)
                    {
                        filterExpression = "madk=" + i.ToString() + " and traituyen=" + j.ToString();
                        sort = "yyyymmdd " + (flag4 ? " asc" : " desc");
                    }
                    else
                    {
                        filterExpression = "";
                        sort = "quyenso,sobienlai,yyyymmdd " + (flag4 ? " asc" : " desc");
                    }
                    foreach (DataRow row8 in dsdulieu.Tables[0].Select(filterExpression, sort))
                    {
                        DataRow row9 = set.Tables[0].NewRow();
                        row9["stt"] = row8["STT"].ToString();
                        row9["hoten"] = row8["HOTEN"].ToString();
                        row9["mathe"] = !flag5 ? row8["sothe"].ToString() : row8["sothe"].ToString().Substring(0, row8["sothe"].ToString().Length - 5);
                        row9["ma_dkbd"] = ((row8["MANOIDK"].ToString().Substring(0, 1) == "0") ? "'" : "") + row8["MANOIDK"].ToString();
                        row9["mabenh"] = row8["MAICD"].ToString();
                        row9["benhkhac"] = row8["maicdkt"].ToString();
                        row9["makhoa"] = row8["tenkp"].ToString();
                        row9["namsinh"] = row8["ngaysinh"].ToString();
                        if (row8["phai"].ToString().Trim() == "1")
                        {
                            row9["gioitinh"] = 2;
                        }
                        else
                        {
                            row9["gioitinh"] = 1;
                        }
                        row9["ngay_vao"] = row8["NGAYVAO"].ToString();
                        row9["ngay_ra"] = row8["NGAYra"].ToString();
                        try
                        {
                            row9["ngay_vao"] = row8["NGAYVAO"].ToString().Substring(6, 4) + row8["NGAYVAO"].ToString().Substring(3, 2) + row8["NGAYVAO"].ToString().Substring(0, 2) + row8["NGAYvv"].ToString().Substring(8);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row9["ngay_ra"] = row8["NGAYra"].ToString().Substring(6, 4) + row8["NGAYra"].ToString().Substring(3, 2) + row8["NGAYra"].ToString().Substring(0, 2) + row8["NGAYrv"].ToString().Substring(8);
                        }
                        catch
                        {
                        }
                        row9["ngaydtr"] = 1;
                        row9["t_xn"] = row8["ST_1"].ToString();
                        int num4 = 0;
                        int num5 = 0;
                        int num6 = 0;
                        int num7 = 0;
                        try
                        {
                            num4 = Convert.ToInt32(row8["ST_7"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            num5 = Convert.ToInt32(row8["ST_6"].ToString());
                        }
                        catch
                        {
                        }
                        row9["t_cdha"] = row8["ST_2"].ToString();
                        row9["t_thuoc"] = Convert.ToDecimal(row8["ST_3"].ToString());
                        row9["t_mau"] = row8["ST_4"].ToString();
                        try
                        {
                            num6 = Convert.ToInt32(row8["ST_8"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            num7 = Convert.ToInt32(row8["ST_10"].ToString());
                        }
                        catch
                        {
                        }
                        row9["t_pttt"] = row8["ST_5"].ToString();
                        row9["t_vtytth"] = row8["ST_6"].ToString();
                        row9["t_dvktc"] = row8["ST_7"].ToString();
                        try
                        {
                            row9["t_ktg"] = row8["ST_8"].ToString();
                        }
                        catch
                        {
                        }
                        try
                        {
                            row9["t_vtyttt"] = row8["ST_14"].ToString();
                        }
                        catch
                        {
                            row9["t_vtyttt"] = 0;
                        }
                        try
                        {
                            row9["t_kham"] = row8["ST_9"].ToString();
                        }
                        catch
                        {
                        }
                        try
                        {
                            row9["t_vchuyen"] = row8["ST_10"].ToString();
                        }
                        catch
                        {
                        }
                        row9["t_tongchi"] = row8["tongcong"].ToString();
                        row9["t_bhtt"] = row8["bhyttra"].ToString();
                        row9["t_bnct"] = row8["bntra"].ToString();
                        row9["t_ngoaids"] = 0;
                        if (row8["madoituong"].ToString() == "6")
                        {
                            row9["t_ngoaids"] = row8["bhyttra"].ToString();
                        }
                        if (row8["traituyen"].ToString().Trim() == "0")
                        {
                            row9["lydo_vv"] = 1;
                        }
                        if (row8["traituyen"].ToString().Trim() == "1")
                        {
                            row9["lydo_vv"] = 0;
                        }
                        try
                        {
                            if (row8["makp"].ToString().Trim() == "99")
                            {
                                row9["lydo_vv"] = 2;
                            }
                        }
                        catch
                        {
                        }
                        try
                        {
                            row9["nam_qt"] = namqt.ToString();
                            row9["thang_qt"] = thangqt.ToString();
                        }
                        catch
                        {
                        }
                        row9["noikcb"] = flag2 ? row9["ma_dkbd"].ToString() : this._lib.MABHXH;
                        row9["tylebhyt"] = row8["tylebhyt"].ToString();
                        row9["diachi"] = row8["diachi"].ToString();
                        row9["gt_tu"] = row8["GTRITU"].ToString();
                        row9["gt_den"] = row8["GTRIDEN"].ToString();
                        row9["loaikcb"] = "NGOAI";
                        row9["sobienlai"] = row8["sobienlai"].ToString();
                        row9["quyenso"] = row8["quyenso"].ToString();
                        row9["mabn"] = row8["mabn"].ToString();
                        row9["tt_tngt"] = (row8["tainangt"].ToString() != "") ? "TNGT" : "";
                        row9["stt"] = ++num;
                        set.Tables[0].Rows.Add(row9);
                        for (int k = 0; k < strArray2.Length; k++)
                        {
                            try
                            {
                                if (row7[strArray2[k]].ToString() == "")
                                {
                                    row7[strArray2[k]] = 0;
                                }
                                if (row2[strArray2[k]].ToString() == "")
                                {
                                    row2[strArray2[k]] = 0;
                                }
                                if (row5[strArray2[k]].ToString() == "")
                                {
                                    row5[strArray2[k]] = 0;
                                }
                                if (row9[strArray2[k]].ToString() == "")
                                {
                                    row9[strArray2[k]] = 0;
                                }
                                row7[strArray2[k]] = decimal.Parse(row7[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                                row2[strArray2[k]] = decimal.Parse(row2[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                                row5[strArray2[k]] = decimal.Parse(row5[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    if (!flag3)
                    {
                        break;
                    }
                    row7["stt"] = (j == 0) ? "I" : "II";
                    row7["hoten"] = "TỔNG " + ((j == 0) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN");
                    if ((row7["t_tongchi"].ToString() != "") && (row7["t_tongchi"].ToString() != "0"))
                    {
                        set.Tables[0].Rows.Add(row7);
                    }
                    else
                    {
                        set.Tables[0].Rows.Remove(row6);
                    }
                }
                if (!flag3)
                {
                    break;
                }
                row5["stt"] = str.Substring(0, 1);
                row5["hoten"] = "cộng " + row5["stt"].ToString();
                if ((row5["t_tongchi"].ToString() != "") && (row5["t_tongchi"].ToString() != "0"))
                {
                    set.Tables[0].Rows.Add(row5);
                }
                else
                {
                    set.Tables[0].Rows.RemoveAt(set.Tables[0].Rows.Count - 1);
                }
            }
            row2["hoten"] = "TỔNG cộng A+B+C";
            set.Tables[0].Rows.Add(row2);
            try
            {
                dsdulieu.Tables[0].Columns.Remove("yyyymmdd");
            }
            catch
            {
            }
            return set;
        }

        private void f_Ngoaitru_exp_excel_mau41_run(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            ds11 = this.f_setSapXepCotTheoThuTu(ds11, this._tsxml.pNgoaiTru_MauExcel41cot_DanhSachCotHienThi, '#');
            DataRow row = ds11.Tables[0].NewRow();
            for (int i = 0; i < ds11.Tables[0].Columns.Count; i++)
            {
                row[i] = "[" + (i + 1) + "]";
            }
            ds11.Tables[0].Rows.InsertAt(row, 0);
            int num2 = 0;
            int num3 = 3;
            int num4 = 5;
            int num5 = ds11.Tables[0].Rows.Count + 5;
            int num6 = ds11.Tables[0].Columns.Count - 1;
            num2 = num5;
            this.tenfile = this._lib.Export_Excel(ds11, "bccpkcb");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num3; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num3 + 7) + num4.ToString(), this._lib.getIndex(ds11.Tables[0].Columns["t_bhtt"].Ordinal) + num5.ToString()).NumberFormat = "#,##0";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(num6) + (num2 - 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num5, this._lib.getIndex(num6 + 3) + num5);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Bold = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num4, this._lib.getIndex(num6 + 3) + num4);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.Font.Bold = true;
                this.orange.RowHeight = 15;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(num6 + 2) + num5.ToString());
                this.orange.Font.Name = fontchu;
                this.orange.Font.Size = 12;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 3] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NGOẠI TR\x00da";
                int num8 = num4;
                for (int k = 1; k < ds11.Tables[0].Rows.Count; k++)
                {
                    num8++;
                    this.orange = this.osheet.get_Range("A" + num8.ToString(), this._lib.getIndex(num6 - 1) + num8.ToString());
                    if (((ds11.Tables[0].Rows[k]["stt"].ToString() == "A") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((ds11.Tables[0].Rows[k]["stt"].ToString() == "C") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "I") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "II"))
                    {
                        this.orange.Font.ColorIndex = 10;
                        this.orange.Font.Bold = true;
                    }
                    if (ds11.Tables[0].Rows[k]["t_tongchi"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range("B" + num8, this._lib.getIndex(num6) + num8);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "") || !char.IsDigit(ds11.Tables[0].Rows[k]["stt"].ToString(), 0))
                    {
                        this.orange = this.osheet.get_Range("B" + num8, this._lib.getIndex(ds11.Tables[0].Columns["ngay_ra"].Ordinal) + num8);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(num6) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.osheet.Cells[2, 3] = "Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "2", this._lib.getIndex(num6) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception.Message, this._lib.Msg);
            }
        }

        private void f_Ngoaitru_exp_excel_mau41_run2(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            ds11 = this.f_setSapXepCotTheoThuTu(ds11, this._tsxml.pNgoaiTru_MauExcel41cot_DanhSachCotHienThi, '#');
            DataRow row = ds11.Tables[0].NewRow();
            for (int i = 0; i < ds11.Tables[0].Columns.Count; i++)
            {
                row[i] = "[" + (i + 1) + "]";
            }
            ds11.Tables[0].Rows.InsertAt(row, 0);
            StringBuilder builder = new StringBuilder();
            builder.Append("<table border=0>");
            for (int j = 0; j < 10; j++)
            {
                builder.Append("<tr></tr>\n");
            }
            builder.Append("</table>\n");
            builder.Append("<table border=1>");
            for (int k = 0; k < ds11.Tables[0].Rows.Count; k++)
            {
            }
            builder.Append("</table>");
        }

        private void f_Ngoaitru_exp_excel_mau808_run(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            this.f_exp_excel_mau808_run(print, ds11, 0, tungay, denngay, fontchu);
        }

        public string f_Ngoaitru_getSql_chiphibn(string d_mmyy, string sTuNgay, string sDenNgay, string sMaDoiTuong, string sMaBacSi, string sMakp, string sSoTheBHYT, int iViTriTheBhyt, int iNhomBaoCao, string sLocMaTheBHYT, bool bToaThuocNgoaiTru, int iNhomKho, string sMaKho, string sNhomBhyt, bool bLaySoLieuThuocTuTruc, bool bInRiengSoLieuThuocTuTruc, bool bLaySoLieuVPNgoaiTru, bool bInRiengSoLieuVPNgoaiTru)
        {
            string str = "";
            string str2 = this.f_get_sql_theBHYT(3, d_mmyy);
            string str3 = this.f_get_sql_theBHYT(2, d_mmyy);
            string str4 = "";
            str4 = " and to_date(a.ngay,'dd/mm/yy')  between to_date('" + sTuNgay + "','dd/mm/yy')  and to_date('" + sDenNgay + "','dd/mm/yy') ";
            string str5 = "";
            string str6 = "";
            string str7 = "";
            string user = this._lib.user;
            str5 = user + d_mmyy;
            if (iNhomBaoCao == 0)
            {
                str6 = "h.idnhombhytmedisoft";
                str7 = "d.idnhombhytmedisoft";
            }
            else
            {
                str6 = "e.id";
                str7 = "b.ma";
            }
            string str9 = (int.Parse(this._sViTriTheMoi.Substring(0, this._sViTriTheMoi.IndexOf(","))) + 1) + "," + int.Parse(this._sViTriTheMoi.Substring(this._sViTriTheMoi.IndexOf(",") + 1));
            str = "select 0 loai,a.maql,a.mavaovien,a.maphu madoituong,a.id,a.sobienlai,to_char(a.ngay,'dd/mm/yyyy') ngay,null vaora,a.traituyen,nvl(bh.sothe,' ') sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,a.makp,bv.tenkp,a.mabn,bn.phai,f.hoten,f.namsinh,bn.sonha||' '||bn.thon||','||bn3.tenpxa||','||bn2.tenquan||','||bn1.tentt diachi,decode(dv.nhombv,null,a.mabv,dv.nhombv) mabv,g.tenbv," + str6 + " nhomvp,0 loaivp,a.congkham as congkham,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a1.nguyennhan=0 then '1' else '' end tngt";
            if (this._iGiaBhyt == 2)
            {
                str = str + ",sum(b.sotien) sotien ";
            }
            else
            {
                str = str + ",sum(b.soluong*" + ((this._iGiaBhyt == 0) ? "b.giamua" : "b.giaban") + ") sotien ";
            }
            str = str + " from " + user + "d" + d_mmyy + ".bhytkb a inner join " + user + "d" + d_mmyy + ".bhytthuoc b on a.id=b.id inner join " + user + "d" + d_mmyy + ".bhytds f on a.mabn=f.mabn inner join " + user + ".d_dmbd c on b.mabd=c.id inner join " + user + ".d_dmnhom d on c.manhom=d.id inner join  " + user + ".btdbn bn on f.mabn=bn.mabn left join  " + user + ".btdtt bn1 on bn1.matt=bn.matt left join  " + user + ".btdquan bn2 on bn2.maqu=bn.maqu left join  " + user + ".btdpxa bn3 on bn3.maphuongxa=bn.maphuongxa left join " + user + ".dmnoicapbhyt g on a.mabv=g.mabv inner join " + user + ".v_nhomvp e on d.nhomvp=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id left join " + user + ".tainantt a1 on a.maql=a1.maql left join " + user + ".dmnhomdv_bhyt dv on a.mabv=dv.mabv left join " + user + ".btdkp_bv bv on a.makp=bv.makp left join (" + str2 + ") bh on a.maql=bh.maql where   1=1 ";
            str = str + " and a.maphu in (" + sMaDoiTuong + ") " + str4;
            if (sMaBacSi != "")
            {
                str = str + " and a.mabs='" + sMaBacSi + "'";
            }
            if (sMakp != "")
            {
                str = str + " and a.makp in (" + sMakp.Substring(0, sMakp.Length - 1) + ")";
            }
            if (sSoTheBHYT != "")
            {
                object obj2 = str;
                str = string.Concat(new object[] { obj2, " and substr(bh.sothe,", iViTriTheBhyt, ",", sSoTheBHYT.Trim().Length, ")='", sSoTheBHYT.Trim(), "'" });
            }
            if (bToaThuocNgoaiTru)
            {
                str = str + " and a.loaiba not in (2)";
            }
            if (sLocMaTheBHYT != "")
            {
                string str10 = str;
                str = str10 + " and  (substr(upper(bh.sothe)," + str9 + ") in (" + sLocMaTheBHYT + ") or substr(upper(bh.sothe),1,2) in (" + sLocMaTheBHYT + ")) ";
            }
            if (iNhomKho != -1)
            {
                str = str + " and a.nhom=" + iNhomKho;
            }
            if (sNhomBhyt != "")
            {
                str = str + " and e.idnhombhyt in " + sNhomBhyt.Substring(0, sNhomBhyt.Length - 1) + ") ";
            }
            if (sMaKho != "")
            {
                str = str + " and b.makho in (" + sMaKho.Substring(0, sMaKho.Length - 1) + ")";
            }
            str = str + " group by 0,a.maql,a.maphu,a.id,a.sobienlai,to_char(a.ngay,'dd/mm/yyyy'),a.traituyen,nvl(bh.sothe,' '),to_char(bh.tungay,'dd/mm/yyyy') ,to_char(bh.denngay,'dd/mm/yyyy') ,a.makp,bv.tenkp,a.mabn,bn.phai,f.hoten,f.namsinh,bn.sonha,bn.thon,bn3.tenpxa,bn2.tenquan,bn1.tentt,a.mabv,dv.nhombv,g.tenbv," + str6 + ",a.congkham,a.chandoan,a.maicd,a1.nguyennhan,a.mavaovien union all select 1 loai,a.maql,a.mavaovien,a.maphu madoituong,a.id,a.sobienlai,to_char(a.ngay,'dd/mm/yyyy') ngay,null vaora,a.traituyen,nvl(bh.sothe,' ') sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,a.makp,bv.tenkp,a.mabn,bn.phai,f.hoten,f.namsinh,bn.sonha||' '||bn.thon||','||bn3.tenpxa||','||bn2.tenquan||','||bn1.tentt diachi,decode(dv.nhombv,null,a.mabv,dv.nhombv) mabv,g.tenbv," + str6 + " nhomvp,d.id loaivp,a.congkham congkham,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a1.nguyennhan=0 then '1' else '' end tngt,sum(b.soluong*b.dongia) sotien ";
            str = str + " from " + user + "d" + d_mmyy + ".bhytkb a inner join " + user + "d" + d_mmyy + ".bhytcls b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id left join " + user + ".tainantt a1 on a.maql=a1.maql inner join " + user + ".v_loaivp d on c.id_loai=d.id inner join " + user + "d" + d_mmyy + ".bhytds f on a.mabn=f.mabn left join " + user + ".btdbn bn on f.mabn=bn.mabn left join  " + user + ".btdtt bn1 on bn1.matt=bn.matt left join  " + user + ".btdquan bn2 on bn2.maqu=bn.maqu left join  " + user + ".btdpxa bn3 on bn3.maphuongxa=bn.maphuongxa left join " + user + ".dmnoicapbhyt g on a.mabv=g.mabv inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id left join " + user + ".dmnhomdv_bhyt dv on a.mabv=dv.mabv left join " + user + ".btdkp_bv bv on a.makp=bv.makp left join (" + str2 + ") bh on a.maql=bh.maql";
            str = str + " where a.maphu in(" + sMaDoiTuong + ") " + str4;
            if (sMaBacSi != "")
            {
                str = str + " and a.mabs='" + sMaBacSi + "'";
            }
            if (sMakp != "")
            {
                str = str + " and a.makp in (" + sMakp.Substring(0, sMakp.Length - 1) + ")";
            }
            if (sSoTheBHYT != "")
            {
                object obj3 = str;
                str = string.Concat(new object[] { obj3, " and substr(bh.sothe,", iViTriTheBhyt, ",", sSoTheBHYT.Trim().Length, ")='", sSoTheBHYT.Trim(), "'" });
            }
            if (bToaThuocNgoaiTru)
            {
                str = str + " and a.loaiba not in (2)";
            }
            if (sLocMaTheBHYT != "")
            {
                string str11 = str;
                str = str11 + " and (substr(upper(bh.sothe)," + str9 + ") in (" + sLocMaTheBHYT + ") or substr(upper(bh.sothe),1,2) in (" + sLocMaTheBHYT + ")) ";
            }
            if (iNhomKho != -1)
            {
                str = str + " and a.nhom=" + iNhomKho;
            }
            if (sNhomBhyt != "")
            {
                str = str + " and e.idnhombhyt in (" + sNhomBhyt.Substring(0, sNhomBhyt.Length - 1) + ") ";
            }
            str = str + " group by 1,a.maql,a.mavaovien,a.maphu,a.id,a.sobienlai,to_char(a.ngay,'dd/mm/yyyy'),a.traituyen,nvl(bh.sothe,' '),to_char(bh.tungay,'dd/mm/yyyy') ,to_char(bh.denngay,'dd/mm/yyyy') ,a.makp,bv.tenkp,a.mabn,bn.phai,f.hoten,f.namsinh,a.mabv,bn.sonha,bn.thon,bn3.tenpxa,bn2.tenquan,bn1.tentt,dv.nhombv,g.tenbv," + str6 + ",d.id,a.congkham,a.chandoan,a.maicd,a1.nguyennhan";
            if (bLaySoLieuThuocTuTruc)
            {
                str = str + " union all ";
                if (bInRiengSoLieuThuocTuTruc)
                {
                    str = "";
                }
                string str12 = this.f_get_sql_BenhAnDT(2, d_mmyy);
                str = str + "select 3 loai,a.maql,f.mavaovien,a1.madoituong,f.id,null sobienlai,to_char(a.ngay,'dd/mm/yyyy') ngay,null vaora,f.traituyen,nvl(f.sothe,' ') sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,f.makp,kp.tenkp,f.mabn,bn.phai,d.hoten,d.namsinh,bn.sonha||' '||bn.thon||','||bn3.tenpxa||','||bn2.tenquan||','||bn1.tentt diachi,decode(dv.nhombv,null,f.mabv,dv.nhombv) mabv,g.tenbv," + str6 + " nhomvp,0 loaivp,sum(0) as congkham,f.chandoan,f.maicd,null maicdkt,null chandoankt,case when a2.nguyennhan=0 then '1' else '' end tngt,";
                if (this._iGiaBhyt == 2)
                {
                    str = str + " sum(a1.sotien) sotien ";
                }
                else
                {
                    str = str + "sum(a1.soluong*" + ((this._iGiaBhyt == 0) ? "a1.giamua" : "a1.giaban") + ") sotien ";
                }
                string str13 = str;
                str = str13 + " from " + user + "d" + d_mmyy + ".d_xuatsdll a inner join " + user + "d" + d_mmyy + ".d_thucxuat a1 on a.id=a1.id inner join " + user + ".d_dmbd b on a1.mabd=b.id inner join " + user + ".d_dmnhom c on b.manhom=c.id left join (select distinct mabn,maql,ngay from (" + str12 + ")) f2 on a.maql=f2.maql and f2.mabn=a.mabn inner join " + user + "d" + d_mmyy + ".bhytkb f on f2.maql=f.maql inner join " + user + "d" + d_mmyy + ".bhytds d on f.mabn=d.mabn left join " + user + ".btdbn bn on d.mabn=bn.mabn left join  " + user + ".btdtt bn1 on bn1.matt=bn.matt left join  " + user + ".btdquan bn2 on bn2.maqu=bn.maqu left join  " + user + ".btdpxa bn3 on bn3.maphuongxa=bn.maphuongxa inner join " + user + ".v_nhomvp e on c.nhomvp=e.ma inner join " + user + ".v_nhombhyt h on e.idnhombhyt=h.id  left join " + user + ".tainantt a2 on f.maql=a2.maql  left join " + user + ".dmnoicapbhyt g on f.mabv=g.mabv inner join " + user + ".btdkp_bv kp on f.makp=kp.makp left join " + user + ".dmnhomdv_bhyt dv on f.mabv=dv.mabv left join (" + str3 + ") bh on f.maql=bh.maql";
                str = str + " where  a.mabn=f.mabn and a.loai=2  and a1.madoituong in(" + sMaDoiTuong + ") and to_date(a.ngay,'dd/mm/yy')  between to_date('" + sTuNgay + "','dd/mm/yy')  and to_date('" + sDenNgay + "','dd/mm/yy')";
                if (sMaBacSi != "")
                {
                    str = str + " and f.mabs='" + sMaBacSi + "'";
                }
                if (sMakp != "")
                {
                    str = str + " and f.makp in (" + sMakp.Substring(0, sMakp.Length - 1) + ")";
                }
                if (sSoTheBHYT != "")
                {
                    object obj4 = str;
                    str = string.Concat(new object[] { obj4, " and substr(f.sothe,", iViTriTheBhyt, ",", sSoTheBHYT.Trim().Length, ")='", sSoTheBHYT.Trim(), "'" });
                }
                if (bToaThuocNgoaiTru)
                {
                    str = str + " and f.loaiba not in (2)";
                }
                if (sLocMaTheBHYT != "")
                {
                    string str14 = str;
                    str = str14 + " and  (substr(upper(f.sothe)," + str9 + ") in (" + sLocMaTheBHYT + ")or substr(upper(f.sothe),1,2) in (" + sLocMaTheBHYT + ")) ";
                }
                if (iNhomKho != -1)
                {
                    str = str + " and f.nhom=" + iNhomKho;
                }
                if (sNhomBhyt != "")
                {
                    str = str + " and e.idnhombhyt in (" + sNhomBhyt.Trim(new char[] { ',' }) + ") ";
                }
                if (sMaKho != "")
                {
                    str = str + " and a1.makho in (" + sMaKho.Trim(new char[] { ',' }) + ")";
                }
                str = str + " group by 3,a.maql,f.mavaovien,a1.madoituong,f.id,to_char(a.ngay,'dd/mm/yyyy'),f.traituyen,nvl(f.sothe,' '),to_char(bh.tungay,'dd/mm/yyyy') ,to_char(bh.denngay,'dd/mm/yyyy') ,f.makp,kp.tenkp,f.mabn,bn.phai,d.hoten,d.namsinh,bn.sonha,bn.thon,bn3.tenpxa,bn2.tenquan,bn1.tentt,f.mabv,dv.nhombv,g.tenbv,f.chandoan,f.maicd,a2.nguyennhan," + str6 + "";
            }
            if (!bLaySoLieuVPNgoaiTru)
            {
                return str;
            }
            if (this._lib.updloaibn == 1)
            {
                string sql = "";
                sql = " update " + str5 + ".v_ttrvll  set loaibn=2 where loaibn<>2 and  id in (select b.id from " + user + ".benhandt a ," + str5 + ".v_ttrvds b where a.mabn=b.mabn and a.maql=b.maql and a.loaiba=2) ";
                this._lib.execute_data(sql);
                sql = " update " + str5 + ".v_ttrvll  set loaibn=1 where loaibn<>1 and  id in (select b.id from " + user + ".benhandt a ," + str5 + ".v_ttrvds b," + user + ".btdkp_bv c where a.mabn=b.mabn and a.maql=b.maql and a.makp=c.makp and a.loaiba=1 and c.loai=0 )";
                this._lib.execute_data(sql);
            }
            string str16 = "select e.id from " + str5 + ".v_hoantra f inner join " + str5 + ".v_ttrvll e on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai inner join " + str5 + ".v_ttrvds g on e.id=g.id WHERE  g.mabn=f.mabn  and e.ngay between to_date('" + sTuNgay + "','dd/mm/yyyy')  and to_date('" + sDenNgay + "','dd/mm/yyyy')";
            string str17 = "select a.id,a.sothe ,decode(dv.nhombv,null,a.mabv,dv.nhombv) mabv,b.tenbv,a.maphu,a.traituyen from " + str5 + ".v_ttrvbhyt a inner join " + user + ".dmnoicapbhyt b on a.mabv=b.mabv left join " + user + ".dmnhomdv_bhyt dv on a.mabv=dv.mabv ";
            str = str + " union all ";
            if (bInRiengSoLieuVPNgoaiTru)
            {
                str = "";
            }
            str = str + " select 4 loai,a.maql,badt.mavaovien,d.madoituong,a.id,f.sobienlai,to_char(f.ngay,'dd/mm/yyyy') ngay,to_char(a.ngayvao,'dd/mm/yyyy')||' '||to_char(a.ngayra,'dd/mm/yyyy') vaora,g.traituyen,nvl(g.sothe,' ') sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,f.makp,bv.tenkp,a.mabn,bn.phai,bn.hoten,bn.namsinh,bn.sonha||' '||bn.thon||','||bn3.tenpxa||','||bn2.tenquan||','||bn1.tentt diachi,g.mabv,g.tenbv,e.id nhomvp,0 loaivp,sum(0) congkham,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a1.nguyennhan=0 then '1' else '' end tngt,sum(d.soluong*d.dongia) sotien ";
            str = str + " from " + str5 + ".v_ttrvds a  inner join " + user + ".btdbn bn on a.mabn=bn.mabn left join  " + user + ".btdtt bn1 on bn1.matt=bn.matt left join  " + user + ".btdquan bn2 on bn2.maqu=bn.maqu left join  " + user + ".btdpxa bn3 on bn3.maphuongxa=bn.maphuongxa inner join " + str5 + ".v_ttrvct d on a.id=d.id  left join " + user + ".tainantt a1 on a1.maql=a.maql inner join (select " + str7 + " id,c.id mavp from " + user + ".d_dmnhom a," + user + ".v_nhomvp b," + user + ".d_dmbd c," + user + ".v_nhombhyt d  where a.nhomvp=b.ma   and a.id=c.manhom and b.idnhombhyt=d.id) e on d.mavp=e.mavp  inner join " + str5 + ".v_ttrvll f on a.id=f.id  left join (" + str17 + ") g on a.id=g.id left join " + user + ".btdkp_bv bv on f.makp=bv.makp left join (" + str3 + ") bh on a.maql=bh.maql left join (" + this.f_get_sql_BenhAnDT(2, d_mmyy) + ") badt on badt.maql=a.maql where 1=1  ";
            str = str + " and a.id not in (" + str16 + ")   and f.loaibn = 2 and d.madoituong in(" + sMaDoiTuong + ")";
            str = str + " and to_date(f.ngay,'dd/mm/yy') between to_date('" + sTuNgay + "','dd/mm/yy') and to_date('" + sDenNgay + "','dd/mm/yy')";
            if (sMakp != "")
            {
                str = str + " and f.makp in (" + sMakp.Substring(0, sMakp.Length - 1) + ")";
            }
            if (sLocMaTheBHYT != "")
            {
                string str18 = str;
                str = str18 + " and  (substr(upper(g.sothe)," + str9 + ") in (" + sLocMaTheBHYT + ") or substr(upper(g.sothe),1,2) in (" + sLocMaTheBHYT + ")) ";
            }
            string str19 = str + " group by 4,a.maql,d.madoituong,a.id,f.sobienlai,to_char(f.ngay,'dd/mm/yyyy'),to_char(a.ngayvao,'dd/mm/yyyy'),to_char(a.ngayra,'dd/mm/yyyy'),g.traituyen,nvl(g.sothe,' '),to_char(bh.tungay,'dd/mm/yyyy') ,to_char(bh.denngay,'dd/mm/yyyy') ,f.makp,bv.tenkp,a.mabn,bn.phai,bn.hoten,bn.namsinh,bn.sonha,bn.thon,bn3.tenpxa,bn2.tenquan,bn1.tentt,g.mabv,g.tenbv,e.id,a.chandoan,a.maicd,a1.nguyennhan union all  select 4 loai,a.maql,d.madoituong,a.id,f.sobienlai,to_char(f.ngay,'dd/mm/yyyy') ngay,to_char(a.ngayvao,'dd/mm/yyyy')||' '||to_char(a.ngayra,'dd/mm/yyyy') vaora,g.traituyen,nvl(g.sothe,' ') sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,f.makp,bv.tenkp,a.mabn,bn.phai,bn.hoten,bn.namsinh,bn.sonha||' '||bn.thon||','||bn3.tenpxa||','||bn2.tenquan||','||bn1.tentt diachi, g.mabv,g.tenbv,e.id nhomvp,e.loaivp,sum(0) congkham,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a1.nguyennhan=0 then '1' else '' end tngt,sum(d.soluong*d.dongia) sotien  ";
            string str20 = str19 + " from " + str5 + ".v_ttrvds a left join " + user + ".tainantt a1 on a1.maql=a.maql inner join " + user + ".btdbn bn on a.mabn=bn.mabn left join  " + user + ".btdtt bn1 on bn1.matt=bn.matt left join  " + user + ".btdquan bn2 on bn2.maqu=bn.maqu left join  " + user + ".btdpxa bn3 on bn3.maphuongxa=bn.maphuongxa inner join " + str5 + ".v_ttrvct d on a.id=d.id inner join (select " + str7 + " id,c.id mavp,a.id loaivp from " + user + ".v_loaivp a," + user + ".v_nhomvp b," + user + ".v_giavp c," + user + ".v_nhombhyt d  where a.id=c.id_loai   and a.id_nhom=b.ma and b.idnhombhyt=d.id) e on d.mavp=e.mavp  inner join " + str5 + ".v_ttrvll f on a.id=f.id  left join (" + str17 + ") g on a.id=g.id left join " + user + ".btdkp_bv bv on f.makp=bv.makp left join (" + str3 + ") bh on a.maql=bh.maql ";
            str = str20 + " where  f.loaibn = 2  and a.id not in (" + str16 + ")   and d.madoituong in(" + sMaDoiTuong + ") and to_date(f.ngay,'dd/mm/yy')  between to_date('" + sTuNgay + "','dd/mm/yy')  and to_date('" + sDenNgay + "','dd/mm/yy')";
            if (sMakp != "")
            {
                str = str + " and f.makp in (" + sMakp.Substring(0, sMakp.Length - 1) + ")";
            }
            if (sLocMaTheBHYT != "")
            {
                string str21 = str;
                str = str21 + " and ( substr(upper(g.sothe)," + str9 + ") in (" + sLocMaTheBHYT + ") or substr(upper(g.sothe),1,2) in (" + sLocMaTheBHYT + "))";
            }
            return (str + " group by 4,a.maql,badt.mavaovien,d.madoituong,a.id,f.sobienlai,to_char(f.ngay,'dd/mm/yyyy'),to_char(a.ngayvao,'dd/mm/yyyy'),to_char(a.ngayra,'dd/mm/yyyy'),g.traituyen,nvl(g.sothe,' '),to_char(bh.tungay,'dd/mm/yyyy') ,to_char(bh.denngay,'dd/mm/yyyy') ,f.makp,bv.tenkp,a.mabn,bn.phai,bn.hoten,bn.namsinh,bn.sonha,bn.thon,bn3.tenpxa,bn2.tenquan,bn1.tentt,g.mabv,g.tenbv,e.id,e.loaivp,a.chandoan,a.maicd,a1.nguyennhan");
        }

        public string f_NgoaiTru_GetSql_CLS(string mmyy, string tungay, string denngay, string madoituong, string makp, string nhomvp_bhyt_medi, bool LaySLVienPhi, bool LaySLNgoaiTruTheoDotDT, bool LaySLThuocTuTruc, bool LaySLBenhAnNgoaiTru, bool InRiengNgoaiTru, bool LaySLVPKhoa, string mavienphi, string userid, string mabv, string notmabv, bool LaySLTheoBenhNhan, bool LayTheoTT2348, string bc_mabn, string bc_quyenso)
        {
            StringBuilder builder = new StringBuilder();
            if (bc_mabn != "")
            {
                bc_mabn = "'" + bc_mabn.Trim(new char[] { ',' }).Replace(",", "','") + "'";
            }
            string str = "3";
            int num = this._lib.iMavp_congkham(1);
            string str2 = this.f_get_sql_BenhAnDT(2, mmyy);
            string user = this._lib.user;
            string str4 = user + "d" + mmyy;
            string str5 = user + mmyy;
            string str6 = "select e.id from " + str5 + ".v_hoantra f," + str5 + ".v_ttrvll e," + str5 + ".v_ttrvds g where e.id=g.id and e.quyenso=f.quyenso(+) and e.sobienlai=f.sobienlai(+) and g.mabn=f.mabn and e.ngay between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')";
            string str7 = mavienphi.Trim(new char[] { ',' }).Replace(",", " or c.id=");
            str7 = " and (c.id=" + str7 + ")";
            if (mavienphi == "")
            {
                str7 = "";
            }
            builder.Append(" select a.loaiba,c.manhom,bh.idnhombhytmedisoft as nhombhyt, 0 as stt,soft.ten as tennhombhyt,0 as loai,a.congkham as congkham, d.ten as tennhom, c.maloai,  o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt, bb.nhomcc, n.ten as nhacc, c.manuoc, m.ten nuocsx, c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c1.mabv,d.thuocyhct,");
            builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,bb.giamua as dongia ");
            if (this._iGiaBhyt == 3)
            {
                builder.Append(", sum(b.soluong*b.gia_bh) sotien ");
            }
            else if (this._iGiaBhyt == 2)
            {
                builder.Append(", sum(b.sotien) sotien ");
            }
            else
            {
                builder.Append(", sum(b.soluong*bb.giamua)as sotien ");
            }
            builder.Append(",0 as bhyttra,c.ma as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bvbh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
            builder.Append(" from " + str4 + ".bhytkb a  ");
            builder.Append(" inner join " + str4 + ".bhytthuoc b on a.id=b.id inner join " + str4 + ".d_theodoi bb on b.sttt=bb.id ");
            builder.Append(" inner join " + user + ".d_dmbd c  on b.mabd=c.id inner join " + user + ".d_dmloai o on c.maloai=o.id inner join " + user + ".d_dmnhom d on c.manhom=d.id ");
            builder.Append(" left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  ");
            builder.Append(" left join " + user + ".tenvien c1 on c.mabv=c1.mabv  ");
            builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
            builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id ");
            builder.Append(" left join " + user + ".d_dmhang h on c.mahang=h.id ");
            builder.Append(" left join " + user + ".d_dmnx n on bb.nhomcc=n.id ");
            builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
            builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
            builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
            builder.Append(" left join (" + this.f_get_sql_BenhAnDT(3, mmyy) + ") bvbh on bvbh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + str5 + ".d_thuocbhytll b2 on b2.id=b.id") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
            builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
            if (bc_mabn != "")
            {
                builder.Append(" and a.mabn in (" + bc_mabn + ")");
            }
            if (madoituong.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.maphu in (" + madoituong.Trim(new char[] { ',' }) + ")");
            }
            if (mabv != "")
            {
                builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (notmabv != "")
            {
                builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (LaySLNgoaiTruTheoDotDT)
            {
                builder.Append(" and a.loaiba not in (2)");
            }
            if (makp.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLVienPhi)
            {
                builder.Append(" and a.quyenso=0 and a.sobienlai=0");
            }
            builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
            if (userid.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
            }
            if (bc_quyenso != "")
            {
                builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
            }
            builder.Append(" group by a.ngay,a.congkham,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat, ");
            builder.Append("c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c1.mabv,d.thuocyhct,");
            builder.Append(" a.loaiba,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, c.ma, c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt, bb.nhomcc, n.ten,bb.giamua " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,a.maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bvbh.malk,a.mavaovien" : ""));
            if (this._iGiaBhyt == 3)
            {
                builder.Append(", b.gia_bh ");
            }
            else if (this._iGiaBhyt == 2)
            {
                builder.Append("");
            }
            else
            {
                builder.Append(", bb.giamua ");
            }
            builder.Append((builder.ToString() == "") ? "" : " union all ");
            builder.Append(" select a.loaiba, d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,0 as loai,a.congkham as congkham, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat, ");
            builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
            builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,b.dongia as dongia, ");
            builder.Append(" sum(b.soluong*b.dongia)as sotien ");
            builder.Append(",0 as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bvbh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
            builder.Append(" from " + str4 + ".bhytkb a ");
            builder.Append(" inner join " + str4 + ".bhytcls b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id");
            builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id ");
            builder.Append(" inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
            builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
            builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
            builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
            builder.Append(" left join (" + this.f_get_sql_BenhAnDT(3, mmyy) + ") bvbh on bvbh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + str5 + ".v_chidinh b2 on b2.id=b.idchidinh") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
            builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
            if (bc_mabn != "")
            {
                builder.Append(" and a.mabn in (" + bc_mabn + ")");
            }
            if (madoituong.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.maphu in (" + madoituong.Trim(new char[] { ',' }) + ")");
            }
            if (mabv != "")
            {
                builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (notmabv != "")
            {
                builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (LaySLNgoaiTruTheoDotDT)
            {
                builder.Append(" and a.loaiba not in (2)");
            }
            if (makp.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLVienPhi)
            {
                builder.Append(" and a.quyenso=0 and a.sobienlai=0");
            }
            if (userid.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
            }
            if (bc_quyenso != "")
            {
                builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
            }
            builder.Append(" and bh.idnhombhytmedisoft not in(9," + str + ")" + str7 + "");
            builder.Append(" group by a.ngay,a.congkham,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,a.loaiba, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,a.maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bvbh.malk,a.mavaovien" : ""));
            builder.Append((builder.ToString() == "") ? "" : " union all ");
            builder.Append(" select a.loaiba, d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,0 as loai,a.congkham as congkham, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat, ");
            builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
            builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,t_kham as dongia, ");
            builder.Append(" sum(b.soluong*t_kham)as sotien ");
            builder.Append(",0 as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bvbh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
            builder.Append(" from " + str4 + ".bhytkb a ");
            builder.Append(" inner join (select 1 as soluong, mavp,makp from btdkp_bv) b on b.makp=a.makp");
            builder.Append(" inner join " + user + ".v_giavp c on b.mavp=c.id");
            builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id ");
            builder.Append(" inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
            builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
            builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
            builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
            builder.Append(" left join (" + this.f_get_sql_BenhAnDT(3, mmyy) + ") bvbh on bvbh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on a.mabs=b21.ma") : ""));
            builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
            if (bc_mabn != "")
            {
                builder.Append(" and a.mabn in (" + bc_mabn + ")");
            }
            if (madoituong.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.maphu in (" + madoituong.Trim(new char[] { ',' }) + ")");
            }
            if (mabv != "")
            {
                builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (notmabv != "")
            {
                builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (LaySLNgoaiTruTheoDotDT)
            {
                builder.Append(" and a.loaiba not in (2)");
            }
            if (makp.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLVienPhi)
            {
                builder.Append(" and a.quyenso=0 and a.sobienlai=0");
            }
            if (userid.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
            }
            if (bc_quyenso != "")
            {
                builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
            }
            builder.Append(str7);
            builder.Append(" group by a.ngay,a.congkham,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,a.loaiba, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, a.t_kham, c.kythuat " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,a.maicd,to_char(a.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bvbh.malk,a.mavaovien" : ""));
            if (LaySLThuocTuTruc)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                builder.Append(" select b.loaiba, c.manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt ,0 as loai,sum(0) as congkham, d.ten as tennhom, c.maloai,  o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, c.mahang, h.ten hangsx, c.kythuat, ");
                builder.Append("c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c1.mabv,d.thuocyhct,");
                builder.Append("to_char(b.ngay,'dd/mm/yyyy') ngayduyet,  sum(a1.soluong) as soluong,bb.giamua as dongia, ");
                builder.Append(" sum(a1.soluong*bb.giamua) sotien ");
                builder.Append(", 0 as bhyttra ,c.ma as mavp1,b.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",b.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,b.maicd as maicd,to_char(a.ngayylenh,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,b.maql as maql,bvbh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append(" from " + str4 + ".bhytkb b  left join (select distinct mabn,maql,ngay from (" + str2 + ")) b2 on b.maql=b2.maql and b.mabn=b2.mabn inner  join " + str4 + ".d_xuatsdll a on b2.maql=a.maql");
                builder.Append(" inner join " + str4 + ".d_thucxuat a1 on a.id=a1.id ");
                builder.Append(" inner join " + str4 + ".d_theodoi bb on  a1.sttt=bb.id ");
                builder.Append(" inner join " + user + ".d_dmbd c on a1.mabd=c.id ");
                builder.Append(" inner join " + user + ".d_dmloai o on c.maloai=o.id");
                builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id ");
                builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                builder.Append(" left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  ");
                builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id ");
                builder.Append(" left join " + user + ".d_dmhang h on c.mahang=h.id ");
                builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on b.makp=bv.makp ");
                builder.Append(" left join (" + this.f_get_sql_BenhAnDT(3, mmyy) + ") bvbh on bvbh.maql=b.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b21.ma=b.mabs") : ""));
                builder.Append(" where to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and b.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (LaySLNgoaiTruTheoDotDT)
                {
                    builder.Append(" and b.loaiba not in (2)");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (LaySLVienPhi)
                {
                    builder.Append(" and b.quyenso=0 and b.sobienlai=0 ");
                }
                builder.Append(" and a.loai=2 ");
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and b.quyenso in (" + bc_quyenso + ")");
                }
                builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
                builder.Append(" group by b.ngay,b.loaiba,c.ma,bh.idnhombhytmedisoft,soft.ten ,b.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id,  c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt ");
                builder.Append(", bb.giamua, c.manuoc, m.ten , c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c1.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",b.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,b.maicd,to_char(a.ngayylenh,'dd/mm/yyyy hh24:mi'),bv.viettat,b.maql,bvbh.malk,a.mavaovien" : ""));
            }
            if (LaySLBenhAnNgoaiTru)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                if (LaySLVPKhoa)
                {
                    if (InRiengNgoaiTru)
                    {
                        builder = new StringBuilder();
                    }
                    builder.Append(" select xv.loaiba as loaiba, c.manhom,bh.idnhombhytmedisoft as nhombhyt, 0 as stt,soft.ten as tennhombhyt,1 as loai,sum(0) as congkham,d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                    builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                    builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c1.mabv,d.thuocyhct,to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,  sum(a.soluong) as soluong, a.dongia, ");
                    builder.Append(" sum(a.soluong*a.dongia) sotien , sum(a.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                    builder.Append(" from " + str5 + ".v_thvpll ab inner join " + str5 + ".v_thvpct a on ab.id=a.id ");
                    builder.Append(" inner join " + str5 + ".v_thvpbhyt a1 on a.id=a1.id  ");
                    builder.Append(" inner join " + user + ".d_dmbd c on a.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id ");
                    builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma ");
                    builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                    builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                    builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft = soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                    builder.Append(" left join (" + this.f_get_sql_theBHYT(3, mmyy) + ") bh on bh.maql=ab.maql ");
                    builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and xv.loaiba not in(1) and a1.sothe is not null  ");
                    }
                    else
                    {
                        builder.Append(" and xv.loaiba in(2)  and a1.sothe is not null ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
                    builder.Append(" group by ab.ngay,xv.loaiba,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                    builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c1.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",a.mabn" : ""));
                    builder.Append(" union all ");
                    builder.Append(" select  xv.loaiba as loaiba, d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,1 as loai,sum(0) as congkham, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                    builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                    builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                    builder.Append(" to_char(ab.ngay,'dd/mm/yyyy') ngayduyet, sum(a.soluong) as soluong, a.dongia as dongia,  sum(a.soluong*a.dongia)as sotien , sum(a.bhyttra)as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                    builder.Append("  from " + str5 + ".v_thvpll ab inner join " + str5 + ".v_thvpct a on ab.id=a.id  ");
                    builder.Append("  inner join " + str5 + ".v_thvpbhyt b on a.id=b.id inner join " + user + ".v_giavp c on a.mavp=c.id ");
                    builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                    builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                    builder.Append(" left join (" + this.f_get_sql_theBHYT(2, mmyy) + ") bh on ab.maql=bh.maql ");
                    builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and xv.loaiba not in(1) and b.sothe is not null   ");
                    }
                    else
                    {
                        builder.Append(" and xv.loaiba in(2) and b.sothe is not null ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
                    builder.Append(" group by ab.ngay,xv.loaiba,bh.idnhombhytmedisoft,soft.ten,c.ma,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, a.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",ab.mabn" : ""));
                }
                else
                {
                    if (InRiengNgoaiTru)
                    {
                        builder = new StringBuilder();
                    }
                    builder.Append(" select a.loaibn as loaiba, c.manhom,bh.idnhombhytmedisoft as nhombhyt, 0 as stt,soft.ten as tennhombhyt,1 as loai,sum(0) as congkham,d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                    builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                    builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c1.mabv,d.thuocyhct,to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(a1.soluong) as soluong, a1.dongia, ");
                    builder.Append(" sum(a1.soluong*a1.dongia) sotien , sum(a1.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bvbh.malk,b2.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append(" from " + str5 + ".v_ttrvds ab inner join " + str5 + ".v_ttrvll a on ab.id=a.id ");
                    builder.Append(" inner join " + str5 + ".v_ttrvct a1 on a.id=a1.id  ");
                    builder.Append(" inner join " + user + ".d_dmbd c on a1.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id ");
                    builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma ");
                    builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                    builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                    builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft = soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join (" + this.f_get_sql_BenhAnDT(2, mmyy) + ") bvbh on ab.maql=bvbh.maql " + (LayTheoTT2348 ? (" left join (" + this.f_get_sql_benhandt(2, mmyy) + ") b2 on b2.maql=bvbh.maql") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and a.loaibn not in(1) and ab.id not in (" + str6 + ") ");
                    }
                    else
                    {
                        builder.Append(" and a.loaibn in(2) and ab.id not in (" + str6 + ") ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    if (userid.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                    }
                    if (bc_quyenso != "")
                    {
                        builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                    }
                    builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
                    builder.Append(" group by a.ngay,a.loaibn,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                    builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a1.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c1.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,ab.maql,bvbh.malk,b2.mavaovien" : ""));
                    builder.Append(" union all ");
                    builder.Append(" select  a.loaibn as loaiba, d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,1 as loai,sum(0) as congkham, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                    builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                    builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,to_char(a.ngay,'dd/mm/yyyy') ngayduyet,");
                    builder.Append(" sum(b.soluong) as soluong,  b.dongia as dongia,  sum(b.soluong*b.dongia)as sotien , sum(b.bhyttra)as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bvbh.malk,b2.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append("  from " + str5 + ".v_ttrvds ab inner join " + str5 + ".v_ttrvll a on ab.id=a.id  ");
                    builder.Append("  inner join " + str5 + ".v_ttrvct b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id ");
                    builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                    builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(2, mmyy) + ")) bvbh on ab.maql=bvbh.maql " + (LayTheoTT2348 ? (" left join " + user + ".benhandt b2 on ab.maql=b2.maql") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and a.loaibn not in(1) and ab.id not in (" + str6 + ") ");
                    }
                    else
                    {
                        builder.Append(" and a.loaibn in(2) and ab.id not in (" + str6 + ") ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str7 + "");
                    if (userid.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                    }
                    if (bc_quyenso != "")
                    {
                        builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                    }
                    builder.Append(" group by a.ngay,a.loaibn,bh.idnhombhytmedisoft,soft.ten,c.ma,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,ab.maql,bvbh.malk,b2.mavaovien" : ""));
                }
            }
            if (LaySLVPKhoa)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                builder.Append(" select a.loaiba, d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,0 as loai,sum(0) as congkham, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat, ");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append(" to_char(a.ngay,'dd/mm/yyyy') ngayduyet, sum(b.soluong) as soluong,b.dongia as dongia, ");
                builder.Append(" sum(b.soluong*b.dongia)as sotien ");
                builder.Append(" ,0 as bhyttra,to_char(c.ma) as mavp1,b.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn"));
                builder.Append(" from " + str5 + ".benhandt a ");
                builder.Append(" inner join " + user + ".btdbn aa on a.mabn=aa.mabn ");
                builder.Append(" inner join " + str5 + ".v_vpkhoa b on a.maql=b.maql inner join " + user + ".v_giavp c on b.mavp=c.id ");
                builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id ");
                builder.Append(" inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on b.makp=bv.makp ");
                builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(3, mmyy) + ")) bvbh on bvbh.maql=a.maql ");
                builder.Append(" where to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and a.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" and a.loaiba in(3,4)");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and bh.idnhombhytmedisoft not in (" + str + ")" + str7 + "");
                builder.Append(" group by a.ngay,a.loaiba,c.ma,bh.idnhombhytmedisoft,soft.ten,b.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai,  d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt, c.ma, c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",a.mabn" : ""));
            }
            return builder.ToString();
        }

        public string f_NgoaiTru_getSql_Thuoc(string mmyy, string tungay, string denngay, string madoituong, string makp, string nhomvp_bhyt_medi, bool LaySLVienPhi, bool LaySLNgoaiTruTheoDotDT, bool LaySLThuocTuTruc, bool LaySLBenhAnNgoaiTru, bool InRiengNgoaiTru, bool LaySLVPKhoa, string userid, string mabv, string notmabv, bool LaySLTheoBenhNhan, bool LayTheoTT2348, string bc_mabn, string bc_quyenso)
        {
            StringBuilder builder = new StringBuilder();
            if (bc_mabn != "")
            {
                bc_mabn = "'" + bc_mabn.Trim(new char[] { ',' }).Replace(",", "','") + "'";
            }
            string user = this._lib.user;
            string str2 = user + "d" + mmyy;
            string str3 = user + mmyy;
            string str4 = "select e.id from " + str3 + ".v_hoantra f," + str3 + ".v_ttrvll e," + str3 + ".v_ttrvds g where e.id=g.id and e.quyenso=f.quyenso(+) and e.sobienlai=f.sobienlai(+) and g.mabn=f.mabn and e.ngay between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')";
            string str5 = " and nhombh.idnhombhytmedisoft in(" + nhomvp_bhyt_medi + ")";
            string str6 = this.f_get_sql_BenhAnDT(2, mmyy);
            builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,a.loaiba,c.manhom, d.ten as tennhom, c.maloai,  o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt, bb.nhomcc, n.ten as nhacc, c.manuoc, m.ten nuocsx, c.mahang, h.ten hangsx, c.kythuat ,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,");
            builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,round(bb.giamua,2) as dongia ");
            if (this._iGiaBhyt == 3)
            {
                builder.Append(", sum(b.soluong*b.gia_bh) sotien ");
            }
            else if (this._iGiaBhyt == 2)
            {
                builder.Append(", sum(b.sotien) sotien ");
            }
            else
            {
                builder.Append(", sum(b.soluong*bb.giamua)as sotien ");
            }
            builder.Append(",0 as bhyttra ,c.ma as mavp1,a.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
            builder.Append(" from " + str2 + ".bhytkb a  ");
            builder.Append(" inner join " + str2 + ".bhytthuoc b on a.id=b.id inner join " + str2 + ".d_theodoi bb on b.sttt=bb.id ");
            builder.Append(" inner join " + user + ".d_dmbd c  on b.mabd=c.id inner join " + user + ".d_dmloai o on c.maloai=o.id inner join " + user + ".d_dmnhom d on c.manhom=d.id ");
            builder.Append(" left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  ");
            builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
            builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
            builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id ");
            builder.Append(" left join " + user + ".d_dmhang h on c.mahang=h.id ");
            builder.Append(" left join " + user + ".d_dmnx n on bb.nhomcc=n.id ");
            builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
            builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
            builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
            builder.Append(" left join (" + this.f_get_sql_BenhAnDT(3, mmyy) + ") bh on bh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + str3 + ".d_thuocbhytll b2 on b2.maql=a.maql") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
            builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
            if (bc_mabn != "")
            {
                builder.Append(" and a.mabn in (" + bc_mabn + ")");
            }
            if (madoituong.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.maphu in (" + madoituong.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLNgoaiTruTheoDotDT)
            {
                builder.Append(" and a.loaiba not in (2)");
            }
            if (makp.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLVienPhi)
            {
                builder.Append(" and a.quyenso=0 and a.sobienlai=0");
            }
            if (userid.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
            }
            if (bc_quyenso != "")
            {
                builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
            }
            builder.Append(str5);
            if (mabv != "")
            {
                builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (notmabv != "")
            {
                builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            builder.Append(" group by a.ngay,nhombh.idnhombhytmedisoft,soft.ten,c.manuoc,c.ma ,a.makp,bv.tenkp, m.ten, c.mahang, h.ten, c.kythuat, c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct, a.loaiba,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, c.ma, c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt, bb.nhomcc, n.ten,bb.giamua " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,a.maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bh.malk,a.mavaovien" : ""));
            if (this._iGiaBhyt == 3)
            {
                builder.Append(", b.gia_bh ");
            }
            else if (this._iGiaBhyt == 2)
            {
                builder.Append("");
            }
            else
            {
                builder.Append(", bb.giamua ");
            }
            builder.Append((builder.ToString() == "") ? "" : " union all ");
            builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,a.loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat, null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
            builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,b.dongia as dongia, ");
            builder.Append(" sum(b.soluong*b.dongia)as sotien ");
            builder.Append(",0 as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp ,null mabv" + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
            builder.Append(" from " + str2 + ".bhytkb a ");
            builder.Append(" inner join " + str2 + ".bhytcls b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id");
            builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id ");
            builder.Append(" inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
            builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
            builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
            builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
            builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(3, mmyy) + ")) bh on bh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + str3 + ".v_chidinh b2 on b.idchidinh=b2.id") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b2.mabs=b21.ma") : ""));
            builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
            if (bc_mabn != "")
            {
                builder.Append(" and a.mabn in (" + bc_mabn + ")");
            }
            if (madoituong.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.maphu in (" + madoituong.Trim(new char[] { ',' }) + ")");
            }
            if (mabv != "")
            {
                builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (notmabv != "")
            {
                builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
            }
            if (LaySLNgoaiTruTheoDotDT)
            {
                builder.Append(" and a.loaiba not in (2)");
            }
            if (makp.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
            }
            if (LaySLVienPhi)
            {
                builder.Append(" and a.quyenso=0 and a.sobienlai=0");
            }
            if (userid.Trim(new char[] { ',' }) != "")
            {
                builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
            }
            if (bc_quyenso != "")
            {
                builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
            }
            builder.Append(str5);
            builder.Append(" group by a.ngay,a.loaiba,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,a.maicd,to_char(b2.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bh.malk,a.mavaovien" : ""));
            if (LaySLThuocTuTruc)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,b.loaiba, c.manhom, d.ten as tennhom, c.maloai,  o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, c.mahang, h.ten hangsx, c.kythuat, ");
                builder.Append("c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,");
                builder.Append("to_char(b.ngay,'dd/mm/yyyy') ngayduyet,  sum(a1.soluong) as soluong,round(bb.giamua,2) as dongia, ");
                builder.Append(" sum(a1.soluong*bb.giamua) sotien ");
                builder.Append(", 0 as bhyttra,c.ma as mavp1,b.makp,bv.tenkp,null mabv " + (LaySLTheoBenhNhan ? ",b.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,b.maicd as maicd,to_char(a.ngayylenh,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,b.maql as maql,bvbh.malk,b.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append(" from " + str2 + ".bhytkb b  inner join (select distinct mabn,maql,ngay from (" + str6 + ")) b2 on b.maql=b2.maql and b.mabn=b2.mabn inner  join " + str2 + ".d_xuatsdll a on b2.maql=a.maql and a.mabn=b2.mabn");
                builder.Append(" inner join " + str2 + ".d_thucxuat a1 on a.id=a1.id ");
                builder.Append(" inner join " + str2 + ".d_theodoi bb on  a1.sttt=bb.id ");
                builder.Append(" inner join " + user + ".d_dmbd c on a1.mabd=c.id ");
                builder.Append(" inner join " + user + ".d_dmloai o on c.maloai=o.id");
                builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id ");
                builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                builder.Append(" left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  ");
                builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id ");
                builder.Append(" left join " + user + ".d_dmhang h on c.mahang=h.id ");
                builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on b.makp=bv.makp ");
                builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(3, mmyy) + ")) bvbh on bvbh.maql=b.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b.mabs=b21.ma") : ""));
                builder.Append(" where to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and b.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (LaySLNgoaiTruTheoDotDT)
                {
                    builder.Append(" and b.loaiba not in (2)");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (LaySLVienPhi)
                {
                    builder.Append(" and b.quyenso=0 and b.sobienlai=0");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and b.quyenso in (" + bc_quyenso + ")");
                }
                builder.Append(" and a.loai= 2 ");
                builder.Append(str5);
                builder.Append(" group by b.ngay,b.loaiba,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,b.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id,  c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt ");
                builder.Append(", bb.giamua, c.manuoc, m.ten , c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",b.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,b.maicd,to_char(a.ngayylenh,'dd/mm/yyyy hh24:mi'),bv.viettat,b.maql,bvbh.malk,b.mavaovien" : ""));
            }
            if (LaySLBenhAnNgoaiTru)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                if (LaySLVPKhoa)
                {
                    if (InRiengNgoaiTru)
                    {
                        builder = new StringBuilder();
                    }
                    builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,xv.loaiba as loaiba, c.manhom, d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                    builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                    builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,   sum(a.soluong) as soluong,round(a.dongia,2) as dongia, ");
                    builder.Append(" sum(a.soluong*a.dongia) sotien , sum(a.bhyttra)as bhyttra ,c.ma as mavp1,a.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,xv.maicd as maicd,to_char(xv.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,xv.maql as maql,bh.malk,xv.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append(" from " + str3 + ".v_thvpll ab inner join " + str3 + ".v_thvpct a on ab.id=a.id ");
                    builder.Append(" inner join " + str3 + ".v_thvpbhyt a1 on a.id=a1.id  ");
                    builder.Append(" inner join " + user + ".d_dmbd c on a.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id ");
                    builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma ");
                    builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                    builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                    builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                    builder.Append(" left join (" + this.f_get_sql_theBHYT(3, mmyy) + ") bh on bh.maql=ab.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on xv.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and xv.loaiba not in (1) and a1.sothe is not null  ");
                    }
                    else
                    {
                        builder.Append(" and xv.loaiba in(2) and a1.sothe is not null  ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    builder.Append(str5);
                    builder.Append(" group by ab.ngay,xv.loaiba,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                    builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat ,xv.maicd,to_char(xv.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,xv.maql,bh.malk,xv.mavaovien" : ""));
                    builder.Append(" union all ");
                    builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft,soft.ten,xv.loaiba as loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                    builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                    builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                    builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                    builder.Append("to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,   sum(a.soluong) as soluong,a.dongia as dongia,  sum(a.soluong*a.dongia)as sotien , sum(a.bhyttra)as bhyttra ,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,xv.maicd as maicd,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,xv.maql as maql,bh.malk,xv.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append("  from " + str3 + ".v_thvpll ab inner join " + str3 + ".v_thvpct a on ab.id=a.id  ");
                    builder.Append("  inner join " + str3 + ".v_thvpbhyt b on a.id=b.id inner join " + user + ".v_giavp c on a.mavp=c.id ");
                    builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                    builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                    builder.Append(" left join (" + this.f_get_sql_theBHYT(3, mmyy) + ") bh on bh.maql=ab.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on xv.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and xv.loaiba not in(1) and b.sothe is not null ");
                    }
                    else
                    {
                        builder.Append(" and xv.loaiba in(2) and b.sothe is not null ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    builder.Append(str5);
                    builder.Append(" group by ab.ngay,xv.loaiba,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, a.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,xv.maicd,to_char(a.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,xv.maql,bh.malk,xv.mavaovien" : ""));
                }
                else
                {
                    if (InRiengNgoaiTru)
                    {
                        builder = new StringBuilder();
                    }
                    builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,a.loaibn as loaiba, c.manhom, d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                    builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                    builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(a1.soluong) as soluong, round(a1.dongia,2) as dongia, ");
                    builder.Append(" sum(a1.soluong*a1.dongia) sotien , sum(a1.bhyttra)as bhyttra ,c.ma as mavp1,a.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,xv.maicd as maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,xv.maql as maql,bvbh.malk,xv.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append(" from " + str3 + ".v_ttrvds ab inner join " + str3 + ".v_ttrvll a on ab.id=a.id ");
                    builder.Append(" inner join " + str3 + ".v_ttrvct a1 on a.id=a1.id  ");
                    builder.Append(" inner join " + user + ".d_dmbd c on a1.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id ");
                    builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma ");
                    builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                    builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                    builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(2, mmyy) + ")) bvbh on bvbh.maql=ab.maql " + (LayTheoTT2348 ? (" left join " + user + ".benhandt xv on xv.maql=ab.maql") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on xv.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and a.loaibn not in (1) and ab.id not in (" + str4 + ") ");
                    }
                    else
                    {
                        builder.Append(" and a.loaibn in(2) and ab.id not in (" + str4 + ") ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    if (userid.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                    }
                    if (bc_quyenso != "")
                    {
                        builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                    }
                    builder.Append(str5);
                    builder.Append(" group by a.ngay,a.loaibn,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                    builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a1.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat ,xv.maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,xv.maql,bvbh.malk,xv.mavaovien" : ""));
                    builder.Append(" union all ");
                    builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft,soft.ten,a.loaibn as loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                    builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                    builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                    builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                    builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,  b.dongia as dongia, sum(b.soluong*b.dongia)as sotien , sum(b.bhyttra)as bhyttra ,to_char(c.ma) as mavp1,a.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,xv.maicd as maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,xv.maql as maql,bvbh.malk,xv.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                    builder.Append("  from " + str3 + ".v_ttrvds ab inner join " + str3 + ".v_ttrvll a on ab.id=a.id  ");
                    builder.Append("  inner join " + str3 + ".v_ttrvct b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id ");
                    builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                    builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                    builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                    builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                    builder.Append(" left join (select distinct maql,mabv,malk from (" + this.f_get_sql_theBHYT(2, mmyy) + ")) bvbh on bvbh.maql=ab.maql " + (LayTheoTT2348 ? (" left join " + user + ".benhandt xv on xv.maql=ab.maql") : "") + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on xv.mabs=b21.ma") : ""));
                    builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                    if (bc_mabn != "")
                    {
                        builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                    }
                    if (madoituong.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                    }
                    if (mabv != "")
                    {
                        builder.Append(" and bvbh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (notmabv != "")
                    {
                        builder.Append(" and bvbh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                    }
                    if (LaySLVienPhi)
                    {
                        builder.Append(" and a.loaibn not in (1) and ab.id not in (" + str4 + ") ");
                    }
                    else
                    {
                        builder.Append(" and a.loaibn in(2) and ab.id not in (" + str4 + ") ");
                    }
                    if (makp.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                    }
                    if (userid.Trim(new char[] { ',' }) != "")
                    {
                        builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                    }
                    if (bc_quyenso != "")
                    {
                        builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                    }
                    builder.Append(str5);
                    builder.Append(" group by a.ngay,a.loaibn,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,xv.maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,xv.maql,bvbh.malk,xv.mavaovien" : ""));
                }
            }
            if (LaySLVPKhoa)
            {
                builder.Append((builder.ToString() == "") ? "" : " union all ");
                builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft,soft.ten,a.loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat, ");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,  sum(b.soluong) as soluong,b.dongia as dongia, ");
                builder.Append(" sum(b.soluong*b.dongia)as sotien ");
                builder.Append(" ,0 as bhyttra,to_char(c.ma) as mavp1,b.makp,bv.tenkp,null mabv" + (LaySLTheoBenhNhan ? ",a.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,a.maicd as maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,a.maql as maql,bh.malk,a.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append(" from " + str3 + ".benhandt a ");
                builder.Append(" inner join " + user + ".btdbn aa on a.mabn=aa.mabn ");
                builder.Append(" inner join " + str3 + ".v_vpkhoa b on a.maql=b.maql inner join " + user + ".v_giavp c on b.mavp=c.id ");
                builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id ");
                builder.Append(" inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on b.makp=bv.makp ");
                builder.Append(" left join (" + this.f_get_sql_theBHYT(2, mmyy) + ") bh on bh.maql=a.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on a.mabs=b21.ma") : ""));
                builder.Append(" where to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and a.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(str5);
                builder.Append(" group by a.ngay,a.loaiba,c.ma,nhombh.idnhombhytmedisoft,soft.ten ,b.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai,  d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt, c.ma, c.ten, c.ten, c.dvt, b.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",a.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat ,a.maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,a.maql,bh.malk,a.mavaovien" : ""));
            }
            return builder.ToString();
        }

        public DataSet f_Ngoaitru_tao_dataset(System.Data.DataTable dtnhomvp)
        {
            DataSet set = new DataSet();
            set = this._lib.get_data("select 0 stt,null sothe1,null sothe2,null sothe3,0 id,null sobienlai,null quyenso,null mabn,null hoten,null ngaysinh,null phai,null sothe,null gtritu,null gtriden,null manoidk,null noidk,null ngay,null chandoan,null maicd,null ngayvao,null ngayra from dual");
            set.Clear();
            set.Tables[0].Columns.Add(new DataColumn("soluotkham", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("SOPHIEU", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("CONGKHAM", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("songay", typeof(string)));
            foreach (DataRow row in dtnhomvp.Select("true", "stt"))
            {
                set.Tables[0].Columns.Add(new DataColumn("ST_" + row["id"].ToString().Trim().Trim(new char[] { '_' }), typeof(decimal)));
            }
            foreach (DataRow row2 in this._lib.get_data("select id from v_nhombhyt_medisoft order by id").Tables[0].Rows)
            {
                try
                {
                    set.Tables[0].Columns.Add(new DataColumn("ST_" + row2["id"].ToString().Trim().Trim(new char[] { '_' }), typeof(decimal)));
                }
                catch
                {
                }
            }
            set.Tables[0].Columns.Add(new DataColumn("DONE", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("TONGCONG", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("BNTRA", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("BHYTTRA", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("MAKP", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TENKP", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TRAITUYEN", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("tainangt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TYLEBHYT", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("NHOM_DT_BHYT", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TEN_NHOMDT_BHYT", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("madk", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("chiphids", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("Lydo", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("benhkhac", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noidkkcb", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("namqt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("thangqt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("tungay", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("denngay", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("diachi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("chandoankt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("maicdkt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("madoituong", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("nguoiduyet", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("malk", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("loaiba", typeof(int)));
            set.Tables[0].Columns.Add(new DataColumn("makhuvuc", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mavaovien", typeof(long)));
            set.Tables[0].Columns.Add(new DataColumn("ngayvv", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngayrv", typeof(long)));
            set.Tables[0].Columns.Add(new DataColumn("userid_thuvp", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ten_thuvp", typeof(string)));
            return set;
        }

        public void f_Ngoaitru_xuatExcel_mau01_TT2348(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("ma_lk", typeof(string));
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
            dset.Tables[0].Columns.Add("ten_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ma_tai_nan", typeof(int));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("muc_huong", typeof(decimal));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_nguonkhac", typeof(decimal));
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal));
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_PTTT_QT", typeof(string));
            decimal d = 0M;
            decimal num2 = this._lib.themoi15_sotien();
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    DataRow row2 = dset.Tables[0].NewRow();
                    try
                    {
                        row2["ma_lk"] = this._lib.Mabv + row["malk"].ToString();
                    }
                    catch
                    {
                        row2["ma_lk"] = row["malk"].ToString();
                    }
                    row2["STT"] = d = decimal.op_Increment(d);
                    row2["ma_bn"] = row["mabn"].ToString();
                    row2["ho_ten"] = row["HOTEN"].ToString();
                    row2["ngay_sinh"] = row["ngaysinh"].ToString();
                    row2["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row2["dia_chi"] = row["diachi"].ToString();
                    try
                    {
                        row2["ma_the"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row2["ma_the"] = row["sothe"].ToString();
                    }
                    row2["ma_dkbd"] = row["MANOIDK"].ToString();
                    try
                    {
                        row2["gt_the_tu"] = row["gtritu"].ToString().Substring(6, 4) + row["gtritu"].ToString().Substring(3, 2) + row["gtritu"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_tu"] = "";
                    }
                    try
                    {
                        row2["gt_the_den"] = row["gtriden"].ToString().Substring(6, 4) + row["gtriden"].ToString().Substring(3, 2) + row["gtriden"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_den"] = "";
                    }
                    row2["ma_benh"] = row["MAICD"].ToString();
                    row2["ma_benhkhac"] = row["maicdkt"].ToString();
                    row2["ten_benh"] = row["chandoan"].ToString();
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
                    row2["ma_tai_nan"] = 0;
                    try
                    {
                        row2["ngay_vao"] = row["NGAYVAO"].ToString().Substring(6, 4) + row["NGAYVAO"].ToString().Substring(3, 2) + row["NGAYVAO"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_vao"] = "";
                    }
                    try
                    {
                        row2["ngay_ra"] = row["NGAYRA"].ToString().Substring(6, 4) + row["NGAYRA"].ToString().Substring(3, 2) + row["NGAYRA"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_ra"] = "";
                    }
                    row2["so_ngay_dtri"] = 1;
                    row2["ket_qua_dtri"] = 0;
                    row2["tinh_trang_rv"] = 0;
                    try
                    {
                        row2["muc_huong"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                    }
                    catch
                    {
                    }
                    row2["t_tongchi"] = row["tongcong"].ToString();
                    row2["t_bntt"] = row["bntra"].ToString();
                    row2["t_bhtt"] = row["bhyttra"].ToString();
                    row2["t_nguonkhac"] = 0;
                    row2["t_ngoaids"] = 0;
                    row2["nam_qt"] = denngay.Substring(6, 4);
                    row2["thang_qt"] = denngay.Substring(3, 2);
                    row2["ma_loaikcb"] = 1;
                    row2["ma_cskcb"] = this._lib.MABHXH;
                    row2["ma_khuvuc"] = row["makhuvuc"].ToString();
                    row2["ma_PTTT_QT"] = "";
                    dset.Tables[0].Rows.Add(row2);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, this._lib.Msg);
                }
            }
            dset.WriteXml("bang01_ngoai.xml", XmlWriteMode.WriteSchema);
            int num3 = 0;
            int num4 = 0;
            int num5 = 0;
            int i = 0;
            int num7 = 0;
            num3 = 3;
            num4 = 5;
            num5 = dset.Tables[0].Rows.Count + 5;
            i = dset.Tables[0].Columns.Count - 1;
            num7 = num5;
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb01");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num3; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num3 + 8) + num4.ToString(), this._lib.getIndex(i - 0x12) + num5.ToString()).NumberFormat = "#,##";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(i) + num7.ToString()).Borders.LineStyle = XlBorderWeight.xlHairline;
                string[] strArray = new string[] { "ma_lk", "ma_bn", "ma_dkbd" };
                for (int k = 0; k < strArray.Length; k++)
                {
                    try
                    {
                        int ordinal = dset.Tables[0].Columns[strArray[k]].Ordinal;
                        this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + num4, this._lib.getIndex(ordinal) + num5);
                        this.orange.NumberFormat = "@";
                    }
                    catch
                    {
                    }
                }
                for (int m = 0; m < dset.Tables[0].Rows.Count; m++)
                {
                    for (int n = 0; n < strArray.Length; n++)
                    {
                        int num13 = dset.Tables[0].Columns[strArray[n]].Ordinal;
                        try
                        {
                            this.osheet.Cells[num4 + m, num13 + 1] = dset.Tables[0].Rows[m][num13].ToString();
                        }
                        catch
                        {
                        }
                    }
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(i + 2) + num5.ToString());
                this.orange.Font.Name = "Arial";
                this.orange.Font.Size = 8;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 4] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NGOẠI TR\x00da ";
                this.osheet.Cells[2, 4] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(i) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception2)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception2.Message, this._lib.Msg);
            }
        }

        public void f_Ngoaitru_xuatExcel_mau01_TT324(bool print, DataSet dsdulieu, string tungay, string denngay, DataSet vdsthuoc, DataSet vdscls, int loaigiamdinh, int kygiamdinh, string namgiamdinh, string ngaylaphoso, bool tachdsbenhnhan, bool mahoadulieu, bool tt324moi, string makp_ngoaitru, bool xuatfileexcel, string v_duongdanxuatxml)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("ma_lk", typeof(string));
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
            dset.Tables[0].Columns.Add("ten_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benhkhac", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ma_tai_nan", typeof(int));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("ngay_ttoan", typeof(string));
            dset.Tables[0].Columns.Add("muc_huong", typeof(decimal));
            dset.Tables[0].Columns.Add("t_thuoc", typeof(decimal));
            dset.Tables[0].Columns.Add("t_vtyt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_nguonkhac", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_PTTT_QT", typeof(string));
            dset.Tables[0].Columns.Add("can_nang", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_tongct", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sobienlai", typeof(string));
            dset.Tables[0].Columns.Add("quyenso", typeof(string));
            dset.Tables[0].Columns.Add("nguoithu", typeof(string));
            for (int i = 1; i <= 15; i++)
            {
                dset.Tables[0].Columns.Add("t_" + i.ToString(), typeof(decimal)).DefaultValue = 0;
            }
            dsdulieu.WriteXml("tam.xml", XmlWriteMode.WriteSchema);
            vdscls.WriteXml("cls.xml", XmlWriteMode.WriteSchema);
            vdsthuoc.WriteXml("thuoc.xml", XmlWriteMode.WriteSchema);
            string malk = "";
            string mabv = this._lib.Mabv;
            string tenbv = this._lib.Tenbv;
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            DataSet set3 = this._lib.f_get_dskhoaphong();
            int num2 = this._lib.iMavp_congkham(1);
            DataSet set4 = new DataSet();
            decimal num3 = 100M;
            set4.Tables.Add();
            set4.Tables[0].Columns.Add("id", typeof(decimal));
            set4.Tables[0].Columns.Add("ten");
            string str4 = "";
            string str5 = "";
            string str6 = "";
            string makhoa = "";
            if (v_duongdanxuatxml == "")
            {
                v_duongdanxuatxml = this._s_xmlpath;
            }
            StreamWriter writer = new StreamWriter(v_duongdanxuatxml + "bhxh_" + denngay.Substring(3, 2) + denngay.Substring(8) + ".xml");
            if (!tachdsbenhnhan)
            {
                str4 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<GIAMDINHHS xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + this._s_arrCap[1] + "<THONGTINDONVI>" + this._s_arrCap[2] + "<MACSKCB>" + mabv + "</MACSKCB>" + this._s_arrCap[1] + "</THONGTINDONVI>" + this._s_arrCap[1] + "<THONGTINHOSO>" + this._s_arrCap[2] + "<NGAYLAP>" + ((ngaylaphoso == "") ? DateTime.Now.ToString("yyyyMMdd") : ngaylaphoso) + "</NGAYLAP>" + this._s_arrCap[2] + "<SOLUONGHOSO>" + dsdulieu.Tables[0].Rows.Count.ToString() + "</SOLUONGHOSO>" + this._s_arrCap[2] + "<DANHSACHHOSO>";
                writer.WriteLine(str4);
                str5 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                str5 = str5 + this._s_arrCap[0] + "<DSACH_CHI_TIET_THUOC>";
                str6 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                str6 = str6 + this._s_arrCap[0] + "<DSACH_CHI_TIET_DVKT>";
            }
            decimal d = 0M;
            decimal num5 = this._lib.themoi15_sotien();
            string exp = "";
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    exp = (((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4")) ? ("mavaovien=" + row["mavaovien"].ToString()) : (" ngayduyet='" + row["ngay"].ToString().Substring(0, 10) + "' and sothe='" + row["sothe"].ToString() + "'")) + " and mabn='" + row["mabn"].ToString() + "'";
                    try
                    {
                        if (this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp) <= 0M)
                        {
                            this._lib.f_write_log("Benh nhan co tong chi phi bang 0: " + row["mabn"].ToString() + " - " + row["HOTEN"].ToString() + " so the: " + row["sothe"].ToString() + "  " + exp);
                            continue;
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        foreach (DataRow row2 in vdsthuoc.Tables[0].Select(exp + " and sotien>gia_bh_toida  and gia_bh_toida>0", "soluong"))
                        {
                            row2["sotien"] = row2["gia_bh_toida"].ToString();
                            row2["dongia"] = decimal.Parse(row2["gia_bh_toida"].ToString()) / decimal.Parse(row2["soluong"].ToString());
                        }
                        vdsthuoc.AcceptChanges();
                    }
                    catch (Exception exception)
                    {
                        this._lib.f_write_log("gia_bh_toida:" + exception.ToString());
                    }
                    try
                    {
                        foreach (DataRow row3 in vdscls.Tables[0].Select(exp + " and sotien>gia_bh_toida  and gia_bh_toida>0", "soluong"))
                        {
                            row3["sotien"] = row3["gia_bh_toida"].ToString();
                            row3["dongia"] = decimal.Parse(row3["gia_bh_toida"].ToString()) / decimal.Parse(row3["soluong"].ToString());
                        }
                        vdscls.AcceptChanges();
                    }
                    catch (Exception exception2)
                    {
                        this._lib.f_write_log("gia_bh_toida:" + exception2.ToString());
                    }
                    try
                    {
                        foreach (DataRow row4 in vdsthuoc.Tables[0].Select(exp + " and soluong<0", "soluong"))
                        {
                            this.f_update_soam_324moi(row4["soluong"].ToString(), exp + " and soluong>0 and mavp=" + row4["mavp"].ToString() + " and dongia=" + row4["dongia"].ToString(), ref vdsthuoc);
                            row4["soluong"] = 0;
                            row4["sotien"] = 0;
                        }
                        vdsthuoc.AcceptChanges();
                    }
                    catch
                    {
                    }
                    DataRow row5 = dset.Tables[0].NewRow();
                    try
                    {
                        row5["ma_lk"] = this._lib.Mabv + row["malk"].ToString();
                    }
                    catch
                    {
                        row5["ma_lk"] = row["malk"].ToString();
                    }
                    row5["STT"] = d = decimal.op_Increment(d);
                    row5["ma_bn"] = row["mabn"].ToString();
                    row5["ho_ten"] = row["HOTEN"].ToString();
                    row5["ngay_sinh"] = row["ngaysinh"].ToString();
                    row5["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row5["dia_chi"] = row["diachi"].ToString();
                    try
                    {
                        row5["ma_the"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row5["ma_the"] = row["sothe"].ToString();
                    }
                    row5["ma_dkbd"] = row["MANOIDK"].ToString();
                    if (row["tungay"].ToString() == "")
                    {
                        try
                        {
                            DataSet set5 = this._lib.f_get_sothebhyt_tungay(row["sothe"].ToString(), row["NGAYVAO"].ToString(), row["NGAYRA"].ToString());
                            row["tungay"] = set5.Tables[0].Rows[0]["tungay"].ToString();
                            row["denngay"] = set5.Tables[0].Rows[0]["denngay"].ToString();
                        }
                        catch
                        {
                        }
                    }
                    try
                    {
                        row5["sobienlai"] = row["sobienlai"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row5["quyenso"] = row["quyenso"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row5["nguoithu"] = row["ten_thuvp"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row5["gt_the_tu"] = row["tungay"].ToString().Substring(6, 4) + row["tungay"].ToString().Substring(3, 2) + row["tungay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row5["gt_the_tu"] = "";
                    }
                    try
                    {
                        row5["gt_the_den"] = row["denngay"].ToString().Substring(6, 4) + row["denngay"].ToString().Substring(3, 2) + row["denngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row5["gt_the_den"] = "";
                    }
                    try
                    {
                        row5["ngay_ttoan"] = row["ngay"].ToString().Substring(6, 4) + row["ngay"].ToString().Substring(3, 2) + row["ngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row5["ngay_ttoan"] = "";
                    }
                    try
                    {
                        row5["ngay_ttoan"] = row5["ngay_ttoan"].ToString() + row["ngay"].ToString().Substring(11, 2) + row["ngay"].ToString().Substring(14, 2);
                    }
                    catch
                    {
                        row5["ngay_ttoan"] = row5["ngay_ttoan"].ToString() + "0000";
                    }
                    row5["ma_benh"] = row["MAICD"].ToString();
                    try
                    {
                        row5["ma_benh"] = row5["ma_benh"].ToString().Split(new char[] { ';' })[0];
                    }
                    catch
                    {
                    }
                    if (row5["ma_benh"].ToString() == "")
                    {
                        row5["ma_benh"] = row["MAICD"].ToString();
                    }
                    row5["ma_benhkhac"] = row["maicdkt"].ToString();
                    row5["ten_benh"] = row["chandoan"].ToString();
                    if (row["traituyen"].ToString() != "0")
                    {
                        row5["ma_lydo_vvien"] = 3;
                    }
                    else if (row["traituyen"].ToString() == "0")
                    {
                        row5["ma_lydo_vvien"] = 1;
                    }
                    else
                    {
                        row5["ma_lydo_vvien"] = 2;
                    }
                    row5["ma_noi_chuyen"] = "";
                    row5["ma_tai_nan"] = 0;
                    try
                    {
                        row5["ngay_vao"] = row["NGAYVV"].ToString();
                    }
                    catch
                    {
                        row5["ngay_vao"] = "";
                    }
                    try
                    {
                        row5["ngay_ra"] = row["NGAYRV"].ToString();
                    }
                    catch
                    {
                        row5["ngay_ra"] = "";
                    }
                    try
                    {
                        if (decimal.Parse(row5["ngay_ra"].ToString()) < decimal.Parse(row5["ngay_vao"].ToString()))
                        {
                            row5["ngay_vao"] = row5["ngay_ra"].ToString();
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        if (decimal.Parse(row5["ngay_ttoan"].ToString()) < decimal.Parse(row5["ngay_ra"].ToString()))
                        {
                            row5["ngay_ttoan"] = row5["ngay_ra"].ToString();
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        row5["so_ngay_dtri"] = row["SONGAY"].ToString();
                    }
                    catch
                    {
                        row5["so_ngay_dtri"] = 1;
                    }
                    row5["ket_qua_dtri"] = 1;
                    row5["tinh_trang_rv"] = 1;
                    try
                    {
                        row5["muc_huong"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                    }
                    catch
                    {
                    }
                    row5["t_thuoc"] = 0;
                    row5["t_vtyt"] = 0;
                    row5["t_tongchi"] = row["tongcong"].ToString();
                    row5["t_bntt"] = row["bntra"].ToString();
                    row5["t_bhtt"] = row["bhyttra"].ToString();
                    row5["t_nguonkhac"] = 0;
                    row5["t_ngoaids"] = 0;
                    row5["nam_qt"] = denngay.Substring(6, 4);
                    row5["thang_qt"] = denngay.Substring(3, 2);
                    row5["ma_loaikcb"] = 1;
                    if (row["loaiba"].ToString() == "1")
                    {
                        row5["ma_loaikcb"] = 3;
                    }
                    else if (row["loaiba"].ToString() == "2")
                    {
                        row5["ma_loaikcb"] = 2;
                    }
                    if (makp_ngoaitru != "")
                    {
                        try
                        {
                            if (makp_ngoaitru == "2")
                            {
                                row5["ma_loaikcb"] = 2;
                            }
                        }
                        catch
                        {
                        }
                    }
                    row5["ma_cskcb"] = this._lib.MABHXH;
                    try
                    {
                        row5["ma_khuvuc"] = row["makhuvuc"].ToString();
                    }
                    catch
                    {
                    }
                    row5["ma_PTTT_QT"] = "";
                    dset.Tables[0].Rows.Add(row5);
                    try
                    {
                        malk = row5["ma_lk"].ToString();
                        decimal num6 = 0M;
                        try
                        {
                            if (row["loaiba"].ToString() == "30")
                            {
                                num6 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + "  and nhombhyt=9");
                                if (num6 == 0M)
                                {
                                    try
                                    {
                                        str4 = "id=" + num2.ToString();
                                        DataRow row6 = set2.Tables[0].Select(str4)[0];
                                        DataRow row7 = vdscls.Tables[0].NewRow();
                                        try
                                        {
                                            row7.ItemArray = vdscls.Tables[0].Select(exp)[0].ItemArray;
                                        }
                                        catch
                                        {
                                        }
                                        vdscls.Tables[0].Rows.Add(row7);
                                        try
                                        {
                                            row7["sothe"] = row5["ma_the"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                        row7["nhombhyt"] = 9;
                                        row7["mavp"] = row6["id"].ToString();
                                        row7["sotien"] = row["congkham"].ToString();
                                        row7["dongia"] = row["congkham"].ToString();
                                        row7["soluong"] = 1;
                                        row7["dvt"] = row6["dvt"].ToString();
                                        row7["mavp1"] = row6["ma"].ToString();
                                        row7["ten"] = row6["ten"].ToString();
                                        try
                                        {
                                            row7["malk"] = row["malk"].ToString();
                                        }
                                        catch
                                        {
                                            row7["malk"] = row["mavaovien"].ToString();
                                        }
                                        row7["mavaovien"] = row["mavaovien"].ToString();
                                        row7["ngayduyet"] = row["ngay"].ToString().Substring(0, 10);
                                        row7["mabn"] = row["mabn"].ToString();
                                        row7["ngayylenh"] = row["ngayvv"].ToString();
                                        try
                                        {
                                            row7["maql"] = row["maql"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                    catch (Exception exception3)
                                    {
                                        this._lib.f_write_log("them cong kham =0: " + exception3.ToString());
                                    }
                                    vdscls.AcceptChanges();
                                }
                                else
                                {
                                    try
                                    {
                                        str4 = "id=" + num2.ToString();
                                        DataRow row8 = set2.Tables[0].Select(str4)[0];
                                        DataRow[] rowArray = vdscls.Tables[0].Select(exp + "  and nhombhyt=9");
                                        try
                                        {
                                            rowArray[0]["sotien"] = row["congkham"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                        rowArray[0]["dongia"] = rowArray[0]["sotien"].ToString();
                                        rowArray[0]["soluong"] = 1;
                                        for (int j = 1; j < rowArray.Length; j++)
                                        {
                                            rowArray[j]["sotien"] = 0;
                                        }
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                        foreach (DataRow row9 in vdscls.Tables[0].Select(exp + " and nhombhyt=11", "soluong"))
                        {
                            try
                            {
                                row9["bhyt"] = 100;
                                if (row9["soluong"].ToString() == "0.5")
                                {
                                    row9["soluong"] = 1;
                                    row9["bhyt"] = 50;
                                    row9["sotien"] = decimal.Parse(row9["dongia"].ToString()) / 2M;
                                }
                                else if (row9["soluong"].ToString().IndexOf(".5") > -1)
                                {
                                    DataRow row10 = vdscls.Tables[0].NewRow();
                                    row10.ItemArray = row9.ItemArray;
                                    row10["soluong"] = 1;
                                    row10["bhyt"] = 50;
                                    row10["sotien"] = decimal.Parse(row9["dongia"].ToString()) / 2M;
                                    row10["bhyttra"] = 0;
                                    vdscls.Tables[0].Rows.Add(row10);
                                    row9["soluong"] = double.Parse(row9["soluong"].ToString()) - 0.5;
                                    row9["sotien"] = decimal.Parse(row9["soluong"].ToString()) * decimal.Parse(row9["dongia"].ToString());
                                }
                            }
                            catch
                            {
                            }
                        }
                        vdscls.AcceptChanges();
                        try
                        {
                            num6 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + " and nhombhyt in(3,4) and sotien>0", true);
                            num6 = Math.Round(num6);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_tongchi"] = Math.Round(Convert.ToDecimal(row5["t_tongchi"].ToString()));
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_vtyt"] = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + " and nhombhyt in(6,14) and sotien>0", true);
                            row5["t_vtyt"] = Math.Round(Convert.ToDouble(row5["t_vtyt"].ToString()));
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_tongct"] = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + " and sotien>0 and nhombhyt not in(3,4,6,14)");
                            row5["t_tongct"] = (Math.Round(Convert.ToDecimal(row5["t_tongct"].ToString())) + num6) + Convert.ToDecimal(row5["t_vtyt"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_bhtt"] = decimal.Parse(row5["t_tongct"].ToString()) * (decimal.Parse(row5["muc_huong"].ToString()) / 100M);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_bntt"] = decimal.Parse(row5["t_tongct"].ToString()) - decimal.Parse(row5["t_bhtt"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            row5["t_thuoc"] = Math.Round(num6);
                        }
                        catch
                        {
                        }
                        try
                        {
                            num6 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + " and nhombhyt in(14,7)");
                            if (num6 > 48400000M)
                            {
                                row5["t_bntt"] = decimal.Parse(row5["t_bntt"].ToString()) + (num6 - 48400000M);
                                row5["t_bhtt"] = decimal.Parse(row5["t_tongct"].ToString()) - decimal.Parse(row5["t_bntt"].ToString());
                            }
                        }
                        catch
                        {
                        }
                        foreach (DataRow row11 in vdscls.Tables[0].Select(exp + " and nhombhyt=10"))
                        {
                            if ((row["traituyen"].ToString() != "1") || (row["sothe"].ToString().IndexOf("TE1") != -1))
                            {
                                decimal num8 = decimal.Parse(row11["sotien"].ToString()) * (1M - (decimal.Parse(row5["muc_huong"].ToString()) / 100M));
                                row5["t_bhtt"] = decimal.Parse(row5["t_bhtt"].ToString()) + num8;
                                row5["t_bntt"] = decimal.Parse(row5["t_bntt"].ToString()) - num8;
                            }
                        }
                        if (Math.Abs((double) (double.Parse(row5["t_bhtt"].ToString()) - double.Parse(row5["t_tongct"].ToString()))) < 5.0)
                        {
                            row5["muc_huong"] = 100;
                        }
                        num3 = 100M;
                        if (row["traituyen"].ToString() != "0")
                        {
                            num3 = 60M;
                        }
                        try
                        {
                            makhoa = set3.Tables[0].Select("makp='" + row["makp"].ToString() + "'")[0]["makp_byt"].ToString();
                        }
                        catch
                        {
                            makhoa = "";
                        }
                        string ngaythanhtoan = row5["ngay_ttoan"].ToString();
                        try
                        {
                            if (decimal.Parse(row5["ngay_ttoan"].ToString()) < decimal.Parse(row5["ngay_ra"].ToString()))
                            {
                                ngaythanhtoan = row5["ngay_ra"].ToString();
                            }
                        }
                        catch
                        {
                        }
                        try
                        {
                            if (tachdsbenhnhan)
                            {
                                writer = new StreamWriter(this._s_xmlpath + "bhxh_" + row5["ma_the"].ToString() + "_" + row["ngayrv"].ToString() + ".xml");
                                str4 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<GIAMDINHHS xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + this._s_arrCap[1] + "<THONGTINDONVI>" + this._s_arrCap[2] + "<MACSKCB>" + mabv + "</MACSKCB>" + this._s_arrCap[1] + "</THONGTINDONVI>" + this._s_arrCap[1] + "<THONGTINHOSO>" + this._s_arrCap[2] + "<NGAYLAP>" + ((ngaylaphoso == "") ? DateTime.Now.ToString("yyyyMMdd") : ngaylaphoso) + "</NGAYLAP>" + this._s_arrCap[2] + "<SOLUONGHOSO>1</SOLUONGHOSO>" + this._s_arrCap[2] + "<DANHSACHHOSO>";
                                writer.WriteLine(str4);
                                str5 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                                str5 = str5 + this._s_arrCap[0] + "<DSACH_CHI_TIET_THUOC>";
                                str6 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                                str6 = str6 + this._s_arrCap[0] + "<DSACH_CHI_TIET_DVKT>";
                            }
                            str4 = this._s_arrCap[3] + "<HOSO>" + this._s_arrCap[4] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML1</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>";
                            writer.WriteLine(str4);
                            str4 = this._s_arrCap[0] + "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                            str4 = str4 + this.f_nodexml_tonghopttbn(1, malk, row5["STT"].ToString(), row5["ma_bn"].ToString(), !tt324moi ? row5["ho_ten"].ToString() : this.f_cdata_tag(row5["ho_ten"].ToString().ToLower()), row5["ngay_sinh"].ToString(), row5["gioi_tinh"].ToString(), !tt324moi ? row5["dia_chi"].ToString() : this.f_cdata_tag(row5["dia_chi"].ToString()), row5["ma_the"].ToString(), row5["ma_dkbd"].ToString(), row5["gt_the_tu"].ToString(), row5["gt_the_den"].ToString(), row5["ma_benh"].ToString(), !tt324moi ? row5["ten_benh"].ToString() : this.f_cdata_tag(row5["ten_benh"].ToString()), row5["ma_lydo_vvien"].ToString(), "", "0", row5["ngay_vao"].ToString(), row5["ngay_ra"].ToString(), (row5["so_ngay_dtri"].ToString() == "0") ? "1" : row5["so_ngay_dtri"].ToString(), row5["ket_qua_dtri"].ToString(), row5["tinh_trang_rv"].ToString(), ngaythanhtoan, row5["muc_huong"].ToString(), row5["t_thuoc"].ToString(), row5["t_vtyt"].ToString(), Math.Round(Convert.ToDecimal(row5["t_tongct"].ToString())), Math.Round(Convert.ToDecimal(row5["t_bntt"].ToString())), Math.Round(Convert.ToDecimal(row5["t_bhtt"].ToString())), "0", denngay.Substring(6, 4), denngay.Substring(3, 2), Convert.ToInt32(row5["ma_loaikcb"].ToString()), makhoa, mabv, "", "0", true);
                        }
                        catch
                        {
                        }
                        str4 = mahoadulieu ? this.f_convert_to_unicode(str4) : str4;
                        string str37 = str4;
                        str4 = str37 + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                        writer.WriteLine(str4);
                        string strText = "";
                        int num9 = 0;
                        foreach (DataRow row12 in vdsthuoc.Tables[0].Select(exp + " and nhombhyt in(3,4) and soluong>0", "nhombhyt"))
                        {
                            string str = "";
                            string mathuoc = "";
                            string manhom = "";
                            string ngayylenh = "";
                            string str15 = "";
                            string str16 = "100";
                            try
                            {
                                str = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["lieuluong"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                mathuoc = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["mabyt"].ToString();
                            }
                            catch
                            {
                            }
                            if (mathuoc == "")
                            {
                                mathuoc = row12["mavp1"].ToString();
                            }
                            try
                            {
                                manhom = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["maloaibhyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str15 = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                                if (str15 == "")
                                {
                                    str15 = row12["ten"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                str16 = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["bhyt"].ToString();
                                str16 = (str16 == "0") ? "100" : str16;
                                if (str16 == "50")
                                {
                                    row12["dongia"] = decimal.Parse(row12["dongia"].ToString()) * 2M;
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                row12["duongdung"] = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["duongdung"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                row12["dvt"] = set2.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0]["dvtbyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                ngayylenh = row12["ngayylenh"].ToString().Substring(6, 4) + row12["ngayylenh"].ToString().Substring(3, 2) + row12["ngayylenh"].ToString().Substring(0, 2);
                                if (row12["ngayylenh"].ToString().Length > 10)
                                {
                                    ngayylenh = ngayylenh + row12["ngayylenh"].ToString().Substring(11, 2) + row12["ngayylenh"].ToString().Substring(14, 2);
                                }
                            }
                            catch
                            {
                                ngayylenh = row["ngayvv"].ToString();
                            }
                            try
                            {
                                if (((manhom == "") || (mathuoc == "")) && (set4.Tables[0].Select("id=" + row12["mavp"].ToString()).Length == 0))
                                {
                                    set4.Tables[0].Rows.Add(new object[] { decimal.Parse(row12["mavp"].ToString()), row12["ten"].ToString() });
                                }
                                row5["t_" + manhom] = decimal.Parse(row5["t_" + manhom].ToString()) + decimal.Parse(row12["sotien"].ToString());
                            }
                            catch
                            {
                            }
                            strText = strText + this.f_nodexml_chitietthuoc_bhxh(1, malk, ++num9, mathuoc, manhom, !tt324moi ? str15 : this.f_cdata_tag(str15), row12["dvt"].ToString(), !tt324moi ? row12["hamluong"].ToString() : this.f_cdata_tag(row12["hamluong"].ToString()), row12["duongdung"].ToString(), !tt324moi ? str : this.f_cdata_tag(str), row12["sodk"].ToString(), Math.Round(Convert.ToDecimal(row12["soluong"].ToString()), 2), Math.Round(Convert.ToDecimal(row12["dongia"].ToString()), 2), Convert.ToDecimal(str16), Math.Round(Convert.ToDecimal(row12["sotien"].ToString()), 2), makhoa, row12["kihieubs"].ToString(), row5["ma_benh"].ToString(), ngayylenh, "0");
                        }
                        foreach (DataRow row13 in vdscls.Tables[0].Select(exp + " and nhombhyt in(3,4) and sotien>0", "nhombhyt"))
                        {
                            string str17 = "";
                            string str18 = "";
                            string str19 = "";
                            string str20 = "";
                            string str21 = "";
                            string str22 = "100";
                            try
                            {
                                str17 = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["lieuluong"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str18 = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["mabyt"].ToString();
                            }
                            catch
                            {
                            }
                            if (str18 == "")
                            {
                                str18 = row13["mavp1"].ToString();
                            }
                            try
                            {
                                str19 = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["maloaibhyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str21 = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                                if (str21 == "")
                                {
                                    str21 = row13["ten"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                str22 = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["bhyt"].ToString();
                                str22 = (str22 == "0") ? "100" : str22;
                                if (str22 == "50")
                                {
                                    row13["dongia"] = decimal.Parse(row13["dongia"].ToString()) * 2M;
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                row13["duongdung"] = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["duongdung"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                row13["dvt"] = set2.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0]["dvtbyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str20 = row13["ngayylenh"].ToString().Substring(6, 4) + row13["ngayylenh"].ToString().Substring(3, 2) + row13["ngayylenh"].ToString().Substring(0, 2);
                                if (row13["ngayylenh"].ToString().Length > 10)
                                {
                                    str20 = str20 + row13["ngayylenh"].ToString().Substring(11, 2) + row13["ngayylenh"].ToString().Substring(14, 2);
                                }
                            }
                            catch
                            {
                                str20 = row["ngayvv"].ToString();
                            }
                            try
                            {
                                if (((str19 == "") || (str18 == "")) && (set4.Tables[0].Select("id=" + row13["mavp"].ToString()).Length == 0))
                                {
                                    set4.Tables[0].Rows.Add(new object[] { decimal.Parse(row13["mavp"].ToString()), row13["ten"].ToString() });
                                }
                                row5["t_" + str19] = decimal.Parse(row5["t_" + str19].ToString()) + decimal.Parse(row13["sotien"].ToString());
                            }
                            catch
                            {
                            }
                            strText = strText + this.f_nodexml_chitietthuoc_bhxh(1, malk, ++num9, str18, str19, !tt324moi ? str21 : this.f_cdata_tag(str21), row13["dvt"].ToString(), !tt324moi ? "" : this.f_cdata_tag(""), "", !tt324moi ? str17 : this.f_cdata_tag(str17), "", Math.Round(Convert.ToDecimal(row13["soluong"].ToString()), 2), Math.Round(Convert.ToDecimal(row13["dongia"].ToString()), 2), Convert.ToDecimal(str22), Math.Round(Convert.ToDecimal(row13["sotien"].ToString()), 2), makhoa, row13["kihieubs"].ToString(), row5["ma_benh"].ToString(), str20, "0");
                        }
                        if (strText != "")
                        {
                            strText = (this._s_arrCap[0] + str5 + strText) + this._s_arrCap[0] + "</DSACH_CHI_TIET_THUOC>";
                            str4 = this._s_arrCap[3] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML2</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>" + (mahoadulieu ? this.f_convert_to_unicode(strText) : strText) + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                            writer.WriteLine(str4);
                        }
                        strText = "";
                        foreach (DataRow row14 in vdsthuoc.Tables[0].Select(exp + " and nhombhyt not in(3,4) and sotien>0", "nhombhyt"))
                        {
                            string mavattu = "";
                            string str24 = "";
                            string ma = "";
                            string str26 = "";
                            string str27 = "";
                            string str28 = "100";
                            try
                            {
                                if ((row14["nhombhyt"].ToString() == "6") || (row14["nhombhyt"].ToString() == "14"))
                                {
                                    mavattu = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["mabyt"].ToString();
                                    if (row14["nhombhyt"].ToString() == "14")
                                    {
                                        try
                                        {
                                            string str29 = vdsthuoc.Tables[0].Select(exp + " and nhombhyt in(7) and sotien>0", "nhombhyt")[0]["mavp"].ToString();
                                            ma = set2.Tables[0].Select("id=" + str29 + "")[0]["mabyt"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                else
                                {
                                    ma = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["mabyt"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                str24 = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["maloaibhyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str27 = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                                if (str27 == "")
                                {
                                    str27 = row14["ten"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                row14["dvt"] = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["dvtbyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str26 = row14["ngayylenh"].ToString().Substring(6, 4) + row14["ngayylenh"].ToString().Substring(3, 2) + row14["ngayylenh"].ToString().Substring(0, 2);
                                if (row14["ngayylenh"].ToString().Length > 10)
                                {
                                    str26 = str26 + row14["ngayylenh"].ToString().Substring(11, 2) + row14["ngayylenh"].ToString().Substring(14, 2);
                                }
                            }
                            catch
                            {
                                str26 = row["ngayvv"].ToString();
                            }
                            try
                            {
                                str28 = set2.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0]["bhyt"].ToString();
                                str28 = (str28 == "0") ? "100" : str28;
                                if (str28 == "50")
                                {
                                    row14["dongia"] = decimal.Parse(row14["dongia"].ToString()) * 2M;
                                }
                                try
                                {
                                    if (row14["nhombhyt"].ToString() == "11")
                                    {
                                        str28 = row14["bhyt"].ToString();
                                    }
                                }
                                catch
                                {
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                if (((str24 == "") || ((ma == "") && (mavattu == ""))) && (set4.Tables[0].Select("id=" + row14["mavp"].ToString()).Length == 0))
                                {
                                    set4.Tables[0].Rows.Add(new object[] { decimal.Parse(row14["mavp"].ToString()), row14["ten"].ToString() });
                                }
                                row5["t_" + str24] = decimal.Parse(row5["t_" + str24].ToString()) + decimal.Parse(row14["sotien"].ToString());
                            }
                            catch
                            {
                            }
                            strText = strText + this.f_nodexml_chitietcls_bhxh(1, malk, ++num9, ma, mavattu, str24, !tt324moi ? str27 : this.f_cdata_tag(str27), row14["dvt"].ToString(), Math.Round(Convert.ToDecimal(row14["soluong"].ToString()), 2), Math.Round(Convert.ToDecimal(row14["dongia"].ToString()), 2), Convert.ToDecimal(str28), Math.Round(Convert.ToDecimal(row14["sotien"].ToString()), 2), "", row14["kihieubs"].ToString(), row5["ma_benh"].ToString(), str26, str26, "0");
                        }
                        foreach (DataRow row15 in vdscls.Tables[0].Select(exp + " and nhombhyt not in(3,4) and sotien>0", "nhombhyt"))
                        {
                            string str30 = "";
                            string str31 = "";
                            string str32 = "";
                            string str33 = "";
                            string str34 = "";
                            string str35 = "100";
                            try
                            {
                                if ((row15["nhombhyt"].ToString() == "6") || (row15["nhombhyt"].ToString() == "14"))
                                {
                                    str30 = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["mabyt"].ToString();
                                    if (row15["nhombhyt"].ToString() == "14")
                                    {
                                        try
                                        {
                                            string str36 = vdscls.Tables[0].Select(exp + " and nhombhyt in(7) and sotien>0", "nhombhyt")[0]["mavp"].ToString();
                                            str32 = set2.Tables[0].Select("id=" + str36 + "")[0]["mabyt"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                else
                                {
                                    str32 = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["mabyt"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                str31 = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["maloaibhyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str34 = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["tenbyt"].ToString();
                                if (str34 == "")
                                {
                                    str34 = row15["ten"].ToString();
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                str33 = row15["ngayylenh"].ToString().Substring(6, 4) + row15["ngayylenh"].ToString().Substring(3, 2) + row15["ngayylenh"].ToString().Substring(0, 2);
                                if (row15["ngayylenh"].ToString().Length > 10)
                                {
                                    str33 = str33 + row15["ngayylenh"].ToString().Substring(11, 2) + row15["ngayylenh"].ToString().Substring(14, 2);
                                }
                            }
                            catch
                            {
                                str33 = row["ngayvv"].ToString();
                            }
                            try
                            {
                                str35 = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["bhyt"].ToString();
                                str35 = (str35 == "0") ? "100" : str35;
                                if (str35 == "50")
                                {
                                    row15["dongia"] = decimal.Parse(row15["dongia"].ToString()) * 2M;
                                }
                                try
                                {
                                    if (row15["nhombhyt"].ToString() == "11")
                                    {
                                        str35 = row15["bhyt"].ToString();
                                    }
                                    else if (row15["nhombhyt"].ToString() == "9")
                                    {
                                        try
                                        {
                                            decimal num10 = decimal.Parse(vdscls.Tables[0].Select(exp + " and nhombhyt=9", "")[0]["congkham"].ToString());
                                            if (num10 == 0M)
                                            {
                                                num10 = decimal.Parse(row15["dongia"].ToString());
                                            }
                                            str35 = ((100 * int.Parse(row15["dongia"].ToString())) / num10);
                                            row15["dongia"] = num10;
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                                catch
                                {
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                row15["dvt"] = set2.Tables[0].Select("id=" + row15["mavp"].ToString() + "")[0]["dvtbyt"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                if (((str31 == "") || ((str32 == "") && (str30 == ""))) && (set4.Tables[0].Select("id=" + row15["mavp"].ToString()).Length == 0))
                                {
                                    set4.Tables[0].Rows.Add(new object[] { decimal.Parse(row15["mavp"].ToString()), row15["ten"].ToString() });
                                }
                                row5["t_" + str31] = decimal.Parse(row5["t_" + str31].ToString()) + decimal.Parse(row15["sotien"].ToString());
                            }
                            catch
                            {
                            }
                            strText = strText + this.f_nodexml_chitietcls_bhxh(1, malk, ++num9, str32, str30, str31, !tt324moi ? str34 : this.f_cdata_tag(str34), row15["dvt"].ToString(), Convert.ToDecimal(row15["soluong"].ToString()), Math.Round(Convert.ToDecimal(row15["dongia"].ToString()), 2), Convert.ToDecimal(str35), Math.Round(Convert.ToDecimal(row15["sotien"].ToString()), 2), "", row15["kihieubs"].ToString(), row5["ma_benh"].ToString(), str33, str33, "0");
                        }
                        if (strText != "")
                        {
                            strText = (this._s_arrCap[0] + str6 + strText) + this._s_arrCap[0] + "</DSACH_CHI_TIET_DVKT>";
                            str4 = this._s_arrCap[3] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML3</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>" + (mahoadulieu ? this.f_convert_to_unicode(strText) : strText) + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                            writer.WriteLine(str4);
                        }
                        writer.WriteLine(this._s_arrCap[3] + "</HOSO>");
                        if (tachdsbenhnhan)
                        {
                            str4 = this._s_arrCap[2] + "</DANHSACHHOSO>" + this._s_arrCap[1] + "</THONGTINHOSO>" + this._s_arrCap[1] + "<CHUKYDONVI />\r\n</GIAMDINHHS>";
                            writer.WriteLine(str4);
                            writer.Close();
                        }
                    }
                    catch (Exception exception4)
                    {
                        this._lib.f_write_log(row["mabn"].ToString() + " - " + row["HOTEN"].ToString() + " so the: " + row["sothe"].ToString() + "  " + exception4.ToString());
                    }
                }
                catch (Exception exception5)
                {
                    MessageBox.Show(exception5.Message, this._lib.Msg);
                }
            }
            if (!tachdsbenhnhan)
            {
                str4 = this._s_arrCap[2] + "</DANHSACHHOSO>" + this._s_arrCap[1] + "</THONGTINHOSO>" + this._s_arrCap[1] + "<CHUKYDONVI />\r\n</GIAMDINHHS>";
                writer.WriteLine(str4);
                writer.Close();
            }
            dset.WriteXml("bang01_ngoai.xml", XmlWriteMode.WriteSchema);
            set4.WriteXml("dsmanhomdichvu.xml", XmlWriteMode.WriteSchema);
            if (xuatfileexcel)
            {
                int num11 = 0;
                int num12 = 0;
                int num13 = 0;
                num11 = dset.Tables[0].Rows.Count + 5;
                num12 = dset.Tables[0].Columns.Count - 1;
                num13 = num11;
                this.tenfile = this._lib.Export_Excel(dset, "bccpkcb01");
                try
                {
                    Process.Start(this.tenfile);
                }
                catch
                {
                }
            }
        }

        public void f_Ngoaitru_xuatExcel_mau01_TT4210(bool print, DataSet dsdulieu, string tungay, string denngay, DataSet vdsthuoc, DataSet vdscls, int loaigiamdinh, int kygiamdinh, string namgiamdinh, string ngaylaphoso, bool tachdsbenhnhan, bool mahoadulieu, bool tt324moi, string makp_ngoaitru, bool xuatfileexcel, string v_duongdanxuatxml)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("ma_lk", typeof(string));
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
            dset.Tables[0].Columns.Add("ten_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benhkhac", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ma_tai_nan", typeof(int));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("ngay_ttoan", typeof(string));
            dset.Tables[0].Columns.Add("muc_huong", typeof(decimal));
            dset.Tables[0].Columns.Add("t_thuoc", typeof(decimal));
            dset.Tables[0].Columns.Add("t_vtyt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_nguonkhac", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_PTTT_QT", typeof(string));
            dset.Tables[0].Columns.Add("can_nang", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_tongct", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("sobienlai", typeof(string));
            dset.Tables[0].Columns.Add("quyenso", typeof(string));
            dset.Tables[0].Columns.Add("nguoithu", typeof(string));
            for (int i = 1; i <= 15; i++)
            {
                dset.Tables[0].Columns.Add("t_" + i.ToString(), typeof(decimal)).DefaultValue = 0;
            }
            try
            {
                vdscls.Tables[0].Columns.Add("bhyt", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                vdsthuoc.Tables[0].Columns.Add("bhyt", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                vdscls.Tables[0].Columns.Add("bntra", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                vdsthuoc.Tables[0].Columns.Add("bntra", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            dsdulieu.WriteXml("tam.xml", XmlWriteMode.WriteSchema);
            vdscls.WriteXml("cls.xml", XmlWriteMode.WriteSchema);
            vdsthuoc.WriteXml("thuoc.xml", XmlWriteMode.WriteSchema);
            DataSet vdsct = new DataSet();
            vdsct.Tables.Add(vdsthuoc.Tables[0].Copy());
            vdsct.Tables[0].TableName = "1";
            vdsct.Tables.Add(vdscls.Tables[0].Copy());
            string malk = "";
            string mabv = this._lib.Mabv;
            string tenbv = this._lib.Tenbv;
            DataSet set3 = this._lib.f_get_dsthuoc_cls();
            DataSet set4 = this._lib.f_get_dskhoaphong();
            int num2 = this._lib.iMavp_congkham(1);
            DataSet set5 = new DataSet();
            decimal num3 = 100M;
            set5.Tables.Add();
            set5.Tables[0].Columns.Add("id", typeof(decimal));
            set5.Tables[0].Columns.Add("ten");
            string str4 = "";
            string str5 = "";
            string str6 = "";
            string str7 = "";
            if (v_duongdanxuatxml == "")
            {
                v_duongdanxuatxml = this._s_xmlpath;
            }
            StreamWriter writer = new StreamWriter(v_duongdanxuatxml + "bhxh_" + denngay.Substring(3, 2) + denngay.Substring(8) + ".xml");
            if (!tachdsbenhnhan)
            {
                str4 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<GIAMDINHHS xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + this._s_arrCap[1] + "<THONGTINDONVI>" + this._s_arrCap[2] + "<MACSKCB>" + mabv + "</MACSKCB>" + this._s_arrCap[1] + "</THONGTINDONVI>" + this._s_arrCap[1] + "<THONGTINHOSO>" + this._s_arrCap[2] + "<NGAYLAP>" + ((ngaylaphoso == "") ? DateTime.Now.ToString("yyyyMMdd") : ngaylaphoso) + "</NGAYLAP>" + this._s_arrCap[2] + "<SOLUONGHOSO>" + dsdulieu.Tables[0].Rows.Count.ToString() + "</SOLUONGHOSO>" + this._s_arrCap[2] + "<DANHSACHHOSO>";
                writer.WriteLine(str4);
                str5 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                str5 = str5 + this._s_arrCap[0] + "<DSACH_CHI_TIET_THUOC>";
                str6 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                str6 = str6 + this._s_arrCap[0] + "<DSACH_CHI_TIET_DVKT>";
            }
            decimal d = 0M;
            decimal num5 = this._lib.themoi15_sotien();
            string exp = "";
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    exp = (((row["loaiba"].ToString() == "1") || (row["loaiba"].ToString() == "4")) ? ("mavaovien=" + row["mavaovien"].ToString()) : (" ngayduyet='" + row["ngay"].ToString().Substring(0, 10) + "' and sothe='" + row["sothe"].ToString() + "'")) + " and mabn='" + row["mabn"].ToString() + "'";
                    try
                    {
                        if (this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp) <= 0M)
                        {
                            this._lib.f_write_log("Benh nhan co tong chi phi bang 0: " + row["mabn"].ToString() + " - " + row["HOTEN"].ToString() + " so the: " + row["sothe"].ToString() + "  " + exp);
                            continue;
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        for (int j = 0; j < vdsct.Tables.Count; j++)
                        {
                            try
                            {
                                foreach (DataRow row2 in vdsct.Tables[j].Select(exp + " and sotien>gia_bh_toida  and gia_bh_toida>0", "soluong"))
                                {
                                    row2["sotien"] = row2["gia_bh_toida"].ToString();
                                    row2["dongia"] = decimal.Parse(row2["gia_bh_toida"].ToString()) / decimal.Parse(row2["soluong"].ToString());
                                }
                            }
                            catch
                            {
                            }
                            try
                            {
                                foreach (DataRow row3 in vdsct.Tables[j].Select(exp + " and soluong<0", "soluong"))
                                {
                                    this.f_update_soam_324moi(row3["soluong"].ToString(), exp + " and soluong>0 and mavp=" + row3["mavp"].ToString() + " and dongia=" + row3["dongia"].ToString(), ref vdsthuoc);
                                    row3["soluong"] = 0;
                                    row3["sotien"] = 0;
                                }
                            }
                            catch
                            {
                            }
                        }
                        vdsct.AcceptChanges();
                    }
                    catch (Exception exception)
                    {
                        this._lib.f_write_log("gia_bh_toida:" + exception.ToString());
                    }
                    DataRow row4 = dset.Tables[0].NewRow();
                    try
                    {
                        row4["ma_lk"] = this._lib.Mabv + row["malk"].ToString();
                    }
                    catch
                    {
                        row4["ma_lk"] = row["malk"].ToString();
                    }
                    row4["STT"] = d = decimal.op_Increment(d);
                    row4["ma_bn"] = row["mabn"].ToString();
                    row4["ho_ten"] = row["HOTEN"].ToString();
                    row4["ngay_sinh"] = row["ngaysinh"].ToString();
                    row4["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row4["dia_chi"] = row["diachi"].ToString();
                    try
                    {
                        row4["ma_the"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row4["ma_the"] = row["sothe"].ToString();
                    }
                    row4["ma_dkbd"] = row["MANOIDK"].ToString();
                    if (row["tungay"].ToString() == "")
                    {
                        try
                        {
                            DataSet set6 = this._lib.f_get_sothebhyt_tungay(row["sothe"].ToString(), row["NGAYVAO"].ToString(), row["NGAYRA"].ToString());
                            row["tungay"] = set6.Tables[0].Rows[0]["tungay"].ToString();
                            row["denngay"] = set6.Tables[0].Rows[0]["denngay"].ToString();
                        }
                        catch
                        {
                        }
                    }
                    try
                    {
                        row4["sobienlai"] = row["sobienlai"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row4["quyenso"] = row["quyenso"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row4["nguoithu"] = row["ten_thuvp"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row4["gt_the_tu"] = row["tungay"].ToString().Substring(6, 4) + row["tungay"].ToString().Substring(3, 2) + row["tungay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row4["gt_the_tu"] = "";
                    }
                    try
                    {
                        row4["gt_the_den"] = row["denngay"].ToString().Substring(6, 4) + row["denngay"].ToString().Substring(3, 2) + row["denngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row4["gt_the_den"] = "";
                    }
                    try
                    {
                        row4["ngay_ttoan"] = row["ngay"].ToString().Substring(6, 4) + row["ngay"].ToString().Substring(3, 2) + row["ngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row4["ngay_ttoan"] = "";
                    }
                    try
                    {
                        row4["ngay_ttoan"] = row4["ngay_ttoan"].ToString() + row["ngay"].ToString().Substring(11, 2) + row["ngay"].ToString().Substring(14, 2);
                    }
                    catch
                    {
                        row4["ngay_ttoan"] = row4["ngay_ttoan"].ToString() + "0000";
                    }
                    row4["ma_benh"] = row["MAICD"].ToString();
                    try
                    {
                        row4["ma_benh"] = row4["ma_benh"].ToString().Split(new char[] { ';' })[0];
                    }
                    catch
                    {
                    }
                    if (row4["ma_benh"].ToString() == "")
                    {
                        row4["ma_benh"] = row["MAICD"].ToString();
                    }
                    row4["ma_benhkhac"] = row["maicdkt"].ToString();
                    row4["ten_benh"] = row["chandoan"].ToString();
                    if (row["traituyen"].ToString() != "0")
                    {
                        row4["ma_lydo_vvien"] = 3;
                    }
                    else if (row["traituyen"].ToString() == "0")
                    {
                        row4["ma_lydo_vvien"] = 1;
                    }
                    else
                    {
                        row4["ma_lydo_vvien"] = 2;
                    }
                    row4["ma_noi_chuyen"] = "";
                    row4["ma_tai_nan"] = 0;
                    try
                    {
                        row4["ngay_vao"] = row["NGAYVV"].ToString();
                    }
                    catch
                    {
                        row4["ngay_vao"] = "";
                    }
                    try
                    {
                        row4["ngay_ra"] = row["NGAYRV"].ToString();
                    }
                    catch
                    {
                        row4["ngay_ra"] = "";
                    }
                    try
                    {
                        if (decimal.Parse(row4["ngay_ra"].ToString()) < decimal.Parse(row4["ngay_vao"].ToString()))
                        {
                            row4["ngay_vao"] = row4["ngay_ra"].ToString();
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        if (decimal.Parse(row4["ngay_ttoan"].ToString()) < decimal.Parse(row4["ngay_ra"].ToString()))
                        {
                            row4["ngay_ttoan"] = row4["ngay_ra"].ToString();
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        row4["so_ngay_dtri"] = row["SONGAY"].ToString();
                    }
                    catch
                    {
                        row4["so_ngay_dtri"] = 1;
                    }
                    row4["ket_qua_dtri"] = 1;
                    row4["tinh_trang_rv"] = 1;
                    try
                    {
                        row4["muc_huong"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                    }
                    catch
                    {
                    }
                    row4["t_thuoc"] = 0;
                    row4["t_vtyt"] = 0;
                    row4["t_tongchi"] = row["tongcong"].ToString();
                    row4["t_bntt"] = row["bntra"].ToString();
                    row4["t_bhtt"] = row["bhyttra"].ToString();
                    row4["t_nguonkhac"] = 0;
                    row4["t_ngoaids"] = 0;
                    row4["nam_qt"] = denngay.Substring(6, 4);
                    row4["thang_qt"] = denngay.Substring(3, 2);
                    row4["ma_loaikcb"] = 1;
                    if (row["loaiba"].ToString() == "1")
                    {
                        row4["ma_loaikcb"] = 3;
                    }
                    else if (row["loaiba"].ToString() == "2")
                    {
                        row4["ma_loaikcb"] = 2;
                    }
                    if (makp_ngoaitru != "")
                    {
                        try
                        {
                            if (makp_ngoaitru == "2")
                            {
                                row4["ma_loaikcb"] = 2;
                            }
                        }
                        catch
                        {
                        }
                    }
                    row4["ma_cskcb"] = this._lib.MABHXH;
                    try
                    {
                        row4["ma_khuvuc"] = row["makhuvuc"].ToString();
                    }
                    catch
                    {
                    }
                    row4["ma_PTTT_QT"] = "";
                    dset.Tables[0].Rows.Add(row4);
                    try
                    {
                        malk = row4["ma_lk"].ToString();
                        decimal num7 = 0M;
                        try
                        {
                            if (row["loaiba"].ToString() == "30")
                            {
                                num7 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", exp + "  and nhombhyt=9");
                                if (num7 == 0M)
                                {
                                    try
                                    {
                                        str4 = "id=" + num2.ToString();
                                        DataRow row5 = set3.Tables[0].Select(str4)[0];
                                        DataRow row6 = vdscls.Tables[0].NewRow();
                                        try
                                        {
                                            row6.ItemArray = vdscls.Tables[0].Select(exp)[0].ItemArray;
                                        }
                                        catch
                                        {
                                        }
                                        vdscls.Tables[0].Rows.Add(row6);
                                        try
                                        {
                                            row6["sothe"] = row4["ma_the"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                        row6["nhombhyt"] = 9;
                                        row6["mavp"] = row5["id"].ToString();
                                        row6["sotien"] = row["congkham"].ToString();
                                        row6["dongia"] = row["congkham"].ToString();
                                        row6["soluong"] = 1;
                                        row6["dvt"] = row5["dvt"].ToString();
                                        row6["mavp1"] = row5["ma"].ToString();
                                        row6["ten"] = row5["ten"].ToString();
                                        try
                                        {
                                            row6["malk"] = row["malk"].ToString();
                                        }
                                        catch
                                        {
                                            row6["malk"] = row["mavaovien"].ToString();
                                        }
                                        row6["mavaovien"] = row["mavaovien"].ToString();
                                        row6["ngayduyet"] = row["ngay"].ToString().Substring(0, 10);
                                        row6["mabn"] = row["mabn"].ToString();
                                        row6["ngayylenh"] = row["ngayvv"].ToString();
                                        try
                                        {
                                            row6["maql"] = row["maql"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                    catch (Exception exception2)
                                    {
                                        this._lib.f_write_log("them cong kham =0: " + exception2.ToString());
                                    }
                                    vdscls.AcceptChanges();
                                }
                                else
                                {
                                    try
                                    {
                                        str4 = "id=" + num2.ToString();
                                        DataRow row7 = set3.Tables[0].Select(str4)[0];
                                        DataRow[] rowArray = vdscls.Tables[0].Select(exp + "  and nhombhyt=9");
                                        try
                                        {
                                            rowArray[0]["sotien"] = row["congkham"].ToString();
                                        }
                                        catch
                                        {
                                        }
                                        rowArray[0]["dongia"] = rowArray[0]["sotien"].ToString();
                                        rowArray[0]["soluong"] = 1;
                                        for (int num8 = 1; num8 < rowArray.Length; num8++)
                                        {
                                            rowArray[num8]["sotien"] = 0;
                                        }
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                        foreach (DataRow row8 in vdsct.Tables[1].Select(exp + " and nhombhyt=110", "soluong"))
                        {
                            try
                            {
                                row8["bhyt"] = 100;
                                if (row8["soluong"].ToString() == "0.5")
                                {
                                    row8["soluong"] = 1;
                                    row8["bhyt"] = 50;
                                    row8["sotien"] = decimal.Parse(row8["dongia"].ToString()) / 2M;
                                }
                                else if (row8["soluong"].ToString().IndexOf(".5") > -1)
                                {
                                    DataRow row9 = vdsct.Tables[1].NewRow();
                                    row9.ItemArray = row8.ItemArray;
                                    row9["soluong"] = 1;
                                    row9["bhyt"] = 50;
                                    row9["sotien"] = decimal.Parse(row8["dongia"].ToString()) / 2M;
                                    row9["bhyttra"] = 0;
                                    row9["bntra"] = decimal.Parse(row9["sotien"].ToString()) - decimal.Parse(row9["bhyttra"].ToString());
                                    vdsct.Tables[1].Rows.Add(row9);
                                    row8["soluong"] = double.Parse(row8["soluong"].ToString()) - 0.5;
                                    row8["sotien"] = decimal.Parse(row8["soluong"].ToString()) * decimal.Parse(row8["dongia"].ToString());
                                }
                                row8["bntra"] = decimal.Parse(row8["sotien"].ToString()) - decimal.Parse(row8["bhyttra"].ToString());
                            }
                            catch
                            {
                            }
                        }
                        vdsct.AcceptChanges();
                        try
                        {
                            num7 = this.f_tinhtong_cttt9324(vdsct, "sotien", exp + " and nhombhyt in(14,7)");
                            if (num7 > 48400000M)
                            {
                                row4["t_bntt"] = decimal.Parse(row4["t_bntt"].ToString()) + (num7 - 48400000M);
                                row4["t_bhtt"] = decimal.Parse(row4["t_tongct"].ToString()) - decimal.Parse(row4["t_bntt"].ToString());
                            }
                        }
                        catch
                        {
                        }
                        DataRow row10 = null;
                        for (int k = 0; k < vdsct.Tables.Count; k++)
                        {
                            foreach (DataRow row11 in vdsct.Tables[k].Select(exp + " and nhombhyt=10"))
                            {
                                if (!(row["traituyen"].ToString() == "1") || (row["sothe"].ToString().IndexOf("TE1") != -1))
                                {
                                    row11["bhyttra"] = row11["sotien"].ToString();
                                }
                            }
                            foreach (DataRow row12 in vdsct.Tables[k].Select(exp + " and nhombhyt<>10"))
                            {
                                row10 = set3.Tables[0].Select("id=" + row12["mavp"].ToString() + "")[0];
                                string str9 = "";
                                try
                                {
                                    str9 = row10["bhyt"].ToString();
                                    str9 = (str9 == "0") ? "100" : str9;
                                    try
                                    {
                                        if (row12["nhombhyt"].ToString() == "9")
                                        {
                                            try
                                            {
                                                decimal num10 = decimal.Parse(vdscls.Tables[0].Select(exp + " and nhombhyt=9", "")[0]["congkham"].ToString());
                                                if (num10 == 0M)
                                                {
                                                    num10 = decimal.Parse(row12["dongia"].ToString());
                                                }
                                                str9 = ((100 * int.Parse(row12["dongia"].ToString())) / num10);
                                            }
                                            catch
                                            {
                                                continue;
                                            }
                                        }
                                    }
                                    catch
                                    {
                                    }
                                    row12["bhyt"] = str9;
                                }
                                catch
                                {
                                }
                                if (row["loaiba"].ToString() == "3")
                                {
                                    try
                                    {
                                        row12["bhyttra"] = (decimal.Parse(row12["sotien"].ToString()) * decimal.Parse(row["tylebhyt"].ToString())) / 100M;
                                    }
                                    catch
                                    {
                                    }
                                }
                                row12["sotien"] = decimal.Parse(row12["sotien"].ToString()).ToString("###.00");
                                row12["bhyttra"] = decimal.Parse(row12["bhyttra"].ToString()).ToString("###.00");
                                row12["bntra"] = decimal.Parse(row12["sotien"].ToString()) - decimal.Parse(row12["bhyttra"].ToString());
                            }
                        }
                        vdsct.AcceptChanges();
                        try
                        {
                            num7 = this.f_tinhtong_cttt9324(vdsct, "sotien", exp + " and nhombhyt in(3,4) and sotien>0");
                            num7 = Math.Round(num7, 2);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_tongchi"] = Math.Round(Convert.ToDecimal(row4["t_tongchi"].ToString()), 2);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_vtyt"] = this.f_tinhtong_cttt9324(vdsct, "sotien", exp + " and nhombhyt in(6,14) and sotien>0");
                            row4["t_vtyt"] = Math.Round(Convert.ToDouble(row4["t_vtyt"].ToString()), 2);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_tongct"] = this.f_tinhtong_cttt9324(vdsct, "sotien", exp + " and sotien>0 and nhombhyt not in(3,4,6,14)");
                            row4["t_tongct"] = (Math.Round(Convert.ToDecimal(row4["t_tongct"].ToString()), 2) + num7) + Convert.ToDecimal(row4["t_vtyt"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_bhtt"] = this.f_tinhtong_cttt9324(vdsct, "bhyttra", exp + " and sotien>0");
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_bntt"] = decimal.Parse(row4["t_tongct"].ToString()) - decimal.Parse(row4["t_bhtt"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            row4["t_thuoc"] = Math.Round(num7, 2);
                        }
                        catch
                        {
                        }
                        if (Math.Abs((double) (double.Parse(row4["t_bhtt"].ToString()) - double.Parse(row4["t_tongct"].ToString()))) < 5.0)
                        {
                            row4["muc_huong"] = 100;
                        }
                        num3 = 100M;
                        if (row["traituyen"].ToString() != "0")
                        {
                            num3 = 60M;
                        }
                        try
                        {
                            str7 = set4.Tables[0].Select("makp='" + row["makp"].ToString() + "'")[0]["makp_byt"].ToString();
                        }
                        catch
                        {
                            str7 = "";
                        }
                        string ngaythanhtoan = row4["ngay_ttoan"].ToString();
                        try
                        {
                            if (decimal.Parse(row4["ngay_ttoan"].ToString()) < decimal.Parse(row4["ngay_ra"].ToString()))
                            {
                                ngaythanhtoan = row4["ngay_ra"].ToString();
                            }
                        }
                        catch
                        {
                        }
                        try
                        {
                            if (tachdsbenhnhan)
                            {
                                writer = new StreamWriter(this._s_xmlpath + "bhxh_" + row4["ma_the"].ToString() + "_" + row["ngayrv"].ToString() + ".xml");
                                str4 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<GIAMDINHHS xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + this._s_arrCap[1] + "<THONGTINDONVI>" + this._s_arrCap[2] + "<MACSKCB>" + mabv + "</MACSKCB>" + this._s_arrCap[1] + "</THONGTINDONVI>" + this._s_arrCap[1] + "<THONGTINHOSO>" + this._s_arrCap[2] + "<NGAYLAP>" + ((ngaylaphoso == "") ? DateTime.Now.ToString("yyyyMMdd") : ngaylaphoso) + "</NGAYLAP>" + this._s_arrCap[2] + "<SOLUONGHOSO>1</SOLUONGHOSO>" + this._s_arrCap[2] + "<DANHSACHHOSO>";
                                writer.WriteLine(str4);
                                str5 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                                str5 = str5 + this._s_arrCap[0] + "<DSACH_CHI_TIET_THUOC>";
                                str6 = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                                str6 = str6 + this._s_arrCap[0] + "<DSACH_CHI_TIET_DVKT>";
                            }
                            str4 = this._s_arrCap[3] + "<HOSO>" + this._s_arrCap[4] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML1</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>";
                            writer.WriteLine(str4);
                            str4 = this._s_arrCap[0] + "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                            str4 = str4 + this.f_nodexml_tonghopttbn_4210(1, malk, row4["STT"].ToString(), row4["ma_bn"].ToString(), !tt324moi ? row4["ho_ten"].ToString() : this.f_cdata_tag(row4["ho_ten"].ToString().ToLower()), row4["ngay_sinh"].ToString(), row4["gioi_tinh"].ToString(), !tt324moi ? row4["dia_chi"].ToString() : this.f_cdata_tag(row4["dia_chi"].ToString()), row4["ma_the"].ToString(), row4["ma_dkbd"].ToString(), row4["gt_the_tu"].ToString(), row4["gt_the_den"].ToString(), row4["ma_benh"].ToString(), !tt324moi ? row4["ten_benh"].ToString() : this.f_cdata_tag(row4["ten_benh"].ToString()), row4["ma_lydo_vvien"].ToString(), "", "0", row4["ngay_vao"].ToString(), row4["ngay_ra"].ToString(), (row4["so_ngay_dtri"].ToString() == "0") ? "1" : row4["so_ngay_dtri"].ToString(), row4["ket_qua_dtri"].ToString(), row4["tinh_trang_rv"].ToString(), ngaythanhtoan, row4["muc_huong"].ToString(), row4["t_thuoc"].ToString(), row4["t_vtyt"].ToString(), row4["t_tongct"].ToString(), "0", row4["t_bhtt"].ToString(), "0", denngay.Substring(6, 4), denngay.Substring(3, 2), Convert.ToInt32(row4["ma_loaikcb"].ToString()), (str7 == "") ? "000" : str7, mabv, "", "0", true, "", row4["t_bntt"].ToString());
                        }
                        catch
                        {
                        }
                        str4 = mahoadulieu ? this.f_convert_to_unicode(str4) : str4;
                        string str34 = str4;
                        str4 = str34 + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                        writer.WriteLine(str4);
                        string strText = "";
                        int num11 = 0;
                        for (int m = 0; m < vdsct.Tables.Count; m++)
                        {
                            foreach (DataRow row13 in vdsct.Tables[m].Select(exp + " and nhombhyt in(3,4) and soluong>0", "nhombhyt"))
                            {
                                string str = "";
                                string duongdung = "";
                                string mathuoc = "";
                                string manhom = "";
                                string ngayylenh = "";
                                string str17 = "";
                                string str18 = "100";
                                string str19 = "";
                                string str20 = "";
                                row10 = set3.Tables[0].Select("id=" + row13["mavp"].ToString() + "")[0];
                                try
                                {
                                    str = row10["tenduongdung"].ToString() + " " + row13["soluong"].ToString() + " " + row13["dvt"].ToString() + " theo toa thuốc.";
                                }
                                catch
                                {
                                    str = "0000";
                                }
                                try
                                {
                                    mathuoc = row10["mabyt"].ToString();
                                }
                                catch
                                {
                                }
                                if (mathuoc == "")
                                {
                                    mathuoc = row13["mavp1"].ToString();
                                }
                                try
                                {
                                    str19 = set4.Tables[0].Select("makp='" + row13["makp"].ToString() + "'")[0]["makp_byt"].ToString();
                                }
                                catch
                                {
                                    str19 = row13["makp"].ToString();
                                }
                                try
                                {
                                    manhom = row10["maloaibhyt"].ToString();
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str17 = row10["tenbyt"].ToString();
                                    if (str17 == "")
                                    {
                                        str17 = row13["ten"].ToString();
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str20 = row10["nhathau"].ToString();
                                }
                                catch
                                {
                                }
                                try
                                {
                                    duongdung = row10["duongdung"].ToString();
                                    if (duongdung == "")
                                    {
                                        duongdung = "0000";
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    string str21 = row10["dvtbyt"].ToString();
                                    if (str21 != "")
                                    {
                                        row13["dvt"] = str21;
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    ngayylenh = row13["ngayylenh"].ToString().Substring(6, 4) + row13["ngayylenh"].ToString().Substring(3, 2) + row13["ngayylenh"].ToString().Substring(0, 2);
                                    if (row13["ngayylenh"].ToString().Length > 10)
                                    {
                                        ngayylenh = ngayylenh + row13["ngayylenh"].ToString().Substring(11, 2) + row13["ngayylenh"].ToString().Substring(14, 2);
                                    }
                                }
                                catch
                                {
                                    ngayylenh = row["ngayvv"].ToString();
                                }
                                try
                                {
                                    if (((manhom == "") || (mathuoc == "")) && (set5.Tables[0].Select("id=" + row13["mavp"].ToString()).Length == 0))
                                    {
                                        set5.Tables[0].Rows.Add(new object[] { decimal.Parse(row13["mavp"].ToString()), row13["ten"].ToString() });
                                    }
                                    row4["t_" + manhom] = decimal.Parse(row4["t_" + manhom].ToString()) + decimal.Parse(row13["sotien"].ToString());
                                }
                                catch
                                {
                                }
                                strText = strText + this.f_nodexml_chitietthuoc_4210(1, malk, ++num11, mathuoc, manhom, !tt324moi ? str17 : this.f_cdata_tag(str17), (row13["dvt"].ToString() == "") ? "0000" : row13["dvt"].ToString(), !tt324moi ? row13["hamluong"].ToString() : this.f_cdata_tag(row13["hamluong"].ToString()), duongdung, !tt324moi ? str : this.f_cdata_tag(str), (row13["sodk"].ToString() == "") ? "0000" : row13["sodk"].ToString(), Math.Round(Convert.ToDecimal(row13["soluong"].ToString()), 2), Math.Round(Convert.ToDecimal(row13["dongia"].ToString()), 2), Convert.ToDecimal(str18), Math.Round(Convert.ToDecimal(row13["sotien"].ToString()), 2), (str19 == "") ? "000" : str19, (row13["kihieubs"].ToString() == "") ? "0000" : row13["kihieubs"].ToString(), row4["ma_benh"].ToString(), ngayylenh, "0", (str20 == "") ? "0000" : str20, 1, Convert.ToInt32((int) ((Convert.ToInt32(str18) * Convert.ToInt32(row4["muc_huong"].ToString())) / 100)), 0M, Math.Round(Convert.ToDecimal(row13["bhyttra"].ToString()), 2), Convert.ToDecimal(row13["sotien"].ToString()) - Math.Round(Convert.ToDecimal(row13["bhyttra"].ToString()), 2));
                            }
                        }
                        if (strText != "")
                        {
                            strText = (this._s_arrCap[0] + str5 + strText) + this._s_arrCap[0] + "</DSACH_CHI_TIET_THUOC>";
                            str4 = this._s_arrCap[3] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML2</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>" + (mahoadulieu ? this.f_convert_to_unicode(strText) : strText) + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                            writer.WriteLine(str4);
                        }
                        strText = "";
                        for (int n = 0; n < vdsct.Tables.Count; n++)
                        {
                            foreach (DataRow row14 in vdsct.Tables[n].Select(exp + " and nhombhyt not in(3,4) and sotien>0", "nhombhyt"))
                            {
                                string mavattu = "";
                                string str23 = "";
                                string str24 = "";
                                string ma = "";
                                string str26 = "";
                                string str27 = "";
                                string str28 = "100";
                                string str29 = "";
                                string str30 = "";
                                string magiuong = "";
                                row10 = set3.Tables[0].Select("id=" + row14["mavp"].ToString() + "")[0];
                                try
                                {
                                    if ((row14["nhombhyt"].ToString() == "6") || (row14["nhombhyt"].ToString() == "14"))
                                    {
                                        mavattu = row10["mabyt"].ToString();
                                        str23 = row10["tenbyt"].ToString();
                                        if (row14["nhombhyt"].ToString() == "14")
                                        {
                                            try
                                            {
                                                string str32 = vdsthuoc.Tables[0].Select(exp + " and nhombhyt in(7) and sotien>0", "nhombhyt")[0]["mavp"].ToString();
                                                ma = set3.Tables[0].Select("id=" + str32 + "")[0]["mabyt"].ToString();
                                            }
                                            catch
                                            {
                                            }
                                        }
                                    }
                                    else
                                    {
                                        ma = row10["mabyt"].ToString();
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str29 = set4.Tables[0].Select("makp='" + row14["makp"].ToString() + "'")[0]["makp_byt"].ToString();
                                }
                                catch
                                {
                                    str29 = row14["makp"].ToString();
                                }
                                try
                                {
                                    str24 = row10["maloaibhyt"].ToString();
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str30 = row10["nhathau"].ToString();
                                    if (row14["nhombhyt"].ToString() == "11")
                                    {
                                        magiuong = row10["ma_giuong_byt"].ToString();
                                        magiuong = (magiuong == "") ? row10["mabyt"].ToString() : magiuong;
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str27 = row10["tenbyt"].ToString();
                                    if (str27 == "")
                                    {
                                        str27 = row14["ten"].ToString();
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    string str33 = row10["dvtbyt"].ToString();
                                    if (str33 != "")
                                    {
                                        row14["dvt"] = str33;
                                    }
                                }
                                catch
                                {
                                }
                                try
                                {
                                    str26 = row14["ngayylenh"].ToString().Substring(6, 4) + row14["ngayylenh"].ToString().Substring(3, 2) + row14["ngayylenh"].ToString().Substring(0, 2);
                                    if (row14["ngayylenh"].ToString().Length > 10)
                                    {
                                        str26 = str26 + row14["ngayylenh"].ToString().Substring(11, 2) + row14["ngayylenh"].ToString().Substring(14, 2);
                                    }
                                }
                                catch
                                {
                                    str26 = row["ngayvv"].ToString();
                                }
                                try
                                {
                                    if (((str24 == "") || ((ma == "") && (mavattu == ""))) && (set5.Tables[0].Select("id=" + row14["mavp"].ToString()).Length == 0))
                                    {
                                        set5.Tables[0].Rows.Add(new object[] { decimal.Parse(row14["mavp"].ToString()), row14["ten"].ToString() });
                                    }
                                    row4["t_" + str24] = decimal.Parse(row4["t_" + str24].ToString()) + decimal.Parse(row14["sotien"].ToString());
                                }
                                catch
                                {
                                }
                                strText = strText + this.f_nodexml_chitietcls_4210(1, malk, ++num11, ma, mavattu, str24, !tt324moi ? str27 : this.f_cdata_tag(str27), (row14["dvt"].ToString() == "") ? "0000" : row14["dvt"].ToString(), Math.Round(Convert.ToDecimal(row14["soluong"].ToString()), 2), Math.Round(Convert.ToDecimal(row14["dongia"].ToString()), 2), Convert.ToDecimal(str28), Math.Round(Convert.ToDecimal(row14["sotien"].ToString()), 2), (str29 == "") ? "000" : str29, (row14["kihieubs"].ToString() == "") ? "0000" : row14["kihieubs"].ToString(), row4["ma_benh"].ToString(), str26, str26, "0", "", str23, 1, (str30 == "") ? "0000.00.0" : str30, (row14["gia_bh_toida"].ToString() == "0") ? "" : row14["gia_bh_toida"].ToString(), 0M, decimal.Parse(row14["bhyttra"].ToString()), Convert.ToDecimal(row14["sotien"].ToString()) - Math.Round(Convert.ToDecimal(row14["bhyttra"].ToString()), 2), magiuong, Convert.ToInt32((int) ((Convert.ToInt32(str28) * Convert.ToInt32(row4["muc_huong"].ToString())) / 100)));
                            }
                        }
                        if (strText != "")
                        {
                            strText = (this._s_arrCap[0] + str6 + strText) + this._s_arrCap[0] + "</DSACH_CHI_TIET_DVKT>";
                            str4 = this._s_arrCap[3] + "<FILEHOSO>" + this._s_arrCap[4] + "<LOAIHOSO>XML3</LOAIHOSO>" + this._s_arrCap[4] + "<NOIDUNGFILE>" + (mahoadulieu ? this.f_convert_to_unicode(strText) : strText) + this._s_arrCap[4] + "</NOIDUNGFILE>" + this._s_arrCap[3] + "</FILEHOSO>";
                            writer.WriteLine(str4);
                        }
                        writer.WriteLine(this._s_arrCap[3] + "</HOSO>");
                        if (tachdsbenhnhan)
                        {
                            str4 = this._s_arrCap[2] + "</DANHSACHHOSO>" + this._s_arrCap[1] + "</THONGTINHOSO>" + this._s_arrCap[1] + "<CHUKYDONVI />\r\n</GIAMDINHHS>";
                            writer.WriteLine(str4);
                            writer.Close();
                        }
                    }
                    catch (Exception exception3)
                    {
                        this._lib.f_write_log(row["mabn"].ToString() + " - " + row["HOTEN"].ToString() + " so the: " + row["sothe"].ToString() + "  " + exception3.ToString());
                    }
                }
                catch (Exception exception4)
                {
                    MessageBox.Show(exception4.Message, this._lib.Msg);
                }
            }
            if (!tachdsbenhnhan)
            {
                str4 = this._s_arrCap[2] + "</DANHSACHHOSO>" + this._s_arrCap[1] + "</THONGTINHOSO>" + this._s_arrCap[1] + "<CHUKYDONVI />\r\n</GIAMDINHHS>";
                writer.WriteLine(str4);
                writer.Close();
            }
            dset.WriteXml("bang01_ngoai.xml", XmlWriteMode.WriteSchema);
            set5.WriteXml("dsmanhomdichvu.xml", XmlWriteMode.WriteSchema);
            if (xuatfileexcel)
            {
                int num14 = 0;
                int num15 = 0;
                int num16 = 0;
                num14 = dset.Tables[0].Rows.Count + 5;
                num15 = dset.Tables[0].Columns.Count - 1;
                num16 = num14;
                this.tenfile = this._lib.Export_Excel(dset, "bccpkcb01");
                try
                {
                    Process.Start(this.tenfile);
                }
                catch
                {
                }
            }
        }

        public void f_Ngoaitru_xuatExcel_mau01_TT9324(bool print, DataSet dsdulieu, string tungay, string denngay, DataSet vdsthuoc, DataSet vdscls)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("ma_lk", typeof(string));
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
            dset.Tables[0].Columns.Add("ten_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_benhkhac", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ma_tai_nan", typeof(int));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("ngay_ttoan", typeof(string));
            dset.Tables[0].Columns.Add("muc_huong", typeof(decimal));
            dset.Tables[0].Columns.Add("t_thuoc", typeof(decimal));
            dset.Tables[0].Columns.Add("t_vtyt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_nguonkhac", typeof(decimal));
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal));
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_PTTT_QT", typeof(string));
            dset.Tables[0].Columns.Add("can_nang", typeof(decimal)).DefaultValue = 0;
            dset.Tables[0].Columns.Add("t_tongct", typeof(decimal)).DefaultValue = 0;
            dsdulieu.WriteXml("tam.xml", XmlWriteMode.WriteSchema);
            vdscls.WriteXml("cls.xml", XmlWriteMode.WriteSchema);
            vdsthuoc.WriteXml("thuoc.xml", XmlWriteMode.WriteSchema);
            string str = "";
            string malk = "";
            string mabv = this._lib.Mabv;
            DataSet set2 = this._lib.f_get_dsthuoc_cls();
            int num = this._lib.iMavp_congkham(1);
            decimal d = 0M;
            decimal num3 = this._lib.themoi15_sotien();
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    try
                    {
                        if (this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", " mavaovien=" + row["mavaovien"].ToString()) <= 0M)
                        {
                            continue;
                        }
                    }
                    catch
                    {
                    }
                    DataRow row2 = dset.Tables[0].NewRow();
                    try
                    {
                        row2["ma_lk"] = this._lib.Mabv + row["malk"].ToString();
                    }
                    catch
                    {
                        row2["ma_lk"] = row["malk"].ToString();
                    }
                    row2["STT"] = d = decimal.op_Increment(d);
                    row2["ma_bn"] = row["mabn"].ToString();
                    row2["ho_ten"] = row["HOTEN"].ToString();
                    row2["ngay_sinh"] = row["ngaysinh"].ToString();
                    row2["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row2["dia_chi"] = row["diachi"].ToString();
                    try
                    {
                        row2["ma_the"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row2["ma_the"] = row["sothe"].ToString();
                    }
                    row2["ma_dkbd"] = row["MANOIDK"].ToString();
                    try
                    {
                        row2["gt_the_tu"] = row["tungay"].ToString().Substring(6, 4) + row["tungay"].ToString().Substring(3, 2) + row["tungay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_tu"] = "";
                    }
                    try
                    {
                        row2["gt_the_den"] = row["denngay"].ToString().Substring(6, 4) + row["denngay"].ToString().Substring(3, 2) + row["denngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_den"] = "";
                    }
                    try
                    {
                        row2["ngay_ttoan"] = row["ngay"].ToString().Substring(6, 4) + row["ngay"].ToString().Substring(3, 2) + row["ngay"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_ttoan"] = "";
                    }
                    row2["ma_benh"] = row["MAICD"].ToString();
                    try
                    {
                        row2["ma_benh"] = row2["ma_benh"].ToString().Split(new char[] { ';' })[0];
                    }
                    catch
                    {
                    }
                    row2["ma_benhkhac"] = row["maicdkt"].ToString();
                    row2["ten_benh"] = row["chandoan"].ToString();
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
                    row2["ma_tai_nan"] = 0;
                    try
                    {
                        row2["ngay_vao"] = row["NGAYVAO"].ToString().Substring(6, 4) + row["NGAYVAO"].ToString().Substring(3, 2) + row["NGAYVAO"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_vao"] = "";
                    }
                    try
                    {
                        row2["ngay_ra"] = row["NGAYRA"].ToString().Substring(6, 4) + row["NGAYRA"].ToString().Substring(3, 2) + row["NGAYRA"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_ra"] = "";
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
                    try
                    {
                        row2["muc_huong"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                    }
                    catch
                    {
                    }
                    row2["t_thuoc"] = 0;
                    row2["t_vtyt"] = 0;
                    row2["t_tongchi"] = row["tongcong"].ToString();
                    row2["t_bntt"] = row["bntra"].ToString();
                    row2["t_bhtt"] = row["bhyttra"].ToString();
                    row2["t_nguonkhac"] = 0;
                    row2["t_ngoaids"] = 0;
                    row2["nam_qt"] = denngay.Substring(6, 4);
                    row2["thang_qt"] = denngay.Substring(3, 2);
                    row2["ma_loaikcb"] = 1;
                    row2["ma_cskcb"] = this._lib.MABHXH;
                    try
                    {
                        row2["ma_khuvuc"] = row["makhuvuc"].ToString();
                    }
                    catch
                    {
                    }
                    row2["ma_PTTT_QT"] = "";
                    dset.Tables[0].Rows.Add(row2);
                    try
                    {
                        StreamWriter writer = new StreamWriter(this._s_xmlpath + row2["ngay_ra"].ToString() + "_" + row2["ma_the"].ToString() + "_CheckOut.xml");
                        str = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>\r\n<CHECKOUT>";
                        malk = row2["ma_lk"].ToString();
                        try
                        {
                            str = str + this.f_nodexml_thongtinbn(1, malk, "0", row2["ngay_vao"].ToString(), row2["ngay_ra"].ToString(), mabv, row2["ma_benh"].ToString(), "1", "", "", "");
                            writer.WriteLine(str);
                        }
                        catch
                        {
                        }
                        decimal num5 = 0M;
                        try
                        {
                            if (row["loaiba"].ToString() == "3")
                            {
                                num5 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", " mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt=9");
                                if (num5 == 0M)
                                {
                                    try
                                    {
                                        str = "id=" + num.ToString();
                                        DataRow row3 = set2.Tables[0].Select(str)[0];
                                        DataRow row4 = vdscls.Tables[0].NewRow();
                                        row4["nhombhyt"] = 9;
                                        row4["mavp"] = row3["id"].ToString();
                                        row4["sotien"] = 0x4268;
                                        row4["dongia"] = 0x4268;
                                        row4["soluong"] = 1;
                                        row4["dvt"] = row3["dvt"].ToString();
                                        row4["mavp1"] = row3["ma"].ToString();
                                        row4["ten"] = row3["ten"].ToString();
                                        row4["malk"] = row["malk"].ToString();
                                        row4["mavaovien"] = row["mavaovien"].ToString();
                                        vdscls.Tables[0].Rows.Add(row4);
                                    }
                                    catch
                                    {
                                    }
                                }
                                foreach (DataRow row5 in vdsthuoc.Tables[0].Select("mavaovien=" + row["mavaovien"].ToString()))
                                {
                                    try
                                    {
                                        row5["bhyttra"] = (decimal.Parse(row5["sotien"].ToString()) * decimal.Parse(row["tylebhyt"].ToString())) / 100M;
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        row5["bntra"] = decimal.Parse(row5["sotien"].ToString()) - decimal.Parse(row5["bhyttra"].ToString());
                                    }
                                    catch
                                    {
                                    }
                                }
                                vdsthuoc.AcceptChanges();
                                foreach (DataRow row6 in vdscls.Tables[0].Select("mavaovien=" + row["mavaovien"].ToString()))
                                {
                                    try
                                    {
                                        row6["bhyttra"] = (decimal.Parse(row6["sotien"].ToString()) * decimal.Parse(row["tylebhyt"].ToString())) / 100M;
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        row6["bntra"] = decimal.Parse(row6["sotien"].ToString()) - decimal.Parse(row6["bhyttra"].ToString());
                                    }
                                    catch
                                    {
                                    }
                                }
                                vdscls.AcceptChanges();
                            }
                        }
                        catch
                        {
                        }
                        try
                        {
                            num5 = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", " mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt=3");
                            num5 = Math.Round(num5);
                        }
                        catch
                        {
                        }
                        try
                        {
                            row2["t_tongchi"] = Math.Round(Convert.ToDecimal(row2["t_tongchi"].ToString()));
                        }
                        catch
                        {
                        }
                        try
                        {
                            row2["t_vtyt"] = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", " mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt=6");
                        }
                        catch
                        {
                        }
                        try
                        {
                            row2["t_tongct"] = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "sotien", " mavaovien=" + row["mavaovien"].ToString());
                        }
                        catch
                        {
                        }
                        try
                        {
                            row2["t_bhtt"] = this.f_tinhtong_cttt9324(vdsthuoc, vdscls, "bhyttra", " mavaovien=" + row["mavaovien"].ToString());
                            row2["t_bhtt"] = Math.Round(Convert.ToDecimal(row2["t_bhtt"].ToString()));
                        }
                        catch
                        {
                        }
                        try
                        {
                            row2["t_bntt"] = Convert.ToDecimal(row2["t_tongchi"].ToString()) - Convert.ToDecimal(row2["t_bhtt"].ToString());
                        }
                        catch
                        {
                        }
                        row2["t_thuoc"] = num5.ToString();
                        try
                        {
                            str = this._s_arrCap[1] + "<THONGTINCHITIET>";
                            str = str + this.f_nodexml_tonghopttbn(2, malk, row2["STT"].ToString(), row2["ma_bn"].ToString(), row2["ho_ten"].ToString(), row2["ngay_sinh"].ToString(), row2["gioi_tinh"].ToString(), row2["dia_chi"].ToString(), row2["ma_the"].ToString(), row2["ma_dkbd"].ToString(), row2["gt_the_tu"].ToString(), row2["gt_the_den"].ToString(), row2["ma_benh"].ToString(), row2["ten_benh"].ToString(), row2["ma_lydo_vvien"].ToString(), "", "0", row2["ngay_vao"].ToString(), row2["ngay_ra"].ToString(), (row2["so_ngay_dtri"].ToString() == "0") ? "1" : row2["so_ngay_dtri"].ToString(), row2["ket_qua_dtri"].ToString(), row2["tinh_trang_rv"].ToString(), row2["ngay_ttoan"].ToString(), row2["muc_huong"].ToString(), Math.Round(num5), "0", row2["t_tongchi"].ToString(), row2["t_bntt"].ToString(), row2["t_bhtt"].ToString(), "0", denngay.Substring(6, 4), denngay.Substring(3, 2), Convert.ToInt32(row2["ma_loaikcb"].ToString()), "", mabv, "", "0", false);
                        }
                        catch
                        {
                        }
                        writer.WriteLine(str);
                        string str5 = this._s_arrCap[2] + "<BANG_CTTHUOC>";
                        int num6 = 0;
                        foreach (DataRow row7 in vdsthuoc.Tables[0].Select("mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt=3 and sotien>0"))
                        {
                            string lieudung = "";
                            string mathuoc = "";
                            string manhom = "";
                            try
                            {
                                lieudung = set2.Tables[0].Select("mavp='" + row7["mavp1"].ToString() + "'")[0]["lieuluong"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                mathuoc = set2.Tables[0].Select("mavp='" + row7["mavp1"].ToString() + "'")[0]["mabyt"].ToString();
                            }
                            catch
                            {
                            }
                            if (mathuoc == "")
                            {
                                mathuoc = row7["mavp1"].ToString();
                            }
                            try
                            {
                                manhom = set2.Tables[0].Select("mavp='" + row7["mavp1"].ToString() + "'")[0]["manhombhyt"].ToString();
                            }
                            catch
                            {
                            }
                            str5 = str5 + this.f_nodexml_chitietthuoc(3, malk, ++num6, mathuoc, manhom, row7["ten"].ToString(), row7["dvt"].ToString(), row7["hamluong"].ToString(), row7["duongdung"].ToString(), lieudung, row7["sodk"].ToString(), Convert.ToDecimal(row7["soluong"].ToString()), Math.Round(Convert.ToDecimal(row7["dongia"].ToString())), Convert.ToDecimal(row2["muc_huong"].ToString()), Math.Round((decimal) ((Convert.ToDecimal(row7["sotien"].ToString()) * Convert.ToDecimal(row2["muc_huong"].ToString())) / 100M)), "", row7["kihieubs"].ToString(), row2["ma_benh"].ToString(), "", "0");
                        }
                        str5 = str5 + this._s_arrCap[2] + "</BANG_CTTHUOC>";
                        writer.WriteLine(str5);
                        str5 = this._s_arrCap[2] + "<BANG_CTDV>";
                        foreach (DataRow row8 in vdsthuoc.Tables[0].Select("mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt<>3 and sotien>0"))
                        {
                            string mavattu = "";
                            string str10 = "";
                            string ma = "";
                            try
                            {
                                mavattu = set2.Tables[0].Select("mavp='" + row8["mavp1"].ToString() + "'")[0]["mavattu"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                ma = set2.Tables[0].Select("mavp='" + row8["mavp1"].ToString() + "'")[0]["mabyt"].ToString();
                            }
                            catch
                            {
                            }
                            if (ma == "")
                            {
                                ma = row8["mavp1"].ToString();
                            }
                            try
                            {
                                str10 = set2.Tables[0].Select("mavp='" + row8["mavp1"].ToString() + "'")[0]["manhombhyt"].ToString();
                            }
                            catch
                            {
                            }
                            str5 = str5 + this.f_nodexml_chitietcls(3, malk, ++num6, ma, mavattu, str10, row8["ten"].ToString(), row8["dvt"].ToString(), Convert.ToDecimal(row8["soluong"].ToString()), Math.Round(Convert.ToDecimal(row8["dongia"].ToString())), Convert.ToDecimal(row2["muc_huong"].ToString()), Math.Round((decimal) ((Convert.ToDecimal(row8["sotien"].ToString()) * Convert.ToDecimal(row2["muc_huong"].ToString())) / 100M)), "", row8["kihieubs"].ToString(), row2["ma_benh"].ToString(), "", "", "0");
                        }
                        foreach (DataRow row9 in vdscls.Tables[0].Select("mavaovien=" + row["mavaovien"].ToString() + " and nhombhyt<>3 and sotien>0"))
                        {
                            string str12 = "";
                            string str13 = "";
                            string str14 = "";
                            try
                            {
                                str12 = set2.Tables[0].Select("mavp='" + row9["mavp1"].ToString() + "'")[0]["mavattu"].ToString();
                            }
                            catch
                            {
                            }
                            try
                            {
                                str14 = set2.Tables[0].Select("mavp='" + row9["mavp1"].ToString() + "'")[0]["mabyt"].ToString();
                            }
                            catch
                            {
                            }
                            if (str14 == "")
                            {
                                str14 = row9["mavp1"].ToString();
                            }
                            try
                            {
                                str13 = set2.Tables[0].Select("mavp='" + row9["mavp1"].ToString() + "'")[0]["manhombhyt"].ToString();
                            }
                            catch
                            {
                            }
                            str5 = str5 + this.f_nodexml_chitietcls(3, malk, ++num6, str14, str12, str13, row9["ten"].ToString(), row9["dvt"].ToString(), Convert.ToDecimal(row9["soluong"].ToString()), Math.Round(Convert.ToDecimal(row9["dongia"].ToString())), Convert.ToDecimal(row2["muc_huong"].ToString()), Math.Round((decimal) ((Convert.ToDecimal(row9["sotien"].ToString()) * Convert.ToDecimal(row2["muc_huong"].ToString())) / 100M)), "", row9["kihieubs"].ToString(), row2["ma_benh"].ToString(), "", "", "0");
                        }
                        str5 = str5 + this._s_arrCap[2] + "</BANG_CTDV>";
                        writer.WriteLine(str5);
                        str = this._s_arrCap[1] + "</THONGTINCHITIET>" + "\r\n</CHECKOUT>";
                        writer.WriteLine(str);
                        writer.Close();
                    }
                    catch
                    {
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, this._lib.Msg);
                }
            }
            dset.WriteXml("bang01_ngoai.xml", XmlWriteMode.WriteSchema);
            int num8 = 0;
            int num9 = 0;
            int num10 = 0;
            int num11 = 0;
            num8 = 5;
            num9 = dset.Tables[0].Rows.Count + 5;
            num10 = dset.Tables[0].Columns.Count - 1;
            num11 = num9;
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb01");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }

        public void f_Ngoaitru_xuatExcel_mau25ct(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add(new DataColumn("STT", typeof(decimal)));
            dset.Tables[0].Columns.Add(new DataColumn("mabn", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("hoten", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("namsinh", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("gioitinh", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("mathe", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ma_dkbd", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ten_bv", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("mabenh", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ngay_vao", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ngay_ra", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ngaydt", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_xn", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_cdha", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_thuoc", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_mau", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_pthuat", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_vtytth", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_vtyttt", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_dvktc", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_ktg", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_kham", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_vchuyen", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_tongchi", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_bhxh", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_benhnhan", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_ngoaids", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("lydo_vv", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ty_le", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("benhkhac", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ma_CSKCB", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("nam_qt", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("thang_qt", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("gt_tu", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("gt_den", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("diachi", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("giamdinh", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_xuattoan", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("lydo_xt", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_datuyen", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("t_vuottran", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("loaikcb", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("noi_ttoan", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("Sophieu", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ma_khoa", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ten_khoa", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("tuyen", typeof(string)));
            dset.Tables[0].Columns.Add(new DataColumn("ngaykham", typeof(string)));
            int num = 0;
            decimal num2 = this._lib.themoi15_sotien();
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    DataRow row2 = dset.Tables[0].NewRow();
                    row2["STT"] = ++num;
                    row2["mabn"] = row["mabn"].ToString();
                    row2["hoten"] = row["HOTEN"].ToString();
                    try
                    {
                        row2["mathe"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row2["mathe"] = row["sothe"].ToString();
                    }
                    row2["ma_dkbd"] = "'" + row["MANOIDK"].ToString();
                    row2["mabenh"] = row["MAICD"].ToString();
                    row2["benhkhac"] = row["chandoan"].ToString();
                    row2["sophieu"] = row["sophieu"].ToString();
                    row2["namsinh"] = row["ngaysinh"].ToString();
                    row2["gioitinh"] = row["phai"].ToString();
                    row2["ngay_vao"] = row["NGAYVAO"].ToString();
                    row2["ngay_ra"] = row["NGAYRA"].ToString();
                    row2["ngaydt"] = row["SONGAY"].ToString();
                    row2["t_xn"] = row["ST_1"].ToString();
                    row2["t_cdha"] = row["ST_2"].ToString();
                    row2["t_thuoc"] = row["ST_3"].ToString();
                    row2["t_mau"] = row["ST_4"].ToString();
                    row2["t_pthuat"] = row["ST_5"].ToString();
                    row2["t_vtytth"] = row["ST_6"].ToString();
                    try
                    {
                        row2["t_vtyttt"] = row["ST_14"].ToString();
                    }
                    catch
                    {
                    }
                    row2["t_dvktc"] = row["ST_7"].ToString();
                    row2["t_ktg"] = 0;
                    row2["t_kham"] = row["ST_9"].ToString();
                    row2["t_vchuyen"] = row["ST_10"].ToString();
                    row2["t_tongchi"] = row["tongcong"].ToString();
                    row2["t_bhxh"] = row["bhyttra"].ToString();
                    row2["ty_le"] = row["tylebhyt"].ToString();
                    row2["t_benhnhan"] = row["bntra"].ToString();
                    if ((decimal.Parse(row["traituyen"].ToString()) == 0M) && (decimal.Parse(row["tongcong"].ToString()) <= num2))
                    {
                        row2["t_bhxh"] = row["tongcong"].ToString();
                        row2["ty_le"] = 100;
                        row2["t_benhnhan"] = 0;
                    }
                    row2["t_ngoaids"] = 0;
                    row2["lydo_vv"] = row["lydo"].ToString();
                    if (row["lydo"].ToString() == "")
                    {
                        row2["lydo_vv"] = (row["traituyen"].ToString() == "1") ? 0 : 1;
                    }
                    row2["nam_qt"] = denngay.Substring(6, 4);
                    row2["thang_qt"] = denngay.Substring(3, 2);
                    row2["ma_CSKCB"] = this._lib.MABHXH;
                    row2["ma_khoa"] = row["makp"].ToString();
                    row2["diachi"] = row["diachi"].ToString();
                    row2["gt_tu"] = row["tungay"].ToString();
                    row2["gt_den"] = row["denngay"].ToString();
                    row2["ten_bv"] = row["noidk"].ToString();
                    row2["ten_khoa"] = row["tenkp"].ToString();
                    try
                    {
                        row2["ngaykham"] = row["ngayvao"].ToString();
                    }
                    catch
                    {
                    }
                    dset.Tables[0].Rows.Add(row2);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, this._lib.Msg);
                }
            }
            string[] strArray = new string[] { "ngaykham" };
            string str = "#" + this._tsxml.pNgoaiTru_MauExcel25CT_themdscot.Trim(new char[] { '#' }) + "#";
            for (int i = 0; i < strArray.Length; i++)
            {
                if (str.IndexOf("#" + strArray[i] + "#") < 0)
                {
                    dset.Tables[0].Columns.Remove(strArray[i]);
                }
            }
            int num4 = 0;
            int num5 = 0;
            int num6 = 0;
            int num7 = 0;
            int num8 = 0;
            num4 = 3;
            num5 = 5;
            num6 = dset.Tables[0].Rows.Count + 5;
            num7 = dset.Tables[0].Columns.Count - 1;
            num8 = num6;
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb25ct");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num4; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num4 + 8) + num5.ToString(), this._lib.getIndex(num7 - 0x12) + num6.ToString()).NumberFormat = "#,##";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(num7) + num8.ToString()).Borders.LineStyle = XlBorderWeight.xlHairline;
                for (int k = dset.Tables[0].Columns["ngay_ra"].Ordinal + 1; k < dset.Tables[0].Columns["lydo_vv"].Ordinal; k++)
                {
                    string[] strArray2 = new string[] { "=SUM(", this._lib.getIndex(k), "5:", this._lib.getIndex(k), (num6 - 1).ToString(), ")" };
                    this.osheet.Cells[num6, k + 1] = string.Concat(strArray2);
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num6.ToString(), this._lib.getIndex(num7 + 3) + num6.ToString());
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Bold = true;
                for (int m = 8; m < (num7 - 0x20); m++)
                {
                    this.orange = this.osheet.get_Range(this._lib.getIndex(m) + num5.ToString(), this._lib.getIndex(m) + num6.ToString());
                    this.orange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(num7 + 2) + num6.ToString());
                this.orange.Font.Name = "Arial";
                this.orange.Font.Size = 8;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 4] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M chữa BỆNH BHYT NGO?I TR\x00da ";
                this.osheet.Cells[2, 4] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("TỪ NG\x00c0Y " + tungay + " d?n ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(num7) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception2)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 s? li?u\n\n" + exception2.Message, this._lib.Msg);
            }
        }

        public void f_Ngoaitru_xuatExcel_mau25th(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dstam = new DataSet();
            dstam.Tables.Add("tam");
            dstam.Tables[0].Columns.Add("stt");
            dstam.Tables[0].Columns.Add("hoten");
            dstam.Tables[0].Columns.Add("soluot");
            dstam.Tables[0].Columns.Add("songay");
            dstam.Tables[0].Columns.Add("st_1");
            dstam.Tables[0].Columns.Add("st_2");
            dstam.Tables[0].Columns.Add("st_3");
            dstam.Tables[0].Columns.Add("st_4");
            dstam.Tables[0].Columns.Add("st_5");
            dstam.Tables[0].Columns.Add("st_6");
            dstam.Tables[0].Columns.Add("st_7");
            dstam.Tables[0].Columns.Add("st_8");
            dstam.Tables[0].Columns.Add("st_9");
            dstam.Tables[0].Columns.Add("st_10");
            dstam.Tables[0].Columns.Add("tongcong");
            dstam.Tables[0].Columns.Add("bntra");
            dstam.Tables[0].Columns.Add("bhyttra");
            dstam.Tables[0].Columns.Add("chiphids");
            dstam.Tables[0].Columns.Add("madk");
            DataSet set2 = new DataSet();
            set2 = dsdulieu.Copy();
            try
            {
                set2.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            set2.Tables[0].Columns.Add("yyyymmdd");
            set2.Tables[0].Columns.Add("soluot");
            for (int i = 0; i < set2.Tables[0].Rows.Count; i++)
            {
                try
                {
                    set2.Tables[0].Rows[i]["madk"] = int.Parse(set2.Tables[0].Rows[i]["MADK"].ToString());
                }
                catch
                {
                }
                try
                {
                    set2.Tables[0].Rows[i]["st_9"] = decimal.Parse(set2.Tables[0].Rows[i]["congkham"].ToString());
                }
                catch
                {
                }
                if (set2.Tables[0].Rows[i]["traituyen"].ToString() == "1")
                {
                    set2.Tables[0].Rows[i]["traituyen"] = 1;
                }
                else
                {
                    set2.Tables[0].Rows[i]["traituyen"] = 2;
                }
                try
                {
                    set2.Tables[0].Rows[i]["yyyymmdd"] = set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(6, 4) + set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(3, 2) + set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(0, 2);
                }
                catch
                {
                    set2.Tables[0].Rows[i]["yyyymmdd"] = 0;
                }
                set2.Tables[0].Rows[i]["soluot"] = 1;
            }
            DataSet set3 = new DataSet();
            set3 = set2.Copy();
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "st_1", "st_2", "st_3", "st_4", "st_5", "st_6", "st_7", "st_8", "st_9", "st_10", "TONGCONG", "bntra", "bhyttra", "chiphids", "soluot", "songay" };
            string str = "";
            DataRow row = dstam.Tables[0].NewRow();
            DataRow row2 = dstam.Tables[0].NewRow();
            for (int j = 1; j <= 3; j++)
            {
                switch (j)
                {
                    case 1:
                        str = "ABỆNH nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str = "BBỆNH nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str = "CBỆNH nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row3 = dstam.Tables[0].NewRow();
                row3["stt"] = str.Substring(0, 1);
                row3["hoten"] = str.Substring(1);
                dstam.Tables[0].Rows.Add(row3);
                DataRow row4 = dstam.Tables[0].NewRow();
                for (int k = 1; k <= 2; k++)
                {
                    DataRow row5 = dstam.Tables[0].NewRow();
                    row5["stt"] = (k == 1) ? "I" : "II";
                    row5["hoten"] = (k == 1) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN";
                    foreach (DataRow row6 in set3.Tables[0].Select("madk=" + j.ToString() + " and traituyen=" + k.ToString(), "yyyymmdd"))
                    {
                        for (int m = 0; m < strArray2.Length; m++)
                        {
                            try
                            {
                                if (row5[strArray2[m]].ToString() == "")
                                {
                                    row5[strArray2[m]] = 0;
                                }
                                if (row[strArray2[m]].ToString() == "")
                                {
                                    row[strArray2[m]] = 0;
                                }
                                if (row4[strArray2[m]].ToString() == "")
                                {
                                    row4[strArray2[m]] = 0;
                                }
                                if (row6[strArray2[m]].ToString() == "")
                                {
                                    row6[strArray2[m]] = 0;
                                }
                                row5[strArray2[m]] = decimal.Parse(row5[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                                row[strArray2[m]] = decimal.Parse(row[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                                row4[strArray2[m]] = decimal.Parse(row4[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    if ((row5["tongcong"].ToString() != "") && (row5["tongcong"].ToString() != "0"))
                    {
                        dstam.Tables[0].Rows.Add(row5);
                    }
                }
                row4["stt"] = str.Substring(0, 1);
                row4["hoten"] = "cộng " + row4["stt"].ToString();
                if ((row4["tongcong"].ToString() != "") && (row4["tongcong"].ToString() != "0"))
                {
                    dstam.Tables[0].Rows.Add(row4);
                }
                else
                {
                    dstam.Tables[0].Rows.RemoveAt(dstam.Tables[0].Rows.Count - 1);
                }
            }
            row["hoten"] = "TỔNG cộng A+B+C";
            dstam.Tables[0].Rows.Add(row);
            dstam.Tables[0].Columns.Remove("madk");
            dstam.Tables[0].Columns["st_1"].ColumnName = "X\x00e9t nghiệm";
            dstam.Tables[0].Columns["st_2"].ColumnName = "CĐHA TDCN";
            dstam.Tables[0].Columns["st_3"].ColumnName = "Thuốc dịch";
            dstam.Tables[0].Columns["st_4"].ColumnName = "M\x00e1u";
            dstam.Tables[0].Columns["st_5"].ColumnName = "Thủ thuật phẫu thuật";
            dstam.Tables[0].Columns["st_6"].ColumnName = "V?t tu y t?";
            dstam.Tables[0].Columns["st_7"].ColumnName = "DVKT cao";
            dstam.Tables[0].Columns["st_8"].ColumnName = "Thu?c k, CTG";
            dstam.Tables[0].Columns["st_9"].ColumnName = "Tiền kh\x00e1m";
            dstam.Tables[0].Columns["st_10"].ColumnName = "CP Vận chuyển";
            dstam.Tables[0].AcceptChanges();
            this.f_Ngoaitru_xuatExcel_mau25th_run(print, dstam, tungay, denngay);
        }

        private void f_Ngoaitru_xuatExcel_mau25th_run(bool print, DataSet dstam, string tungay, string denngay)
        {
            this._lib.check_process_Excel();
            try
            {
                DataRow row = dstam.Tables[0].NewRow();
                for (int i = 0; i < dstam.Tables[0].Columns.Count; i++)
                {
                    row[i] = i + 1;
                }
                dstam.Tables[0].Rows.InsertAt(row, 0);
                int num2 = 6;
                int num3 = 5;
                int count = dstam.Tables[0].Columns.Count;
                this.tenfile = this._lib.Export_Excel(dstam, "bccpkcb_25th");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num2; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).Font.Name = "Arial";
                this.osheet.get_Range(this._lib.getIndex(2) + num3, this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).NumberFormat = "#,##0";
                this.osheet.Cells[1, 1] = this._lib.Tenbv;
                this.osheet.Cells[1, count] = "Mẫu số: 25A-TH/BHYT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "1", this._lib.getIndex((count - 1) - 2) + "1");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Size = 8;
                this.orange.MergeCells = true;
                this.osheet.Cells[2, 1] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M chữa BỆNH BHYT NGO?I TR\x00da ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "2", this._lib.getIndex(count - 1) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 0x10;
                this.osheet.Cells[3, 1] = "TỪ NG\x00c0Y " + tungay + " d?n ng\x00e0y " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "3", this._lib.getIndex(count - 1) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num2 + 2));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                int num6 = -1;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "TT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.ColumnWidth = 5;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Ch? ti\x00eau";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "S? lu?t";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Số ng\x00e0y di?u tr?";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num3 + 1, num6 + 1] = "CHI PH\x00cd PH\x00c1T SINH T?I CO S? KH\x00c1M chữa BỆNH".ToUpper();
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num3 + 1), this._lib.getIndex((count - 1) - 3) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 2] = "Người bệnh c\x00f9ng chi trả";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 2) + (num2 + 1), this._lib.getIndex((count - 1) - 2) + (num3 + 1));
                this.orange.ColumnWidth = 10;
                this.orange.MergeCells = true;
                this.osheet.Cells[num3 + 1, count - 1] = "Chi ph\x00ed đề nghị BHXH thanh to\x00e1n";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num3 + 1), this._lib.getIndex(count - 1) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 1] = "Số tiền";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num2 + 1), this._lib.getIndex((count - 1) - 1) + (num2 + 1));
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, count] = "Trong đ\x00f3 chi ph\x00ed ngo\x00e0i quỹ định suất";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, count - 3] = "TỔNG cộng";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 4) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                int num7 = num2 + 2;
                for (int k = 1; k < dstam.Tables[0].Rows.Count; k++)
                {
                    num7++;
                    this.orange = this.osheet.get_Range("A" + num7.ToString(), this._lib.getIndex(count - 1) + num7.ToString());
                    if (((dstam.Tables[0].Rows[k]["stt"].ToString() == "A") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((dstam.Tables[0].Rows[k]["stt"].ToString() == "C") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                        if ((dstam.Tables[0].Rows[k][1].ToString().Substring(0, 1) != "C") && (dstam.Tables[0].Rows[k][1].ToString().Substring(0, 1) != "T"))
                        {
                            this.orange = this.osheet.get_Range("B" + num7.ToString(), this._lib.getIndex(count - 1) + num7.ToString());
                            this.orange.MergeCells = true;
                        }
                    }
                }
                this.oxl.ActiveWindow.DisplayZeros = false;
                try
                {
                    this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                    this.osheet.PageSetup.LeftMargin = 20.0;
                    this.osheet.PageSetup.RightMargin = 20.0;
                    this.osheet.PageSetup.TopMargin = 30.0;
                    this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                }
                catch
                {
                }
                decimal num9 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["tongcong"].ToString()), 0);
                string str = new numbertotext().doiraso(num9.ToString());
                this.osheet.Cells[((num2 + 1) + dstam.Tables[0].Rows.Count) + 1, 2] = "Số tiền đề nghị thanh to\x00e1n (viết bằng chữ): " + str.Substring(0, 1).ToUpper() + str.Substring(1) + " đồng chẵn.";
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public void f_Ngoaitru_xuatExcel_mau79(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            string str = "";
            DataSet dstam = new DataSet();
            dstam.Tables.Add("tam");
            dstam.Tables[0].Columns.Add("stt");
            dstam.Tables[0].Columns.Add("hoten");
            dstam.Tables[0].Columns.Add("ngaysinh_nam");
            dstam.Tables[0].Columns.Add("ngaysinh_nu");
            dstam.Tables[0].Columns.Add("sothe");
            dstam.Tables[0].Columns.Add("manoidk");
            dstam.Tables[0].Columns.Add("maicd");
            dstam.Tables[0].Columns.Add("ngayvao");
            dstam.Tables[0].Columns.Add("tongcong");
            dstam.Tables[0].Columns.Add("st_1");
            dstam.Tables[0].Columns.Add("st_2");
            dstam.Tables[0].Columns.Add("st_3");
            dstam.Tables[0].Columns.Add("st_4");
            dstam.Tables[0].Columns.Add("st_5");
            dstam.Tables[0].Columns.Add("st_6");
            dstam.Tables[0].Columns.Add("st_14");
            dstam.Tables[0].Columns.Add("st_7");
            dstam.Tables[0].Columns.Add("st_8");
            dstam.Tables[0].Columns.Add("st_9");
            dstam.Tables[0].Columns.Add("st_10");
            dstam.Tables[0].Columns.Add("bntra");
            dstam.Tables[0].Columns.Add("bhyttra");
            dstam.Tables[0].Columns.Add("chiphids");
            dstam.Tables[0].Columns.Add("madk");
            dstam.Tables[0].Columns.Add("ngaytoikham");
            str = str + "ngaytoikham;";
            dstam.Tables[0].Columns.Add("nguoiduyet");
            str = str + "nguoiduyet;";
            dstam.Tables[0].Columns.Add("mabn");
            str = str + "mabn;";
            dstam.Tables[0].Columns.Add("diachi");
            str = str + "diachi;";
            dstam.Tables[0].Columns.Add("denngay");
            str = str + "denngay;";
            dstam.Tables[0].Columns.Add("tenbenh");
            str = str + "tenbenh;";
            dstam.Tables[0].Columns.Add("sophieu");
            str = str + "sophieu;";
            dstam.Tables[0].Columns.Add("mabv");
            str = str + "mabv;";
            dstam.Tables[0].Columns.Add("tenbvdk");
            str = str + "tenbvdk;";
            dstam.Tables[0].Columns.Add("tungay");
            str = str + "tungay;";
            dstam.Tables[0].Columns.Add("lydo_vv");
            str = str + "lydo_vv;";
            dstam.Tables[0].Columns.Add("tenkp");
            str = str + "tenkp;";
            DataSet set2 = new DataSet();
            set2 = dsdulieu.Copy();
            set2.Tables[0].Columns.Add("ngaysinh_nam");
            set2.Tables[0].Columns.Add("ngaysinh_nu");
            try
            {
                set2.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            set2.Tables[0].Columns.Add("yyyymmdd");
            for (int i = 0; i < set2.Tables[0].Rows.Count; i++)
            {
                try
                {
                    try
                    {
                        set2.Tables[0].Rows[i]["manoidk"] = ((set2.Tables[0].Rows[i]["manoidk"].ToString().Substring(0, 1) == "0") ? "'" : "") + set2.Tables[0].Rows[i]["manoidk"].ToString();
                    }
                    catch
                    {
                        set2.Tables[0].Rows[i]["manoidk"] = "";
                    }
                    if (set2.Tables[0].Rows[i]["phai"].ToString() == "0")
                    {
                        set2.Tables[0].Rows[i]["phai"] = "Nam";
                    }
                    else
                    {
                        set2.Tables[0].Rows[i]["phai"] = "Nữ";
                    }
                    if (set2.Tables[0].Rows[i]["phai"].ToString().ToUpper() == "NAM")
                    {
                        set2.Tables[0].Rows[i]["ngaysinh_nam"] = set2.Tables[0].Rows[i]["ngaysinh"].ToString();
                    }
                    else
                    {
                        set2.Tables[0].Rows[i]["ngaysinh_nu"] = set2.Tables[0].Rows[i]["ngaysinh"].ToString();
                    }
                    try
                    {
                        set2.Tables[0].Rows[i]["madk"] = int.Parse(set2.Tables[0].Rows[i]["MADK"].ToString());
                    }
                    catch
                    {
                    }
                    try
                    {
                        if (set2.Tables[0].Rows[i]["st_9"].ToString() == "0")
                        {
                            set2.Tables[0].Rows[i]["st_9"] = decimal.Parse(set2.Tables[0].Rows[i]["congkham"].ToString());
                        }
                    }
                    catch
                    {
                    }
                    if (set2.Tables[0].Rows[i]["traituyen"].ToString() == "0")
                    {
                        set2.Tables[0].Rows[i]["traituyen"] = 1;
                    }
                    else
                    {
                        set2.Tables[0].Rows[i]["traituyen"] = 2;
                    }
                    try
                    {
                        set2.Tables[0].Rows[i]["yyyymmdd"] = set2.Tables[0].Rows[i]["ngayvao"].ToString().Substring(6, 4) + set2.Tables[0].Rows[i]["ngayvao"].ToString().Substring(3, 2) + set2.Tables[0].Rows[i]["ngayvao"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        set2.Tables[0].Rows[i]["yyyymmdd"] = 0;
                    }
                    if (set2.Tables[0].Rows[i]["sothe"].ToString().Length > 15)
                    {
                        set2.Tables[0].Rows[i]["sothe"] = set2.Tables[0].Rows[i]["sothe"].ToString().Substring(0, 15);
                    }
                    try
                    {
                        set2.Tables[0].Rows[i]["denngay"] = set2.Tables[0].Rows[i]["gtriden"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        set2.Tables[0].Rows[i]["tungay"] = set2.Tables[0].Rows[i]["gtritu"].ToString();
                    }
                    catch
                    {
                    }
                }
                catch (Exception exception)
                {
                    this._lib.f_write_log(exception.ToString());
                }
            }
            DataSet set3 = set2.Copy();
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "st_1", "st_2", "st_3", "st_4", "st_5", "st_6", "st_7", "st_8", "st_9", "st_10", "st_14", "TONGCONG", "bntra", "bhyttra", "chiphids" };
            string str2 = "";
            DataRow row = dstam.Tables[0].NewRow();
            DataRow row2 = dstam.Tables[0].NewRow();
            int num2 = 0;
            for (int j = 1; j <= 3; j++)
            {
                switch (j)
                {
                    case 1:
                        str2 = "ABệnh nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str2 = "BBệnh nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str2 = "CBệnh nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row3 = dstam.Tables[0].NewRow();
                row3["stt"] = str2.Substring(0, 1);
                row3["hoten"] = str2.Substring(1);
                dstam.Tables[0].Rows.Add(row3);
                DataRow row4 = dstam.Tables[0].NewRow();
                for (int k = 1; k <= 2; k++)
                {
                    DataRow row5 = dstam.Tables[0].NewRow();
                    row5["stt"] = row3["stt"].ToString() + ((k == 1) ? "I" : "II");
                    row5["hoten"] = (k == 1) ? "BỆNH NH\x00c2N KCB Đ\x00daNG TUYẾN (C\x00d3 GIẤY CHUYỂN VIỆN HOẶC TH CẤP CỨU)" : "BỆNH NH\x00c2N KCB TR\x00c1I TUYẾN (KH\x00d4NG C\x00d3 GIẤY CHUYỂN VIỆN HOẶC KH\x00d4NG PHẢI TRƯỜNG HỢP CẤP CỨU)";
                    dstam.Tables[0].Rows.Add(row5);
                    DataRow row6 = dstam.Tables[0].NewRow();
                    foreach (DataRow row7 in set3.Tables[0].Select("madk=" + j.ToString() + " and traituyen=" + k.ToString(), "yyyymmdd"))
                    {
                        DataRow row8 = dstam.Tables[0].NewRow();
                        for (int m = 0; m < dstam.Tables[0].Columns.Count; m++)
                        {
                            if (dstam.Tables[0].Columns[m].ColumnName == "ngaytoikham")
                            {
                                try
                                {
                                    row8[m] = row7["ngayvao"].ToString();
                                }
                                catch
                                {
                                }
                                try
                                {
                                    row8["ngayvao"] = row7["ngay"].ToString();
                                }
                                catch
                                {
                                }
                            }
                            else
                            {
                                try
                                {
                                    row8[m] = row7[dstam.Tables[0].Columns[m].ColumnName].ToString();
                                }
                                catch
                                {
                                }
                            }
                        }
                        try
                        {
                            row8["nguoiduyet"] = row7["nguoiduyet"].ToString();
                        }
                        catch
                        {
                            row8["nguoiduyet"] = "";
                        }
                        try
                        {
                            row8["diachi"] = row7["diachi"].ToString();
                        }
                        catch
                        {
                            row8["diachi"] = "";
                        }
                        try
                        {
                            row8["denngay"] = row7["denngay"].ToString();
                        }
                        catch
                        {
                            row8["denngay"] = "";
                        }
                        try
                        {
                            row8["tungay"] = row7["tungay"].ToString();
                        }
                        catch
                        {
                            row8["tungay"] = "";
                        }
                        try
                        {
                            row8["lydo_vv"] = row7["lydo"].ToString();
                        }
                        catch
                        {
                            row8["lydo_vv"] = "";
                        }
                        try
                        {
                            row8["tenbenh"] = row7["chandoan"].ToString();
                        }
                        catch
                        {
                            row8["tenbenh"] = "";
                        }
                        try
                        {
                            row8["sophieu"] = row7["sophieu"].ToString();
                        }
                        catch
                        {
                            row8["sophieu"] = "";
                        }
                        try
                        {
                            row8["mabv"] = row7["manoidk"].ToString();
                        }
                        catch
                        {
                            row8["mabv"] = "";
                        }
                        try
                        {
                            row8["tenbvdk"] = row7["noidk"].ToString();
                        }
                        catch
                        {
                            row8["tenbvdk"] = "";
                        }
                        row8["stt"] = ++num2;
                        dstam.Tables[0].Rows.Add(row8);
                        for (int n = 0; n < strArray2.Length; n++)
                        {
                            try
                            {
                                if (row6[strArray2[n]].ToString() == "")
                                {
                                    row6[strArray2[n]] = 0;
                                }
                                if (row[strArray2[n]].ToString() == "")
                                {
                                    row[strArray2[n]] = 0;
                                }
                                if (row4[strArray2[n]].ToString() == "")
                                {
                                    row4[strArray2[n]] = 0;
                                }
                                if (row7[strArray2[n]].ToString() == "")
                                {
                                    row7[strArray2[n]] = 0;
                                }
                                row6[strArray2[n]] = decimal.Parse(row6[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                                row[strArray2[n]] = decimal.Parse(row[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                                row4[strArray2[n]] = decimal.Parse(row4[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    row6["stt"] = (k == 1) ? "I" : "II";
                    row6["hoten"] = "TỔNG " + ((k == 1) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN");
                    if ((row6["tongcong"].ToString() != "") && (row6["tongcong"].ToString() != "0"))
                    {
                        dstam.Tables[0].Rows.Add(row6);
                    }
                    else
                    {
                        dstam.Tables[0].Rows.Remove(row5);
                    }
                }
                row4["stt"] = str2.Substring(0, 1);
                row4["hoten"] = "cộng " + row4["stt"].ToString();
                if ((row4["tongcong"].ToString() != "") && (row4["tongcong"].ToString() != "0"))
                {
                    dstam.Tables[0].Rows.Add(row4);
                }
                else
                {
                    dstam.Tables[0].Rows.RemoveAt(dstam.Tables[0].Rows.Count - 1);
                }
            }
            row["hoten"] = "TỔNG CỘNG A+B+C";
            dstam.Tables[0].Rows.Add(row);
            dstam.Tables[0].Columns.Remove("madk");
            dstam.WriteXml("dsxmlexcel79.xml", XmlWriteMode.WriteSchema);
            foreach (string str3 in str.Trim(new char[] { ';' }).Split(new char[] { ';' }))
            {
                if (this._tsxml.pNgoaiTru_MauExcel79HD_themdscot.IndexOf(str3 + ";") < 0)
                {
                    dstam.Tables[0].Columns.Remove(str3);
                }
            }
            dstam.Tables[0].Columns["ngaysinh_nam"].ColumnName = "Nam";
            dstam.Tables[0].Columns["ngaysinh_nu"].ColumnName = "NỮ";
            dstam.Tables[0].Columns["ngayvao"].ColumnName = "ngay";
            dstam.Tables[0].Columns["st_1"].ColumnName = "X\x00e9t nghiệm";
            dstam.Tables[0].Columns["st_2"].ColumnName = "CĐHA TDCN";
            dstam.Tables[0].Columns["st_3"].ColumnName = "Thuốc dịch";
            dstam.Tables[0].Columns["st_4"].ColumnName = "M\x00e1u";
            dstam.Tables[0].Columns["st_5"].ColumnName = "Thủ thuật phẫu thuật";
            dstam.Tables[0].Columns["st_6"].ColumnName = "Vật tư y tế ti\x00eau hao";
            dstam.Tables[0].Columns["st_7"].ColumnName = "DVKT cao";
            dstam.Tables[0].Columns["st_8"].ColumnName = "Thuốc k, thải gh\x00e9p";
            dstam.Tables[0].Columns["st_9"].ColumnName = "Tiền kh\x00e1m";
            dstam.Tables[0].Columns["st_10"].ColumnName = "Vận chuyển";
            dstam.Tables[0].Columns["st_14"].ColumnName = "Vật tư thay thế";
            dstam.Tables[0].AcceptChanges();
            this.f_Ngoaitru_xuatExcel_mau79_run(print, dstam, tungay, denngay);
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
            string mabv = this._lib.Mabv;
            decimal d = 0M;
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    DataRow row2 = dset.Tables[0].NewRow();
                    row2["STT"] = d = decimal.op_Increment(d);
                    row2["ma_bn"] = row["mabn"].ToString();
                    row2["ho_ten"] = this._lib.f_hoten(row["HOTEN"].ToString());
                    row2["ngay_sinh"] = row["ngaysinh"].ToString();
                    row2["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row2["dia_chi"] = row["diachi"].ToString();
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
                    row2["tenkp"] = row["tenkp"].ToString();
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
                    row2["ma_cskcb"] = this._lib.Mabv;
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
                catch (Exception exception)
                {
                    this._lib.f_write_log(exception.ToString());
                }
            }
            dset.Tables[0].Columns.Remove("maql");
            for (int i = 0; i < dset.Tables[0].Columns.Count; i++)
            {
                dset.Tables[0].Columns[i].ColumnName = dset.Tables[0].Columns[i].ColumnName.ToUpper();
            }
            dset.AcceptChanges();
            dset.WriteXml("bang7980.xml", XmlWriteMode.WriteSchema);
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb02");
            try
            {
                Process.Start(this.tenfile);
            }
            catch
            {
            }
        }

        private void f_Ngoaitru_xuatExcel_mau79_run(bool print, DataSet dstam, string tungay, string denngay)
        {
            this._lib.check_process_Excel();
            try
            {
                DataRow row = dstam.Tables[0].NewRow();
                for (int i = 0; i < dstam.Tables[0].Columns.Count; i++)
                {
                    if (i <= dstam.Tables[0].Columns["ngay"].Ordinal)
                    {
                        row[i] = Convert.ToChar((int) (i + 0x41));
                    }
                    else
                    {
                        row[i] = i - dstam.Tables[0].Columns["ngay"].Ordinal;
                    }
                }
                dstam.Tables[0].Rows.InsertAt(row, 0);
                int num2 = 9;
                int num3 = 7;
                int ordinal = dstam.Tables[0].Columns["chiphids"].Ordinal;
                int count = dstam.Tables[0].Columns.Count;
                this.tenfile = this._lib.Export_Excel(dstam, "bccpkcb");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num2; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).Font.Name = "Arial";
                this.osheet.get_Range(this._lib.getIndex(8) + num3, this._lib.getIndex(ordinal - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).NumberFormat = "#,##0";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.osheet.Cells[1, count] = "Mẫu số: C79a-HD\n(Ban h\x00e0nh theo Th\x00f4ng tư số 178/TT \nng\x00e0y 23/10/2012 của Bộ T\x00e0i Ch\x00ednh)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "1", this._lib.getIndex((count - 1) - 4) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Size = 8;
                this.orange.MergeCells = true;
                this.osheet.Cells[3, 3] = "DANH S\x00c1CH NGƯỜI BỆNH BẢO HIỂM Y TẾ KH\x00c1M CHỮA BỆNH NGOẠI TR\x00da ĐỀ NGHỊ THANH TO\x00c1N";
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "3", this._lib.getIndex((count - 1) - 2) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 0x10;
                this.osheet.Cells[5, 3] = "TỪ NG\x00c0Y " + tungay + " ĐẾN " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "5", this._lib.getIndex((count - 1) - 2) + "5");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.osheet.Cells[6, count] = "ĐVT: Đồng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "6", this._lib.getIndex((count - 1) - 2) + "6");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.MergeCells = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 2), this._lib.getIndex(count - 1) + (num2 + 2));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num3 + 1));
                this.orange.RowHeight = 0x2d;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num2, this._lib.getIndex(count - 1) + num2);
                this.orange.RowHeight = 30;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.RowHeight = 0x37;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count - 1) + 1);
                this.orange.RowHeight = 20;
                int num7 = -1;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "STT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.ColumnWidth = 5;
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "Họ v\x00e0 t\x00ean";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num3 + 1, num7 + 1] = "Năm sinh";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num3 + 1), this._lib.getIndex(num7 + 1) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, num7 + 1] = "Nam";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + num2);
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "Nữ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + num2);
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "M\x00e3 thẻ BHYT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "M\x00e3 ĐK BĐ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.MergeCells = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + ((num2 + 1) + 1), this._lib.getIndex(num7) + (((num2 + 1) + 1) + dstam.Tables[0].Rows.Count));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "M\x00e3 bệnh";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2 + 1, num7 + 1] = "Ng\x00e0y kh\x00e1m";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + (num3 + 1));
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num3 + 1, num7 + 1] = "Tổng chi ph\x00ed kh\x00e1m chữa bệnh bhyt".ToUpper();
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num3 + 1), this._lib.getIndex(ordinal - 3) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, num7 + 1] = "Tổng cộng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + (num2 + 1), this._lib.getIndex(num7) + num2);
                this.orange.MergeCells = true;
                num7++;
                this.osheet.Cells[num2, num7 + 1] = "Trong đ\x00f3";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num7) + num2, this._lib.getIndex(ordinal - 3) + num2);
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, ordinal - 1] = "Người bệnh c\x00f9ng chi trả";
                this.orange = this.osheet.get_Range(this._lib.getIndex((ordinal - 1) - 1) + (num2 + 1), this._lib.getIndex((ordinal - 1) - 1) + (num3 + 1));
                this.orange.ColumnWidth = 10;
                this.orange.MergeCells = true;
                this.osheet.Cells[num3 + 1, ordinal] = "Chi ph\x00ed đề nghị cơ quan BHYT thanh to\x00e1n";
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal - 1) + (num3 + 1), this._lib.getIndex(ordinal) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, ordinal] = "Số tiền";
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal - 1) + (num2 + 1), this._lib.getIndex(ordinal - 1) + num2);
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, ordinal + 1] = "Trong đ\x00f3 chi ph\x00ed ngo\x00e0i quỹ định suất";
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + (num2 + 1), this._lib.getIndex(ordinal) + num2);
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                string[] strArray = this._tsxml.pNgoaiTru_MauExcel79HD_themdscot.Split(new char[] { '#' });
                for (int k = 0; k < strArray.Length; k++)
                {
                    try
                    {
                        int num9 = dstam.Tables[0].Columns[strArray[k].Split(new char[] { ';' })[0]].Ordinal;
                        this.osheet.Cells[num2 + 1, num9 + 1] = strArray[k].Split(new char[] { ';' })[1];
                        this.orange = this.osheet.get_Range(this._lib.getIndex(num9) + (num2 + 1), this._lib.getIndex(num9) + (num2 - 1));
                        this.orange.MergeCells = true;
                    }
                    catch
                    {
                    }
                }
                int num10 = num2 + 2;
                string[] strArray2 = new string[] { "manoidk", "mabv", "ngaytoikham", "mabn" };
                for (int m = 0; m < strArray2.Length; m++)
                {
                    try
                    {
                        int num12 = dstam.Tables[0].Columns[strArray2[m]].Ordinal;
                        this.orange = this.osheet.get_Range(this._lib.getIndex(num12) + (num10 + 1), this._lib.getIndex(num12) + ((num10 + dstam.Tables[0].Rows.Count) - 1));
                        this.orange.NumberFormat = "@";
                    }
                    catch
                    {
                    }
                }
                for (int n = 1; n < dstam.Tables[0].Rows.Count; n++)
                {
                    num10++;
                    this.orange = this.osheet.get_Range("A" + num10.ToString(), this._lib.getIndex(count - 1) + num10.ToString());
                    if (((dstam.Tables[0].Rows[n]["stt"].ToString() == "A") || (dstam.Tables[0].Rows[n]["stt"].ToString() == "B")) || ((dstam.Tables[0].Rows[n]["stt"].ToString() == "C") || (dstam.Tables[0].Rows[n]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                    }
                    else if (((dstam.Tables[0].Rows[n]["hoten"].ToString().IndexOf("Đ\x00daNG TUYẾN") > -1) || (dstam.Tables[0].Rows[n]["hoten"].ToString().IndexOf("TR\x00c1I TUYẾN") > -1)) || ((dstam.Tables[0].Rows[n]["hoten"].ToString().IndexOf("Đ\x00daNG TUYẾN") > -1) || (dstam.Tables[0].Rows[n]["hoten"].ToString().IndexOf("TR\x00c1I TUYẾN") > -1)))
                    {
                        this.orange.Font.ColorIndex = 10;
                        this.orange.Font.Bold = true;
                    }
                    if (dstam.Tables[0].Rows[n]["tongcong"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range("B" + num10, this._lib.getIndex(count - 1) + num10);
                        this.orange.MergeCells = true;
                    }
                    else if ((dstam.Tables[0].Rows[n]["stt"].ToString() == "") || !char.IsDigit(dstam.Tables[0].Rows[n]["stt"].ToString(), 0))
                    {
                        this.orange = this.osheet.get_Range("B" + num10, this._lib.getIndex(dstam.Tables[0].Columns["ngay"].Ordinal) + num10);
                        this.orange.MergeCells = true;
                    }
                    else
                    {
                        for (int num14 = 0; num14 < strArray2.Length; num14++)
                        {
                            try
                            {
                                int num15 = dstam.Tables[0].Columns[strArray2[num14]].Ordinal;
                                this.osheet.Cells[num10, num15 + 1] = dstam.Tables[0].Rows[n][num15].ToString();
                            }
                            catch
                            {
                            }
                        }
                    }
                }
                this.orange.MergeCells = true;
                this.oxl.ActiveWindow.DisplayZeros = false;
                try
                {
                    this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                    this.osheet.PageSetup.LeftMargin = 20.0;
                    this.osheet.PageSetup.RightMargin = 20.0;
                    this.osheet.PageSetup.TopMargin = 30.0;
                    this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                }
                catch
                {
                }
                decimal num16 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["bhyttra"].ToString()), 0);
                string str = new numbertotext().doiraso(num16.ToString());
                this.osheet.Cells[((num2 + 1) + dstam.Tables[0].Rows.Count) + 1, 2] = "Số tiền đề nghị thanh to\x00e1n (viết bằng chữ): " + str.Substring(0, 1).ToUpper() + str.Substring(1) + " đồng chẵn.";
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 2] = "Người lập";
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 2] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 6] = "Trưởng ph\x00f2ng KHTH";
                this.orange = this.osheet.get_Range(this._lib.getIndex(5) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(7) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 6] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(5) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(7) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 15] = "Kế to\x00e1n trưởng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(14) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(15) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 15] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(14) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(15) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 3, count] = "Ng\x00e0y " + DateTime.Now.ToString("dd") + " th\x00e1ng " + DateTime.Now.ToString("MM") + " nam " + DateTime.Now.ToString("yyyy");
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 3), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 3));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, count] = "Thủ trưởng đơn vị";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, count] = "(K\x00fd, họ t\x00ean, đ\x00f3ng dấu)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void f_Ngoaitru_xuatExcel_mau79_run2(bool print, DataSet dstam, string tungay, string denngay)
        {
            string path = Environment.CurrentDirectory + "//Excel//excelmau79hd.xls";
            string str2 = "Arial";
            StringBuilder builder = new StringBuilder();
            StreamWriter writer = new StreamWriter(path, false, Encoding.Unicode);
            builder.Append("<table>");
            builder.Append("<tr>");
            builder.Append("<td colspan=3 style=\"font-family:" + str2 + ";align=left\">" + this._lib.Syte + "</td>");
            for (int i = 3; i < (dstam.Tables[0].Columns.Count - 1); i++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<td colspan=1 align=right style=\"font-family:" + str2 + ";font-size:8pt\">Mẫu số: C79a-HD(Ban h\x00e0nh theo Th\x00f4ng tư số 178/TT ng\x00e0y 23/10/2012 của Bộ T\x00e0i Ch\x00ednh)</td>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td colspan=3 style=\"font-family:" + str2 + ";align=left\">" + this._lib.Tenbv + "</td>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<th colspan=", dstam.Tables[0].Columns.Count, " align=centre style=\"font-family: ", str2, "; font-size: 16pt\">DANH S\x00c1CH NGƯỜI BỆNH BẢO HIỂM Y TẾ KH\x00c1M CHỮA BỆNH NGOẠI TR\x00da ĐỀ NGHỊ THANH TO\x00c1N</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<th colspan=", dstam.Tables[0].Columns.Count, " align=centre style=\"font-family: ", str2, "; font-size: 12pt\">", (tungay == denngay) ? ("Ng\x00e0y " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay), "</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count, " align=right style=\"font-family: ", str2, "; font-size: 10pt\">ĐVT: Đồng</td>" }));
            builder.Append("</tr>");
            builder.Append("<tr></tr>");
            builder.Append("</table>");
            builder.Append("<table border=1>");
            builder.Append("<tr>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">STT</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Họ v\x00e0 t\x00ean</th>");
            builder.Append("<th rowspan=1 colspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Năm sinh</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 thẻ BHYT</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 ĐK BĐ</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 bệnh</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Ng\x00e0y kh\x00e1m</th>");
            builder.Append("<th rowspan=1 colspan=12 align=centre style=\"font-family: " + str2 + "\">TỔNG CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT</th>");
            builder.Append("<th rowspan=3 align=centre style=\"font-family: " + str2 + ";height:30px\">Người bệnh \nc\x00f9ng chi trả</th>");
            builder.Append("<th rowspan=1 colspan=2 align=centre style=\"font-family: " + str2 + ";height:30pt\">Chi ph\x00ed đề nghị\n cơ quan BHYT \nthanh to\x00e1n</th>");
            if (this._tsxml.pNgoaiTru_MauExcel79HD_themdscot != "")
            {
                string[] strArray = this._tsxml.pNgoaiTru_MauExcel79HD_themdscot.Split(new char[] { '#' });
                for (int num2 = 0; num2 < strArray.Length; num2++)
                {
                    builder.Append("<th rowspan=3 align=centre style=\"font-family: " + str2 + "\">" + strArray[num2].Split(new char[] { ';' })[1] + "</th>");
                }
            }
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Nam</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Nữ</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Tổng cộng</th>");
            builder.Append("<th rowspan=1 colspan=11 height=30 align=centre style=\"font-family: " + str2 + "\">Trong đ\x00f3</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Số tiền</th>");
            builder.Append("<th rowspan=2 align=centre style=\"font-family: " + str2 + ";height:30px; width:10px\">Trong đ\x00f3 \n\rchi ph\x00ed ngo\x00e0i quỹ \n\rđịnh suất</th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">X\x00e9t nghiệm</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">CĐHA TDCN</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thuốc dịch</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">M\x00e1u</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thủ thuật phẫu thuật</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vật tư y tế ti\x00eau hao</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vật tư thay thế</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">DVKT cao</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thuốc k, thải gh\x00e9p</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Tiền kh\x00e1m</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vận chuyển</th>");
            builder.Append("</tr>");
            builder.Append("<tr >");
            for (int j = 0; j < dstam.Tables[0].Columns.Count; j++)
            {
                if (j <= dstam.Tables[0].Columns["ngay"].Ordinal)
                {
                    builder.Append(string.Concat(new object[] { "<th style=\"font-family: ", str2, ";font-size: 11pt\">", Convert.ToChar((int) (j + 0x41)), "</th>" }));
                }
                else
                {
                    builder.Append(string.Concat(new object[] { "<th style=\"font-family: ", str2, ";font-size: 11pt\">", j - dstam.Tables[0].Columns["ngay"].Ordinal, "</th>" }));
                }
            }
            builder.Append("</tr>\n");
            builder.Append("</table>");
            writer.Write(builder);
            builder = new StringBuilder();
            builder.Append("<table border=1 style=\"font-family:" + str2 + "\">");
            decimal num4 = 0M;
            decimal num5 = 0M;
            for (int k = 0; k < dstam.Tables[0].Rows.Count; k++)
            {
                num4 = 0M;
                num5 = 0M;
                try
                {
                    num4 = decimal.Parse(dstam.Tables[0].Rows[k]["tongcong"].ToString());
                }
                catch
                {
                }
                try
                {
                    num5 = decimal.Parse(dstam.Tables[0].Rows[k]["stt"].ToString());
                }
                catch
                {
                }
                builder.Append("<tr>");
                for (int num7 = 0; num7 < dstam.Tables[0].Columns.Count; num7++)
                {
                    if ((num5 == 0M) && (num4 == 0M))
                    {
                        string str3 = "style=\"font-family:" + str2 + ";font-size:10pt; font-weight: bold";
                        if (dstam.Tables[0].Rows[k][0].ToString().Length == 1)
                        {
                            str3 = str3 + ";color:blue";
                        }
                        else
                        {
                            str3 = str3 + ";color:DarkGreen";
                        }
                        str3 = str3 + "\"";
                        if (num7 == 0)
                        {
                            builder.Append("<td " + str3 + ">" + dstam.Tables[0].Rows[k][num7].ToString() + "</td>");
                            continue;
                        }
                        builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count - 1, " ", str3, ">", dstam.Tables[0].Rows[k][num7].ToString(), "</td>" }));
                        break;
                    }
                    if (num5 == 0M)
                    {
                        string str4 = "style=\"font-family:" + str2 + ";font-size:10pt; font-weight: bold";
                        if ((dstam.Tables[0].Rows[k][0].ToString().Length < 1) || ((dstam.Tables[0].Rows[k][0].ToString().Length >= 1) && ("A,B,C".IndexOf(dstam.Tables[0].Rows[k][0].ToString().Substring(0, 1)) > -1)))
                        {
                            str4 = str4 + ";color:blue";
                        }
                        else
                        {
                            str4 = str4 + ";color:DarkGreen";
                        }
                        str4 = str4 + "\"";
                        switch (num7)
                        {
                            case 0:
                            {
                                builder.Append("<td " + str4 + ">" + dstam.Tables[0].Rows[k][num7].ToString() + "</td>");
                                continue;
                            }
                            case 1:
                            {
                                builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns["tongcong"].Ordinal - 1, " ", str4, ">", dstam.Tables[0].Rows[k][num7].ToString(), "</td>" }));
                                num7 = dstam.Tables[0].Columns["tongcong"].Ordinal - 1;
                                continue;
                            }
                        }
                        if ((num7 >= dstam.Tables[0].Columns["tongcong"].Ordinal) && (num7 <= dstam.Tables[0].Columns["chiphids"].Ordinal))
                        {
                            try
                            {
                                decimal num8 = decimal.Parse(dstam.Tables[0].Rows[k][num7].ToString());
                                builder.Append("<td " + str4 + ">" + num8.ToString("###,###,###") + "</td>");
                            }
                            catch
                            {
                                builder.Append("<td></td>");
                            }
                        }
                        else
                        {
                            builder.Append("<td " + str4 + ">" + dstam.Tables[0].Rows[k][num7].ToString() + "</td>");
                        }
                    }
                    else if ((num7 >= dstam.Tables[0].Columns["tongcong"].Ordinal) && (num7 <= dstam.Tables[0].Columns["chiphids"].Ordinal))
                    {
                        try
                        {
                            decimal num9 = decimal.Parse(dstam.Tables[0].Rows[k][num7].ToString());
                            builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt\">" + num9.ToString("###,###,###") + "</td>");
                        }
                        catch
                        {
                            builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt\"></td>");
                        }
                    }
                    else
                    {
                        try
                        {
                            if (dstam.Tables[0].Rows[k][num7].ToString().IndexOf('/') > -1)
                            {
                                builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt;align: left\">" + dstam.Tables[0].Rows[k][num7].ToString().Trim(new char[] { '\'' }) + "</td>");
                            }
                            else
                            {
                                builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt;align: left\">" + dstam.Tables[0].Rows[k][num7].ToString() + "</td>");
                            }
                        }
                        catch
                        {
                            builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt;align: left\">" + dstam.Tables[0].Rows[k][num7].ToString() + "</td>");
                        }
                    }
                }
                builder.Append("</tr>");
            }
            builder.Append("</table>");
            writer.Write(builder);
            builder = new StringBuilder();
            builder.Append("<table>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            decimal num10 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["bhyttra"].ToString()), 0);
            string str5 = new numbertotext().doiraso(num10.ToString());
            builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count - 1, " style=\"font-family:", str2, ";font-size:11pt\">Số tiền d? ngh? thanh to\x00e1n (viết bằng chữ): ", str5.Substring(0, 1).ToUpper(), str5.Substring(1), " đồng chẵn.</td>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            for (int m = 0; m < (dstam.Tables[0].Columns.Count - 3); m++)
            {
                builder.Append("<td></td>");
            }
            builder.Append(string.Concat(new object[] { "<th colspan=3 style=\"font-family:", str2, "\">Ng\x00e0y ", DateTime.Now.ToString("dd"), " th\x00e1ng ", DateTime.Now.ToString("MM"), " nam ", DateTime.Now.Year, "</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            builder.Append("<th style=\"font-family:" + str2 + "\">Người lập</th>");
            builder.Append("<td></td><td></td><td></td>");
            builder.Append("<th colspan=3 style=\"font-family:" + str2 + "\">Trưởng ph\x00f2ng KHTH</th>");
            builder.Append("<td></td><td></td><td></td><td></td><td></td><td></td>");
            builder.Append("<th colspan=2 style=\"font-family:" + str2 + "\">Kế to\x00e1n trưởng</th>");
            for (int n = 0x10; n < (dstam.Tables[0].Columns.Count - 3); n++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<th colspan=3 style=\"font-family:" + str2 + "\">Thủ trưởng đơn vị</th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            builder.Append("<td align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("<td></td><td></td><td></td>");
            builder.Append("<td colspan=3 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("<td></td><td></td><td></td><td></td><td></td><td></td>");
            builder.Append("<td colspan=2 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            for (int num13 = 0x10; num13 < (dstam.Tables[0].Columns.Count - 3); num13++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<td colspan=3 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("</tr>");
            builder.Append("</table>");
            writer.Write(builder);
            writer.Close();
            try
            {
                this._sIDProcessExcelCurrent = this._lib.getid_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(dstam.Tables[0].Columns.Count - 3) + "1", this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "2");
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                int ordinal = dstam.Tables[0].Columns["bntra"].Ordinal;
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "8", this._lib.getIndex(ordinal) + "8");
                this.orange.ColumnWidth = 14;
                ordinal = dstam.Tables[0].Columns["bhyttra"].Ordinal + 1;
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "9", this._lib.getIndex(ordinal) + "9");
                this.orange.ColumnWidth = 14;
                this.orange.RowHeight = 60;
                this.orange = this.osheet.get_Range(this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "7", this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "7");
                this.orange.ColumnWidth = 12;
                ordinal = dstam.Tables[0].Columns["tongcong"].Ordinal;
                int num15 = dstam.Tables[0].Columns["bntra"].Ordinal - 1;
                for (ordinal = ordinal; ordinal <= num15; ordinal++)
                {
                    this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "9", this._lib.getIndex(ordinal) + "9");
                    this.orange.ColumnWidth = 13;
                }
                this.owb.Save();
                this.oxl.Quit();
                string idprocesslast = this._lib.getid_process_Excel();
                this._lib.f_end_process_Excel(this._sIDProcessExcelCurrent, idprocesslast);
                Process.Start(path);
            }
            catch
            {
                Process.Start(path);
            }
        }

        public void f_Ngoaitru_xuatExcel_maumoi_41(bool print, DataSet dsdulieu, string tungay, string denngay, string fontchu)
        {
            DataSet set = this.f_Ngoaitru_excel_mau41_getdata(dsdulieu, int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)));
            this.f_Ngoaitru_exp_excel_mau41_run(print, set, tungay, denngay, fontchu);
        }

        public void f_Ngoaitru_xuatExcel_maumoi_808(bool print, DataSet dsdulieu, string tungay, string denngay, string fontchu)
        {
            DataSet set = this.f_Ngoaitru_excel_mau41_getdata(dsdulieu, int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)));
            string str = "stt#hoten#namsinh#gioitinh#mathe#ma_dkbd#mabenh#ngay_vao#ngay_ra#ngaydtr#t_tongchi#t_xn#t_cdha#t_thuoc#t_mau#t_pttt#t_vtytth#t_vtyttt#t_dvktc#t_ktg#t_kham#t_vchuyen#t_bnct#t_bhtt#t_ngoaids#lydo_vv#benhkhac#noikcb#nam_qt#thang_qt#gt_tu#gt_den#diachi#tt_tngt#";
            for (int i = 0; i < set.Tables[0].Columns.Count; i++)
            {
                if (str.IndexOf(set.Tables[0].Columns[i].ColumnName + "#") == -1)
                {
                    set.Tables[0].Columns.RemoveAt(i--);
                }
            }
            this.f_Ngoaitru_exp_excel_mau808_run(print, set, tungay, denngay, fontchu);
        }

        private string f_nodexml_chitietcls(int cap, string malk, string stt, string ma, string mavattu, string manhom, string ten, string dvt, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ngaykq, string ma_pttt)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CTDV>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_DICH_VU", ma) + str + this.f_getstr_TagXML("MA_VAT_TU", mavattu) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("TEN_DICH_VU", ten) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("NGAY_KQ", ngaykq) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CTDV>");
        }

        private string f_nodexml_chitietcls_4210(int cap, string malk, string stt, string ma, string mavattu, string manhom, string ten, string dvt, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ngaykq, string ma_pttt, string goi_vtyt, string ten_vat_tu, int phamvi, string thongtinthau, string mucthanhtoantoida, decimal bntra, decimal bhtra, decimal bncungchitra, string magiuong, int muchuong)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CHI_TIET_DVKT>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_DICH_VU", ma) + str + this.f_getstr_TagXML("MA_VAT_TU", mavattu) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("GOI_VTYT", goi_vtyt) + str + this.f_getstr_TagXML("TEN_VAT_TU", ten_vat_tu) + str + this.f_getstr_TagXML("TEN_DICH_VU", ten) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("PHAM_VI", phamvi.ToString()) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("TT_THAU", thongtinthau) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("T_TRANTT", mucthanhtoantoida) + str + this.f_getstr_TagXML("MUC_HUONG", muchuong.ToString()) + str + this.f_getstr_TagXML("T_NGUONKHAC", "0") + str + this.f_getstr_TagXML("T_BNTT", bntra.ToString()) + str + this.f_getstr_TagXML("T_BHTT", bhtra.ToString()) + str + this.f_getstr_TagXML("T_BNCCT", bncungchitra.ToString()) + str + this.f_getstr_TagXML("T_NGOAIDS", "0") + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_GIUONG", magiuong) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("NGAY_KQ", ngaykq) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CHI_TIET_DVKT>");
        }

        private string f_nodexml_chitietcls_bhxh(int cap, string malk, string stt, string ma, string mavattu, string manhom, string ten, string dvt, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ngaykq, string ma_pttt)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CHI_TIET_DVKT>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_DICH_VU", ma) + str + this.f_getstr_TagXML("MA_VAT_TU", mavattu) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("TEN_DICH_VU", ten) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("NGAY_KQ", ngaykq) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CHI_TIET_DVKT>");
        }

        private string f_nodexml_chitietthuoc(int cap, string malk, string stt, string mathuoc, string manhom, string tenthuoc, string dvt, string hamluong, string duongdung, string lieudung, string sodk, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ma_pttt)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CTTHUOC>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_THUOC", mathuoc) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("ten_thuoc", tenthuoc) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("HAM_LUONG", hamluong) + str + this.f_getstr_TagXML("DUONG_DUNG", duongdung) + str + this.f_getstr_TagXML("lieu_DUNG", lieudung) + str + this.f_getstr_TagXML("SO_DANG_KY", sodk) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CTTHUOC>");
        }

        private string f_nodexml_chitietthuoc_4210(int cap, string malk, string stt, string mathuoc, string manhom, string tenthuoc, string dvt, string hamluong, string duongdung, string lieudung, string sodk, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ma_pttt, string thongtinthau, int phamvi, int muchuong, decimal bntra, decimal bhtra, decimal bncungchitra)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CHI_TIET_THUOC>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_THUOC", mathuoc) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("ten_thuoc", tenthuoc) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("HAM_LUONG", hamluong) + str + this.f_getstr_TagXML("DUONG_DUNG", duongdung) + str + this.f_getstr_TagXML("lieu_DUNG", lieudung) + str + this.f_getstr_TagXML("SO_DANG_KY", sodk) + str + this.f_getstr_TagXML("tt_thau", thongtinthau) + str + this.f_getstr_TagXML("pham_vi", phamvi.ToString()) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("muc_huong", muchuong.ToString()) + str + this.f_getstr_TagXML("T_NGUONKHAC", "0") + str + this.f_getstr_TagXML("T_BNTT", bntra.ToString()) + str + this.f_getstr_TagXML("T_BHTT", bhtra.ToString()) + str + this.f_getstr_TagXML("T_BNCCT", bncungchitra.ToString()) + str + this.f_getstr_TagXML("T_NGOAIDS", "0") + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CHI_TIET_THUOC>");
        }

        private string f_nodexml_chitietthuoc_bhxh(int cap, string malk, string stt, string mathuoc, string manhom, string tenthuoc, string dvt, string hamluong, string duongdung, string lieudung, string sodk, decimal soluong, decimal dongia, decimal tyle, decimal thanhtien, string makhoa, string mabs, string mabenh, string ngayylenh, string ma_pttt)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<CHI_TIET_THUOC>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("MA_THUOC", mathuoc) + str + this.f_getstr_TagXML("MA_NHOM", manhom) + str + this.f_getstr_TagXML("ten_thuoc", tenthuoc) + str + this.f_getstr_TagXML("DON_VI_TINH", dvt) + str + this.f_getstr_TagXML("HAM_LUONG", hamluong) + str + this.f_getstr_TagXML("DUONG_DUNG", duongdung) + str + this.f_getstr_TagXML("lieu_DUNG", lieudung) + str + this.f_getstr_TagXML("SO_DANG_KY", sodk) + str + this.f_getstr_TagXML("so_luong", soluong.ToString()) + str + this.f_getstr_TagXML("don_gia", dongia.ToString()) + str + this.f_getstr_TagXML("TYLE_TT", tyle.ToString()) + str + this.f_getstr_TagXML("THANH_TIEN", thanhtien.ToString()) + str + this.f_getstr_TagXML("MA_KHOA", makhoa) + str + this.f_getstr_TagXML("MA_BAC_SI", mabs) + str + this.f_getstr_TagXML("MA_BENH", mabenh) + str + this.f_getstr_TagXML("NGAY_YL", ngayylenh) + str + this.f_getstr_TagXML("ma_pttt", ma_pttt) + this._s_arrCap[cap] + "</CHI_TIET_THUOC>");
        }

        private string f_nodexml_thongtinbn(int cap, string malk, string sochuyentuyen, string ngaygiovao, string ngaygiora, string mabv, string maicd, string trangthai, string ketqua, string dienthoailh, string nguoilienhe)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + "<THONGTINBENHNHAN>" + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("NGAYGIOVAO", ngaygiovao) + str + this.f_getstr_TagXML("NGAYGIORA", ngaygiora) + str + this.f_getstr_TagXML("MABENHVIEN", mabv) + str + this.f_getstr_TagXML("CHANDOAN", maicd) + str + this.f_getstr_TagXML("TRANGTHAI", trangthai) + str + this.f_getstr_TagXML("KETQUA", ketqua) + str + this.f_getstr_TagXML("SODIENTHOAI_LH", dienthoailh) + str + this.f_getstr_TagXML("NGUOILIENHE", nguoilienhe) + this._s_arrCap[cap] + "</THONGTINBENHNHAN>");
        }

        private string f_nodexml_tonghopttbn(int cap, string malk, string stt, string mabn, string hoten, string ngaysinh, string gioitinh, string diachi, string mathe, string madkbd, string tungay, string denngay, string maicd, string tenbenh, string malydovv, string manoichuyen, string matainan, string ngayvao, string ngayra, string songaydt, string ketquadt, string tinhtrangrv, string ngaythanhtoan, string tyle, string tienthuoc, string tienvtyt, string tongcong, string bntra, string bhtra, string nguonkhac, string namqt, string thangqt, int loaikcb, string makhoa, string macskcb, string makhuvuc, string cannangtreem, bool bctt324)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + (bctt324 ? "<TONG_HOP>" : "<TONGHOP>") + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("ma_bn", mabn) + str + this.f_getstr_TagXML("ho_ten", hoten) + str + this.f_getstr_TagXML("ngay_sinh", ngaysinh) + str + this.f_getstr_TagXML("gioi_tinh", gioitinh) + str + this.f_getstr_TagXML("dia_chi", diachi) + str + this.f_getstr_TagXML("ma_the", mathe) + str + this.f_getstr_TagXML("ma_dkbd", madkbd) + str + this.f_getstr_TagXML("gt_the_tu", tungay) + str + this.f_getstr_TagXML("gt_the_den", denngay) + str + this.f_getstr_TagXML("ten_benh", tenbenh) + str + this.f_getstr_TagXML("ma_benh", maicd) + str + this.f_getstr_TagXML("ma_benhkhac", "") + str + this.f_getstr_TagXML("ma_lydo_vvien", malydovv) + str + this.f_getstr_TagXML("ma_noi_chuyen", manoichuyen) + str + this.f_getstr_TagXML("ma_tai_nan", matainan) + str + this.f_getstr_TagXML("ngay_vao", ngayvao) + str + this.f_getstr_TagXML("ngay_ra", ngayra) + str + this.f_getstr_TagXML("so_ngay_dtri", songaydt) + str + this.f_getstr_TagXML("ket_qua_dtri", ketquadt) + str + this.f_getstr_TagXML("tinh_trang_rv", tinhtrangrv) + str + this.f_getstr_TagXML("ngay_ttoan", ngaythanhtoan) + str + this.f_getstr_TagXML("muc_huong", tyle) + str + this.f_getstr_TagXML("t_thuoc", tienthuoc) + str + this.f_getstr_TagXML("t_vtyt", tienvtyt) + str + this.f_getstr_TagXML("t_tongchi", tongcong) + str + this.f_getstr_TagXML("t_bntt", bntra) + str + this.f_getstr_TagXML("t_bhtt", bhtra) + str + this.f_getstr_TagXML("t_nguonkhac", nguonkhac) + str + this.f_getstr_TagXML("t_ngoaids", "0") + str + this.f_getstr_TagXML("nam_qt", namqt) + str + this.f_getstr_TagXML("thang_qt", thangqt) + str + this.f_getstr_TagXML("ma_loai_kcb", loaikcb.ToString()) + str + this.f_getstr_TagXML("ma_khoa", makhoa) + str + this.f_getstr_TagXML("ma_cskcb", macskcb) + str + this.f_getstr_TagXML("ma_khuvuc", makhuvuc) + str + this.f_getstr_TagXML("ma_pttt_qt", "") + str + this.f_getstr_TagXML("can_nang", cannangtreem) + this._s_arrCap[cap] + (bctt324 ? "</TONG_HOP>" : "</TONGHOP>"));
        }

        private string f_nodexml_tonghopttbn_4210(int cap, string malk, string stt, string mabn, string hoten, string ngaysinh, string gioitinh, string diachi, string mathe, string madkbd, string tungay, string denngay, string maicd, string tenbenh, string malydovv, string manoichuyen, string matainan, string ngayvao, string ngayra, string songaydt, string ketquadt, string tinhtrangrv, string ngaythanhtoan, string tyle, string tienthuoc, string tienvtyt, string tongcong, string bntra, string bhtra, string nguonkhac, string namqt, string thangqt, int loaikcb, string makhoa, string macskcb, string makhuvuc, string cannangtreem, bool bctt324, string ngaybhyt5nam, string bncungchitra)
        {
            string str = this._s_arrCap[cap + 1];
            return (this._s_arrCap[cap] + (bctt324 ? "<TONG_HOP>" : "<TONGHOP>") + str + this.f_getstr_TagXML("ma_lk", malk) + str + this.f_getstr_TagXML("stt", stt) + str + this.f_getstr_TagXML("ma_bn", mabn) + str + this.f_getstr_TagXML("ho_ten", hoten) + str + this.f_getstr_TagXML("ngay_sinh", ngaysinh) + str + this.f_getstr_TagXML("gioi_tinh", gioitinh) + str + this.f_getstr_TagXML("dia_chi", diachi) + str + this.f_getstr_TagXML("ma_the", mathe) + str + this.f_getstr_TagXML("ma_dkbd", madkbd) + str + this.f_getstr_TagXML("gt_the_tu", tungay) + str + this.f_getstr_TagXML("gt_the_den", denngay) + str + this.f_getstr_TagXML("mien_cung_ct", ngaybhyt5nam) + str + this.f_getstr_TagXML("ten_benh", tenbenh) + str + this.f_getstr_TagXML("ma_benh", maicd) + str + this.f_getstr_TagXML("ma_benhkhac", "") + str + this.f_getstr_TagXML("ma_lydo_vvien", malydovv) + str + this.f_getstr_TagXML("ma_noi_chuyen", manoichuyen) + str + this.f_getstr_TagXML("ma_tai_nan", matainan) + str + this.f_getstr_TagXML("ngay_vao", ngayvao) + str + this.f_getstr_TagXML("ngay_ra", ngayra) + str + this.f_getstr_TagXML("so_ngay_dtri", songaydt) + str + this.f_getstr_TagXML("ket_qua_dtri", ketquadt) + str + this.f_getstr_TagXML("tinh_trang_rv", tinhtrangrv) + str + this.f_getstr_TagXML("ngay_ttoan", ngaythanhtoan) + str + this.f_getstr_TagXML("t_thuoc", tienthuoc) + str + this.f_getstr_TagXML("t_vtyt", tienvtyt) + str + this.f_getstr_TagXML("t_tongchi", tongcong) + str + this.f_getstr_TagXML("t_bntt", bntra) + str + this.f_getstr_TagXML("t_bncct", bncungchitra) + str + this.f_getstr_TagXML("t_bhtt", bhtra) + str + this.f_getstr_TagXML("t_nguonkhac", nguonkhac) + str + this.f_getstr_TagXML("t_ngoaids", "0") + str + this.f_getstr_TagXML("nam_qt", namqt) + str + this.f_getstr_TagXML("thang_qt", thangqt) + str + this.f_getstr_TagXML("ma_loai_kcb", loaikcb.ToString()) + str + this.f_getstr_TagXML("ma_khoa", makhoa) + str + this.f_getstr_TagXML("ma_cskcb", macskcb) + str + this.f_getstr_TagXML("ma_khuvuc", makhuvuc) + str + this.f_getstr_TagXML("ma_pttt_qt", "") + str + this.f_getstr_TagXML("can_nang", cannangtreem) + this._s_arrCap[cap] + (bctt324 ? "</TONG_HOP>" : "</TONGHOP>"));
        }

        private DataSet f_Noitru_excel_mau38_getdata(DataSet dsdulieu, int namqt, int thangqt)
        {
            DataSet set = new DataSet();
            set.Tables.Add("Table");
            set.Tables[0].Columns.Add(new DataColumn("sobienlai", typeof(int)));
            set.Tables[0].Columns.Add(new DataColumn("quyenso", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngaythu", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("loaidk", typeof(int)));
            set.Tables[0].Columns.Add(new DataColumn("mabn", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("hoten", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("namsinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gioitinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("diachi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mathe", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_tu", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_den", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ma_dkbd", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noidkbd", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("chandoan", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mabenh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_vao", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_ra", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngaydtr", typeof(int)));
            set.Tables[0].Columns.Add(new DataColumn("t_xn", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_cdha", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_thuoc", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_mau", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_pttt", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtytth", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_dvktc", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_kham", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_vchuyen", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_giuong", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtyttt", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_ktg", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("t_tongchi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bnct", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bhtt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("tenkp", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("traituyen", typeof(int)));
            set.Tables[0].Columns.Add(new DataColumn("tyle", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("manhomdt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("tennhomdt", typeof(string)));
            foreach (DataRow row in dsdulieu.Tables[0].Select("", "quyenso,sobienlai"))
            {
                DataRow row2 = set.Tables[0].NewRow();
                row2["sobienlai"] = row["sobienlai"].ToString();
                row2["quyenso"] = row["quyenso"].ToString();
                row2["ngaythu"] = row["ngaythu"].ToString();
                try
                {
                    row2["loaidk"] = Convert.ToInt32(row["madk"].ToString());
                }
                catch
                {
                    row2["loaidk"] = 3;
                }
                row2["mabn"] = row["mabn"].ToString();
                row2["hoten"] = row["HOTEN"].ToString();
                row2["namsinh"] = row["ngaysinh"].ToString();
                if (row["phai"].ToString().Trim() == "1")
                {
                    row2["gioitinh"] = "Nữ";
                }
                else
                {
                    row2["gioitinh"] = "Nam";
                }
                row2["diachi"] = row["diachi"].ToString();
                try
                {
                    row2["mathe"] = (row["sothe"].ToString().Length > 15) ? row["sothe"].ToString().Substring(0, row["sothe"].ToString().Length - 5) : row["sothe"].ToString();
                }
                catch
                {
                }
                row2["gt_tu"] = row["tungay"].ToString();
                row2["gt_den"] = row["denngay"].ToString();
                row2["ma_dkbd"] = row["MANOIDK"].ToString();
                row2["noidkbd"] = row["NOIDK"].ToString();
                row2["chandoan"] = row["chandoan"].ToString();
                row2["mabenh"] = row["MAICD"].ToString();
                row2["ngay_vao"] = row["NGAYVAO"].ToString();
                row2["ngay_ra"] = row["NGAYRA"].ToString();
                row2["ngaydtr"] = row["songay"].ToString();
                row2["t_xn"] = row["ST_1"].ToString();
                row2["t_cdha"] = row["ST_2"].ToString();
                try
                {
                    row2["t_thuoc"] = Convert.ToDecimal(row["ST_3"].ToString());
                }
                catch
                {
                    row2["t_thuoc"] = 0;
                }
                try
                {
                    row2["t_mau"] = row["ST_4"].ToString();
                }
                catch
                {
                    row2["t_mau"] = 0;
                }
                try
                {
                    row2["t_pttt"] = row["ST_5"].ToString();
                }
                catch
                {
                    row2["t_pttt"] = 0;
                }
                try
                {
                    row2["t_vtytth"] = row["ST_6"].ToString();
                }
                catch
                {
                    row2["t_vtytth"] = 0;
                }
                try
                {
                    row2["t_dvktc"] = row["ST_7"].ToString();
                }
                catch
                {
                    row2["t_dvktc"] = 0;
                }
                try
                {
                    row2["t_ktg"] = row["ST_8"].ToString();
                }
                catch
                {
                    row2["t_ktg"] = 0;
                }
                try
                {
                    row2["t_kham"] = 0;
                }
                catch
                {
                }
                try
                {
                    row2["t_giuong"] = decimal.Parse(row["ST_9"].ToString()) + decimal.Parse(row["ST_11"].ToString());
                }
                catch
                {
                    row2["t_giuong"] = 0;
                }
                try
                {
                    row2["t_vtyttt"] = row["ST_14"].ToString();
                }
                catch
                {
                    row2["t_vtyttt"] = 0;
                }
                try
                {
                    row2["t_vchuyen"] = row["st_10"].ToString();
                }
                catch
                {
                }
                row2["t_tongchi"] = row["tongcong"].ToString();
                row2["t_bhtt"] = row["bhyttra"].ToString();
                row2["t_bnct"] = row["bntra"].ToString();
                row2["tenkp"] = row["khoa"].ToString();
                row2["traituyen"] = row["traituyen"].ToString();
                try
                {
                    row2["tyle"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                }
                catch
                {
                    row2["tyle"] = 0;
                }
                row2["manhomdt"] = row["nhom_dt_bhyt"].ToString();
                row2["tennhomdt"] = row["ten_nhomdt_bhyt"].ToString();
                set.Tables[0].Rows.Add(row2);
            }
            return set;
        }

        private DataSet f_Noitru_excel_mau41_getdata(DataSet dsdulieu, int namqt, int thangqt)
        {
            try
            {
                dsdulieu.Tables[0].Columns.Add("yyyymmdd");
            }
            catch
            {
            }
            try
            {
                dsdulieu.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            bool flag = this._tsxml.pChung_MauExcel41cot_hoten;
            if (dsdulieu.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in dsdulieu.Tables[0].Rows)
                {
                    try
                    {
                        row["manoidk"] = ((row["manoidk"].ToString().Substring(0, 1) == "0") ? "'" : "") + row["manoidk"].ToString();
                    }
                    catch
                    {
                    }
                    row["yyyymmdd"] = row["ngayra"].ToString().Substring(6, 4) + row["ngayra"].ToString().Substring(3, 2) + row["ngayra"].ToString().Substring(0, 2);
                    if (flag)
                    {
                        row["hoten"] = row["hoten"].ToString().ToLower();
                    }
                }
            }
            DataSet set = new DataSet();
            set.Tables.Add("Table");
            set.Tables[0].Columns.Add(new DataColumn("stt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("hoten", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("namsinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gioitinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mathe", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ma_dkbd", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("makhoa", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mabenh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_vao", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngay_ra", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ngaydtr", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_tongchi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_xn", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_cdha", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_thuoc", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_mau", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_pttt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtytth", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vtyttt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_dvktc", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_ktg", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_kham", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vchuyen", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bnct", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_bhtt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_ngoaids", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("lydo_vv", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("benhkhac", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noikcb", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("nam_qt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("thang_qt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_tu", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("gt_den", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("diachi", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("giamdinh", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_xuattoan", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("lydo_xt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_datuyen", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("t_vuottran", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("loaikcb", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("noi_ttoan", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("tt_tngt", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("mabn", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("sobienlai", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("quyenso", typeof(string)));
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "t_tongchi", "ngaydtr", "t_xn", "t_cdha", "t_thuoc", "t_mau", "t_pttt", "t_vtytth", "t_vtyttt", "t_dvktc", "t_ktg", "t_kham", "t_vchuyen", "t_bnct", "t_bhtt", "t_ngoaids" };
            string str = "";
            DataRow row2 = set.Tables[0].NewRow();
            DataRow row3 = set.Tables[0].NewRow();
            int num = 0;
            bool flag2 = this._tsxml.pNoitru_MauExcel41cot_noikcb;
            bool flag3 = !this._tsxml.pNoitru_MauExcel41cot_groupnhom;
            bool flag4 = this._tsxml.pChung_MauExcel41cot_SapXepNgay;
            bool flag5 = this._tsxml.pChung_MauExcel41cot_TheBHYTBo5KiTuCuoi;
            string filterExpression = "";
            string sort = "";
            for (int i = 1; i <= 3; i++)
            {
                switch (i)
                {
                    case 1:
                        str = "ABỆNH nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str = "BBỆNH nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str = "CBỆNH nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row4 = set.Tables[0].NewRow();
                row4["stt"] = str.Substring(0, 1);
                row4["hoten"] = str.Substring(1);
                if (flag3)
                {
                    set.Tables[0].Rows.Add(row4);
                }
                DataRow row5 = set.Tables[0].NewRow();
                for (int j = 0; j <= 1; j++)
                {
                    DataRow row6 = set.Tables[0].NewRow();
                    row6["stt"] = (j == 0) ? "I" : "II";
                    row6["hoten"] = (j == 0) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN";
                    if (flag3)
                    {
                        set.Tables[0].Rows.Add(row6);
                    }
                    DataRow row7 = set.Tables[0].NewRow();
                    if (flag3)
                    {
                        filterExpression = "madk=" + i.ToString() + " and traituyen=" + j.ToString();
                        sort = "yyyymmdd " + (flag4 ? " asc" : " desc");
                    }
                    else
                    {
                        filterExpression = "";
                        sort = "quyenso,sobienlai,yyyymmdd " + (flag4 ? " asc" : " desc");
                    }
                    foreach (DataRow row8 in dsdulieu.Tables[0].Select(filterExpression, sort))
                    {
                        DataRow row9 = set.Tables[0].NewRow();
                        row9["stt"] = row8["STT"].ToString();
                        row9["hoten"] = row8["HOTEN"].ToString();
                        row9["mathe"] = !flag5 ? row8["sothe"].ToString() : row8["sothe"].ToString().Substring(0, row8["sothe"].ToString().Length - 5);
                        row9["ma_dkbd"] = row8["MANOIDK"].ToString();
                        row9["mabenh"] = row8["MAICD"].ToString();
                        row9["benhkhac"] = row8["maicdkt"].ToString();
                        row9["makhoa"] = row8["khoa"].ToString();
                        row9["namsinh"] = row8["ngaysinh"].ToString();
                        row9["mabn"] = row8["mabn"].ToString();
                        row9["tt_tngt"] = (row8["tngt"].ToString() != "") ? "TNGT" : "";
                        if (row8["phai"].ToString().Trim() == "1")
                        {
                            row9["gioitinh"] = 2;
                        }
                        else
                        {
                            row9["gioitinh"] = 1;
                        }
                        row9["ngay_vao"] = row8["NGAYvv"].ToString();
                        row9["ngay_ra"] = row8["NGAYrv"].ToString();
                        row9["ngaydtr"] = row8["songay"].ToString();
                        row9["t_xn"] = row8["ST_1"].ToString();
                        row9["t_cdha"] = row8["ST_2"].ToString();
                        try
                        {
                            row9["t_thuoc"] = Convert.ToDecimal(row8["ST_3"].ToString());
                        }
                        catch
                        {
                            row9["t_thuoc"] = 0;
                        }
                        try
                        {
                            row9["t_mau"] = row8["ST_4"].ToString();
                        }
                        catch
                        {
                            row9["t_mau"] = 0;
                        }
                        try
                        {
                            row9["t_pttt"] = row8["ST_5"].ToString();
                        }
                        catch
                        {
                            row9["t_pttt"] = 0;
                        }
                        try
                        {
                            row9["t_vtytth"] = row8["ST_6"].ToString();
                        }
                        catch
                        {
                            row9["t_vtytth"] = 0;
                        }
                        try
                        {
                            row9["t_dvktc"] = row8["ST_7"].ToString();
                        }
                        catch
                        {
                            row9["t_dvktc"] = 0;
                        }
                        try
                        {
                            row9["t_ktg"] = row8["ST_8"].ToString();
                        }
                        catch
                        {
                            row9["t_ktg"] = 0;
                        }
                        try
                        {
                            row9["t_kham"] = decimal.Parse(row8["ST_9"].ToString()) + decimal.Parse(row8["ST_11"].ToString());
                        }
                        catch
                        {
                            row9["t_kham"] = 0;
                        }
                        try
                        {
                            row9["t_vtyttt"] = row8["ST_14"].ToString();
                        }
                        catch
                        {
                            row9["t_vtyttt"] = 0;
                        }
                        try
                        {
                            row9["t_vchuyen"] = row8["st_10"].ToString();
                        }
                        catch
                        {
                        }
                        row9["t_tongchi"] = row8["tongcong"].ToString();
                        row9["t_bhtt"] = row8["bhyttra"].ToString();
                        row9["t_bnct"] = row8["bntra"].ToString();
                        row9["t_ngoaids"] = 0;
                        if (row8["madoituong"].ToString() == "6")
                        {
                            row9["t_ngoaids"] = row8["bhyttra"].ToString();
                        }
                        if (row8["traituyen"].ToString().Trim() == "0")
                        {
                            row9["lydo_vv"] = 1;
                        }
                        if (row8["traituyen"].ToString().Trim() == "1")
                        {
                            row9["lydo_vv"] = 0;
                        }
                        if (row8["capcuu"].ToString() == "1")
                        {
                            row9["lydo_vv"] = 2;
                        }
                        try
                        {
                            if (row8["makp"].ToString().Trim() == "99")
                            {
                                row9["lydo_vv"] = 2;
                            }
                        }
                        catch
                        {
                        }
                        try
                        {
                            row9["nam_qt"] = namqt.ToString();
                            row9["thang_qt"] = thangqt.ToString();
                        }
                        catch
                        {
                        }
                        row9["noikcb"] = flag2 ? row9["ma_dkbd"].ToString() : this._lib.MABHXH;
                        row9["diachi"] = row8["diachi"].ToString();
                        row9["gt_tu"] = row8["tungay"].ToString();
                        row9["gt_den"] = row8["denngay"].ToString();
                        row9["loaikcb"] = "NOI";
                        row9["sobienlai"] = row8["sobienlai"].ToString();
                        row9["quyenso"] = row8["quyenso"].ToString();
                        row9["stt"] = ++num;
                        set.Tables[0].Rows.Add(row9);
                        for (int k = 0; k < strArray2.Length; k++)
                        {
                            try
                            {
                                if (row7[strArray2[k]].ToString() == "")
                                {
                                    row7[strArray2[k]] = 0;
                                }
                                if (row2[strArray2[k]].ToString() == "")
                                {
                                    row2[strArray2[k]] = 0;
                                }
                                if (row5[strArray2[k]].ToString() == "")
                                {
                                    row5[strArray2[k]] = 0;
                                }
                                if (row9[strArray2[k]].ToString() == "")
                                {
                                    row9[strArray2[k]] = 0;
                                }
                                row7[strArray2[k]] = decimal.Parse(row7[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                                row2[strArray2[k]] = decimal.Parse(row2[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                                row5[strArray2[k]] = decimal.Parse(row5[strArray2[k]].ToString()) + decimal.Parse(row9[strArray2[k]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    if (!flag3)
                    {
                        break;
                    }
                    row7["stt"] = (j == 0) ? "I" : "II";
                    row7["hoten"] = "TỔNG " + ((j == 0) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN");
                    if ((row7["t_tongchi"].ToString() != "") && (row7["t_tongchi"].ToString() != "0"))
                    {
                        set.Tables[0].Rows.Add(row7);
                    }
                    else
                    {
                        set.Tables[0].Rows.Remove(row6);
                    }
                }
                if (!flag3)
                {
                    break;
                }
                row5["stt"] = str.Substring(0, 1);
                row5["hoten"] = "cộng " + row5["stt"].ToString();
                if ((row5["t_tongchi"].ToString() != "") && (row5["t_tongchi"].ToString() != "0"))
                {
                    set.Tables[0].Rows.Add(row5);
                }
                else
                {
                    set.Tables[0].Rows.RemoveAt(set.Tables[0].Rows.Count - 1);
                }
            }
            row2["hoten"] = "Tổng cộng A+B+C";
            set.Tables[0].Rows.Add(row2);
            try
            {
                dsdulieu.Tables[0].Columns.Remove("yyyymmdd");
            }
            catch
            {
            }
            return set;
        }

        private void f_Noitru_exp_excel_mau38_run(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            int num = 0;
            int num2 = 3;
            int num3 = 5;
            int num4 = ds11.Tables[0].Rows.Count + 5;
            int i = ds11.Tables[0].Columns.Count - 1;
            num = num4;
            this.tenfile = this._lib.Export_Excel(ds11, "bccpkcb");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num2; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(ds11.Tables[0].Columns["ngaydtr"].Ordinal) + num3.ToString(), this._lib.getIndex(ds11.Tables[0].Columns["t_bhtt"].Ordinal) + num4.ToString()).NumberFormat = "#,##0";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(i) + num).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num4, this._lib.getIndex(i + 3) + num4);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Bold = true;
                this.orange.RowHeight = 15;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(i + 2) + num4.ToString());
                this.orange.Font.Name = fontchu;
                this.orange.Font.Size = 8;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = false;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 3] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NỘI TR\x00da";
                int num7 = num3;
                string[] strArray = new string[] { "mabn", "ma_dkbd" };
                for (int k = 0; k < strArray.Length; k++)
                {
                    try
                    {
                        int ordinal = ds11.Tables[0].Columns[strArray[k]].Ordinal;
                        this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + num3, this._lib.getIndex(ordinal) + num4);
                        this.orange.NumberFormat = "@";
                    }
                    catch
                    {
                    }
                }
                for (int m = 0; m < ds11.Tables[0].Rows.Count; m++)
                {
                    for (int num11 = 0; num11 < strArray.Length; num11++)
                    {
                        int num12 = ds11.Tables[0].Columns[strArray[num11]].Ordinal;
                        try
                        {
                            this.osheet.Cells[num3 + m, num12 + 1] = ds11.Tables[0].Rows[m][num12].ToString();
                        }
                        catch
                        {
                        }
                    }
                }
                string[] strArray2 = new string[] { "t_xn", "t_cdha", "t_thuoc", "t_mau", "t_pttt", "t_vtytth", "t_dvktc", "t_kham", "t_vanchuyen", "t_giuong", "t_vtyttt", "t_ktg", "t_tongchi", "t_bnct", "t_bhtt" };
                for (int n = 0; n < strArray2.Length; n++)
                {
                    try
                    {
                        int num14 = ds11.Tables[0].Columns[strArray2[n]].Ordinal;
                        this.osheet.Cells[num4, num14 + 1] = "=sum(" + (this._lib.getIndex(num14) + num3) + ":" + (this._lib.getIndex(num14) + (num4 - 1)) + ")";
                    }
                    catch
                    {
                    }
                }
                strArray = new string[] { 
                    "SBL", "Quyển sổ", "NG\x00c0Y THU", "LOẠI ĐK", "M\x00c3 BN", "HỌ T\x00caN", "NĂM SINH", "PH\x00c1I", "ĐỊA CHỈ", "SỐ THẺ", "TỪ NG\x00c0Y", "ĐẾN NG\x00c0Y", "M\x00c3 ĐKKCB", "NƠI ĐKKCB", "CHẨN ĐO\x00c1N", "M\x00c3 ICD", 
                    "NG\x00c0Y V\x00c0O", "NG\x00c0Y RA", "SỐ NG\x00c0Y", "X\x00e9t nghiệm, TDCN", "Ch\x00e2̉n đoán hình ảnh", "Thu\x00f4́c, dịch truy\x00eàn", "Máu", "Dịch vụ kỹ thu\x00e2̣t th\x00f4ng thường", "V\x00e2̣t tư y t\x00eá", "Dịch vụ kỹ thu\x00e2̣t cao", "C\x00f4ng khám", "Chi ph\x00ed v\x00e2̣n chuy\x00eản", "Ti\x00eàn giường", "Vật tư thay thế", "Thuốc K, chống thải gh\x00e9p", "TỔNG CỘNG", 
                    "BN THANH TO\x00c1N", "BHYT THANH TO\x00c1N", "KHOA", "TR\x00c1I TUYẾN", "TỶ LỆ BHYT TRẢ", "M\x00c3 NH\x00d3M ĐT", "NH\x00d3M ĐT"
                 };
                for (int num15 = 0; num15 < strArray.Length; num15++)
                {
                    this.osheet.Cells[num3 - 1, num15 + 1] = strArray[num15];
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(i) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.osheet.Cells[2, 3] = (tungay != denngay) ? ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay) : ("Ng\x00e0y " + tungay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "2", this._lib.getIndex(i) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception.Message, this._lib.Msg);
            }
        }

        private void f_Noitru_exp_excel_mau41_run(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            ds11 = this.f_setSapXepCotTheoThuTu(ds11, this._tsxml.pNoiTru_MauExcel41cot_DanhSachCotHienThi, '#');
            DataRow row = ds11.Tables[0].NewRow();
            for (int i = 0; i < ds11.Tables[0].Columns.Count; i++)
            {
                row[i] = "[" + (i + 1) + "]";
            }
            ds11.Tables[0].Rows.InsertAt(row, 0);
            int num2 = 0;
            int num3 = 3;
            int num4 = 5;
            int num5 = ds11.Tables[0].Rows.Count + 5;
            int num6 = ds11.Tables[0].Columns.Count - 1;
            num2 = num5;
            this.tenfile = this._lib.Export_Excel(ds11, "bccpkcb");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num3; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num3 + 7) + num4.ToString(), this._lib.getIndex(ds11.Tables[0].Columns["t_bhtt"].Ordinal) + num5.ToString()).NumberFormat = "#,##0";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(num6) + (num2 - 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num5, this._lib.getIndex(num6 + 3) + num5);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Bold = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num4, this._lib.getIndex(num6 + 3) + num4);
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.Font.Bold = true;
                this.orange.RowHeight = 15;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(num6 + 2) + num5.ToString());
                this.orange.Font.Name = fontchu;
                this.orange.Font.Size = 12;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 3] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NỘI TR\x00da";
                int num8 = num4;
                for (int k = 1; k < ds11.Tables[0].Rows.Count; k++)
                {
                    num8++;
                    this.orange = this.osheet.get_Range("A" + num8.ToString(), this._lib.getIndex(num6 - 1) + num8.ToString());
                    if (((ds11.Tables[0].Rows[k]["stt"].ToString() == "A") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((ds11.Tables[0].Rows[k]["stt"].ToString() == "C") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "I") || (ds11.Tables[0].Rows[k]["stt"].ToString() == "II"))
                    {
                        this.orange.Font.ColorIndex = 10;
                        this.orange.Font.Bold = true;
                    }
                    if (ds11.Tables[0].Rows[k]["t_tongchi"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range("B" + num8, this._lib.getIndex(num6) + num8);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                    else if ((ds11.Tables[0].Rows[k]["stt"].ToString() == "") || !char.IsDigit(ds11.Tables[0].Rows[k]["stt"].ToString(), 0))
                    {
                        this.orange = this.osheet.get_Range("B" + num8, this._lib.getIndex(ds11.Tables[0].Columns["ngay_ra"].Ordinal) + num8);
                        this.orange.Font.Bold = true;
                        this.orange.MergeCells = true;
                    }
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(num6) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.osheet.Cells[2, 3] = "Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "2", this._lib.getIndex(num6) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception.Message, this._lib.Msg);
            }
        }

        private void f_Noitru_exp_excel_mau808_run(bool print, DataSet ds11, string tungay, string denngay, string fontchu)
        {
            this.f_exp_excel_mau808_run(print, ds11, 1, tungay, denngay, fontchu);
        }

        public string f_Noitru_getSql_chiphibn_theokhoa(string mmyy, string tungay, string denngay, string sothe, string makp, string madoituong, string vitritheBHYT, bool tunguyen)
        {
            string str = "";
            string user = this._lib.user;
            string str3 = user + mmyy;
            str = "select 0 loai,b.madoituong,a.id,a.mabn,g.hoten,g.sonha||' '||g.thon||' '||g.cholam diachi,g.namsinh,g.phai as gioitinh,nvl(c.sothe,' ') sothe,c.noicap mabv,h.tenbv,a.chandoan,a.maicd,a1.maicd maicdkt,a1.chandoan chandoankt,a2.mabn tngt,to_char(a.ngayra,'dd/mm/yyyy') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao,to_char(a.ngayra-a.ngayvao,'dd') as songay,to_char(a.ngayra,'dd/mm/yyyy') as ngay ,i.id nhomvp,b.sotien,0 as sobienlai, d.kythuat, to_char(a.ngayvao,'dd/mm/yyyy') as ngaythu, a.mabn as quyenso ";
            str = str + " from " + str3 + ".v_thvpll a inner join " + str3 + ".v_thvpct b on a.id=b.id  inner join " + str3 + ".v_thvpbhyt c on a.id=c.id  inner join " + user + ".d_dmbd d on b.mavp=d.id  inner join " + user + ".d_dmnhom e on d.manhom=e.id  inner join " + user + ".v_nhomvp f on e.nhomvp=f.ma  inner join " + user + ".btdbn g on a.mabn=g.mabn  inner join " + user + ".dmnoicapbhyt h on c.noicap=h.mabv  inner join " + user + ".v_nhombhyt i on f.idnhombhyt=i.id left join " + user + ".cdkemtheo a1 on a1.maql=a.maql  left join " + user + ".tainantt a2 on a2.maql=a.maql ";
            str = str + " where  b.madoituong in(" + madoituong + ") and a.ngayra between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy') and c.sothe is not null";
            if (tunguyen && (sothe != ""))
            {
                str = str + " and substr(c.sothe," + vitritheBHYT + ",2) not in (" + sothe + ")";
            }
            else if (sothe != "")
            {
                string str4 = str;
                str = str4 + " and substr(c.sothe," + vitritheBHYT + ",2) in (" + sothe + ")";
            }
            if (makp != "")
            {
                str = str + " and b.makp in(" + makp + ")";
            }
            string str5 = str + " union all select 1 loai,b.madoituong,a.id,a.mabn,g.hoten,g.sonha||' '||g.thon||' '||g.cholam diachi,g.namsinh,g.phai as gioitinh,nvl(c.sothe,' ') sothe,c.noicap mabv,h.tenbv,a.chandoan,a.maicd,a1.maicd maicdkt,a1.chandoan chandoankt,a2.mabn tngt,to_char(a.ngayra,'dd/mm/yyyy') as ngayra,to_char(a.ngayvao,'dd/mm/yyyy') as ngayvao,to_char(a.ngayra-a.ngayvao,'dd') as songay,to_char(a.ngayra,'dd/mm/yyyy') as ngay, i.id nhomvp,b.sotien,0 as sobienlai, d.kythuat, to_char(a.ngayvao,'dd/mm/yyyy') as ngaythu, a.mabn as quyenso ";
            str = str5 + " from " + str3 + ".v_thvpll a inner join " + str3 + ".v_thvpct b on a.id=b.id  inner join " + str3 + ".v_thvpbhyt c on a.id=c.id inner join " + user + ".v_giavp d on b.mavp=d.id inner join " + user + ".v_loaivp e d.id_loai=e.id inner join " + user + ".v_nhomvp f on e.id_nhom=f.ma inner join " + user + ".btdbn g on a.mabn=g.mabn inner join " + user + ".dmnoicapbhyt h on c.noicap=h.mabv inner join " + user + ".v_nhombhyt i on f.idnhombhyt=i.id left join " + user + ".cdkemtheo a1 on a.maql=a1.maql left join " + user + ".tainantt a2 on a.maql=a2.maql where and b.madoituong in(" + madoituong + ")";
            str = str + " and a.ngayra between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy') and c.sothe is not null";
            if (tunguyen && (sothe != ""))
            {
                string str6 = str;
                str = str6 + " and substr(c.sothe," + vitritheBHYT + ",2) not in (" + sothe + ")";
            }
            else if (sothe != "")
            {
                str = str + " and substr(c.sothe," + vitritheBHYT + ",2) in (" + sothe + ")";
            }
            if (makp != "")
            {
                str = str + " and b.makp in(" + makp + ")";
            }
            return str;
        }

        public string f_Noitru_getSql_chiphibn_theovp(string d_mmyy, string tungay, string denngay, string madoituong, string makp, string mathe, string quyenso, string userthutien, int Nhombaocao, bool thuoc, bool vienphi, bool BNHoanTraBHYT, bool TinhCPPhongLuuNhuNoiTru, bool LayCPPhongLuu, bool LaySLBNNoiVien, bool LaySLBNNgoaiVien)
        {
            string str = "";
            string str2 = "";
            string str3 = "";
            string user = this._lib.user;
            string str5 = user + d_mmyy;
            string str6 = (int.Parse(this._sViTriTheMoi.Substring(0, this._sViTriTheMoi.IndexOf(","))) + 1) + "," + int.Parse(this._sViTriTheMoi.Substring(this._sViTriTheMoi.IndexOf(",") + 1));
            if (Nhombaocao == 0)
            {
                str2 = "e.id";
                str3 = "d.idnhombhytmedisoft";
            }
            else
            {
                str2 = "e.id";
                str3 = "b.ma";
            }
            string str7 = "";
            if ((LaySLBNNgoaiVien && LaySLBNNoiVien) || (!LaySLBNNgoaiVien && !LaySLBNNoiVien))
            {
                str7 = "";
            }
            else if (LaySLBNNgoaiVien)
            {
                str7 = " <>'" + this._lib.Mabv + "'";
            }
            else
            {
                str7 = " ='" + this._lib.Mabv + "'";
            }
            string str8 = this.f_get_sql_theBHYT(1, d_mmyy);
            string str9 = "select e.id  from " + str5 + ".v_hoantra f left join " + str5 + ".v_ttrvll e on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai  inner join " + str5 + ".v_ttrvds g on e.id=g.id and g.mabn=f.mabn  where  e.ngay between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')";
            string str10 = "select a.id,a.sothe,decode(dv.nhombv,null,a.mabv,dv.nhombv) mabv,b.tenbv,a.maphu,a.traituyen  from " + str5 + ".v_ttrvbhyt a inner join " + user + ".dmnoicapbhyt b on a.mabv=b.mabv left join " + user + ".dmnhomdv_bhyt dv on a.mabv=dv.mabv ";
            string str11 = "";
            string str12 = "";
            if (thuoc)
            {
                string str13 = "select " + str3 + " id,b.ma *- 1 as loaivp,c.id mavp, c.kythuat  from " + user + ".d_dmnhom a inner join " + user + ".v_nhomvp b on a.nhomvp=b.ma  inner join " + user + ".d_dmbd c on a.id=c.manhom inner join " + user + ".v_nhombhyt d on b.idnhombhyt=d.id ";
                str12 = " select 0 loai,a.maql,d.madoituong,a.id,a.mabn,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a2.nguyennhan=0 then '1' else '' end tngt,to_char(a.ngayvao,'dd/mm/yyyy') ngayvao,to_char(f.ngay,'dd/mm/yyyy') ngay, to_char(a.ngayra,'dd/mm/yyyy') ngayra,to_date(a.ngayra,'dd/mm/yy')-to_date(a.ngayvao,'dd/mm/yy') as songay,g.traituyen,b.hoten,b.namsinh,b.phai gioitinh,g.sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,d.mavp," + str2 + " nhomvp,e.loaivp,d.soluong*d.dongia sotien,g.mabv,g.tenbv,b.sonha||' '||b.thon||' '||b.cholam||','||xa.tenpxa||','||qu.tenquan||','||tt.tentt diachi,d.makp,kp.tenkp, g.maphu, f.sobienlai, e.kythuat,a1.nhantu , to_char(f.ngay,'dd/mm/yyyy') ngaythu, qs.sohieu as quyenso , d.bhyttra as a_bhyttra";
                str12 = str12 + " from " + str5 + ".v_ttrvds a  inner join " + str5 + ".v_ttrvll f on  a.id=f.id  inner join (select distinct nhantu,maql,mabn from " + user + ".benhandt) a1 on a.maql=a1.maql and a.mabn=a1.mabn  left join (" + str8 + ") bh on a1.maql=bh.maql left join (select distinct nguyennhan,maql from " + user + ".tainantt) a2 on a1.maql=a2.maql  left join (select distinct maql,makp from " + user + ".xuatvien) a3 on a1.maql=a3.maql  left join (" + str9 + ") i on f.id=i.id left join (" + str10 + ") g on  a.id=g.id inner join " + str5 + ".v_ttrvct d on a.id=d.id inner join  (" + str13 + ") e on  d.mavp=e.mavp   left join " + user + ".btdkp_bv kp on kp.makp=d.makp inner join " + user + ".v_quyenso qs on  f.quyenso=qs.id inner join " + user + ".btdbn b on a.mabn=b.mabn inner join " + user + ".btdtt tt on  b.matt=tt.matt inner join " + user + ".btdquan qu on  b.maqu=qu.maqu  inner join " + user + ".btdpxa xa on b.maphuongxa=xa.maphuongxa ";
                str12 = str12 + " where 1=1  and to_date(to_char(f.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')  and d.madoituong in(" + madoituong + ") ";
                if (this._tsxml.pNoitru_Mau80HD_LocMakpTheoXuatVien)
                {
                    str12 = str12 + " and a3.maql is not null ";
                }
                if (makp != "")
                {
                    if (this._tsxml.pNoitru_Mau80HD_LocMakpTheoXuatVien)
                    {
                        str12 = str12 + " and a3.makp in(" + makp + ")";
                    }
                    else
                    {
                        str12 = str12 + " and d.makp in(" + makp + ")";
                    }
                }
                if (mathe.Trim().Trim(new char[] { ',' }) != "")
                {
                    str12 = str12 + " and  (substr(upper(g.sothe)," + str6 + ") in (" + mathe.Trim().Trim(new char[] { ',' }) + ") or substr(upper(g.sothe),1,2) in (" + mathe.Trim().Trim(new char[] { ',' }) + "))";
                }
                if (quyenso.Trim().Trim(new char[] { ',' }) != "")
                {
                    str12 = str12 + " and f.quyenso in(" + quyenso.Trim().Trim(new char[] { ',' }) + ")";
                }
                if (userthutien.Trim().Trim(new char[] { ',' }) != "")
                {
                    str12 = str12 + " and f.userid in(" + userthutien.Trim().Trim(new char[] { ',' }) + ")";
                }
                if (LayCPPhongLuu)
                {
                    str12 = str12 + " and f.loaibn in (1) and f.makp<>'99'";
                }
                else if (TinhCPPhongLuuNhuNoiTru)
                {
                    str12 = str12 + " and f.loaibn in (1,4)";
                }
                else
                {
                    str12 = str12 + " and f.makp<>'99' and f.loaibn in ('1')";
                }
                if (BNHoanTraBHYT)
                {
                    str12 = str12 + " and i.id is not null ";
                }
                else
                {
                    str12 = str12 + " and i.id is null ";
                }
                if (str7 != "")
                {
                    str12 = str12 + " and g.mabv " + str7;
                }
                str = str12;
            }
            if (vienphi)
            {
                string str14 = "select " + str3 + " id,a.id loaivp,c.id mavp, c.kythuat  from " + user + ".v_loaivp a inner join " + user + ".v_nhomvp b on a.id_nhom=b.ma  inner join " + user + ".v_giavp c on a.id=c.id_loai inner join " + user + ".v_nhombhyt d on b.idnhombhyt=d.id ";
                str11 = " select 1 loai,a.maql,d.madoituong,a.id,a.mabn,a.chandoan,a.maicd,null maicdkt,null chandoankt,case when a2.nguyennhan=0 then '1' else '' end tngt,to_char(a.ngayvao,'dd/mm/yyyy') ngayvao,to_char(f.ngay,'dd/mm/yyyy') ngay, to_char(a.ngayra,'dd/mm/yyyy') ngayra,to_date(a.ngayra,'dd/mm/yy')-to_date(a.ngayvao,'dd/mm/yy') as songay,g.traituyen,b.hoten,b.namsinh,b.phai gioitinh,g.sothe,to_char(bh.tungay,'dd/mm/yyyy') tungay,to_char(bh.denngay,'dd/mm/yyyy') denngay,d.mavp," + str2 + " nhomvp,e.loaivp,d.soluong*d.dongia sotien,g.mabv,g.tenbv,b.sonha||' '||b.thon||' '||b.cholam||','||xa.tenpxa||','||qu.tenquan||','||tt.tentt diachi,d.makp,kp.tenkp, g.maphu, f.sobienlai, e.kythuat,a1.nhantu , to_char(f.ngay,'dd/mm/yyyy') ngaythu, qs.sohieu as quyenso , d.bhyttra as a_bhyttra";
                str11 = str11 + " from " + str5 + ".v_ttrvds a  inner join (select distinct nhantu,maql,mabn from " + user + ".benhandt) a1 on a.maql=a1.maql and a1.mabn=a.mabn left join (select distinct nguyennhan,maql from " + user + ".tainantt) a2 on a1.maql=a2.maql  left join (select distinct maql,makp from " + user + ".xuatvien) a3 on a1.maql=a3.maql  left join (" + str8 + ") bh on a1.maql=bh.maql inner join " + str5 + ".v_ttrvct d on a.id=d.id inner join  (" + str14 + ") e on  d.mavp=e.mavp   inner join " + str5 + ".v_ttrvll f on  a.id=f.id  left join " + user + ".btdkp_bv kp on kp.makp=d.makp inner join " + user + ".v_quyenso qs on  f.quyenso=qs.id left join (" + str9 + ") i on f.id=i.id left join (" + str10 + ") g on  a.id=g.id inner join " + user + ".btdbn b on a.mabn=b.mabn inner join " + user + ".btdtt tt on  b.matt=tt.matt inner join " + user + ".btdquan qu on  b.maqu=qu.maqu  inner join " + user + ".btdpxa xa on b.maphuongxa=xa.maphuongxa ";
                str11 = str11 + " where 1=1  and to_date(to_char(f.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')  and d.madoituong in(" + madoituong + ") ";
                if (this._tsxml.pNoitru_Mau80HD_LocMakpTheoXuatVien)
                {
                    str12 = str12 + " and a3.maql is not null ";
                }
                if (makp != "")
                {
                    if (this._tsxml.pNoitru_Mau80HD_LocMakpTheoXuatVien)
                    {
                        str11 = str11 + " and a3.makp in(" + makp + ")";
                    }
                    else
                    {
                        str11 = str11 + " and d.makp in(" + makp + ")";
                    }
                }
                if (mathe.Trim().Trim(new char[] { ',' }) != "")
                {
                    str11 = str11 + " and  (substr(upper(g.sothe)," + str6 + ") in (" + mathe.Trim().Trim(new char[] { ',' }) + ") or substr(upper(g.sothe),1,2) in (" + mathe.Trim().Trim(new char[] { ',' }) + "))";
                }
                if (quyenso.Trim().Trim(new char[] { ',' }) != "")
                {
                    str11 = str11 + " and f.quyenso in(" + quyenso.Trim().Trim(new char[] { ',' }) + ")";
                }
                if (userthutien.Trim().Trim(new char[] { ',' }) != "")
                {
                    str11 = str11 + " and f.userid in(" + userthutien.Trim().Trim(new char[] { ',' }) + ")";
                }
                if (LayCPPhongLuu)
                {
                    str11 = str11 + " and f.loaibn in (1) and f.makp<>'99'";
                }
                else if (TinhCPPhongLuuNhuNoiTru)
                {
                    str11 = str11 + " and f.loaibn in (1,4)";
                }
                else
                {
                    str11 = str11 + " and f.makp<>'99' and f.loaibn in ('1')";
                }
                if (BNHoanTraBHYT)
                {
                    str11 = str11 + " and i.id is not null ";
                }
                else
                {
                    str11 = str11 + " and i.id is null ";
                }
                if (str7 != "")
                {
                    str11 = str11 + " and g.mabv " + str7;
                }
                str = str + " # " + str11;
            }
            if (!thuoc && !vienphi)
            {
                str = str12 + " # " + str11;
            }
            return str.Trim(new char[] { '#' }).Replace("#", " union all ");
        }

        public string f_NoiTru_getSql_CLS(string mmyy, string tungay, string denngay, string madoituong, string makp, string nhomvp_bhyt_medi, string loaibenhan, bool LaySLVPKhoa, string mavienphi, string userid, string mabv, string notmabv, bool LaySLTheoBenhNhan, bool LayTheoTT2348, string bc_mabn, string bc_quyenso)
        {
            StringBuilder builder = new StringBuilder();
            if (bc_mabn != "")
            {
                bc_mabn = "'" + bc_mabn.Trim(new char[] { ',' }).Replace(",", "','") + "'";
            }
            string str = "3";
            string user = this._lib.user;
            string str3 = user + "d" + mmyy;
            string str4 = user + mmyy;
            string str5 = "";
            string str6 = "select e.id from " + str4 + ".v_ttrvll e  left join " + str4 + ".v_hoantra f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai inner join " + str4 + ".v_ttrvds g on e.id=g.id where  g.mabn=f.mabn and e.ngay between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')";
            if (mavienphi != "")
            {
                str5 = mavienphi.Trim(new char[] { ',' });
                str5 = " and c.id in(" + str5 + ")";
            }
            if (LaySLVPKhoa)
            {
                builder.Append(" select xv.loaiba as loaiba, c.manhom,bh.idnhombhytmedisoft as nhombhyt, 0 as stt,soft.ten as tennhombhyt,1 as loai,0 as congkham, d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,to_char(ab.ngay,'dd/mm/yyyy') ngayduyet, sum(a.soluong) as soluong, a.dongia, sum(a.soluong*a.dongia) sotien , sum(a.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                builder.Append(" from " + str4 + ".v_thvpll ab inner join " + str4 + ".v_thvpct a on ab.id=a.id  inner join " + str4 + ".v_thvpbhyt a1 on a.id=a1.id   inner join " + user + ".d_dmbd c on a.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id  inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  left join " + user + ".tenvien c1 on c1.mabv=c.mabv   left join " + user + ".d_dmnx c2 on c2.id=c.madv   left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id  left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id  left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id  left join " + user + ".btdkp_bv bv on a.makp=bv.makp  left join " + user + ".benhandt xv on ab.maql=xv.maql  left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on ab.maql=bh.maql ");
                builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and xv.loaiba in (" + loaibenhan + ") and a1.sothe is not null ");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str5 + "");
                builder.Append(" group by ab.ngay,xv.loaiba,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct" + (LaySLTheoBenhNhan ? ",ab.mabn" : ""));
                builder.Append(" union all ");
                builder.Append(" select xv.loaiba as loaiba,d.id_nhom as manhom,bh.idnhombhytmedisoft as nhombhyt,0 as stt,  soft.ten as tennhombhyt,1 as loai,0 as congkham ,e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append("to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,  sum(a.soluong) as soluong, a.dongia as dongia,  sum(a.soluong*a.dongia)as sotien , sum(a.bhyttra)as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                builder.Append("  from " + str4 + ".v_thvpll ab inner join " + str4 + ".v_thvpct a on ab.id=a.id  ");
                builder.Append("  inner join " + str4 + ".v_thvpbhyt b on a.id=b.id inner join " + user + ".v_giavp c on a.mavp=c.id ");
                builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft = soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                builder.Append(" left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on ab.maql=bh.maql ");
                builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and xv.loaiba in (" + loaibenhan + ") and b.sothe is not null ");
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and bh.idnhombhytmedisoft not in(" + str + ")" + str5 + "");
                builder.Append(" group by ab.ngay,xv.loaiba,c.ma,bh.idnhombhytmedisoft,soft.ten ,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, a.dongia, c.kythuat " + (LaySLTheoBenhNhan ? ",ab.mabn" : ""));
            }
            else
            {
                builder.Append(" select a.loaibn as loaiba,c.manhom, nhombh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,1 as loai,0 as congkham,  d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c1.mabv,d.thuocyhct,to_char(a.ngay,'dd/mm/yyyy') ngayduyet, sum(a1.soluong) as soluong,  a1.dongia,  sum(a1.soluong*a1.dongia) sotien , sum(a1.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp,bh.mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bh.malk,badt.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append(" from " + str4 + ".v_ttrvds ab inner join " + str4 + ".v_ttrvll a on ab.id=a.id  inner join " + str4 + ".v_ttrvct a1 on a.id=a1.id   inner join " + user + ".d_dmbd c on a1.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id  inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  left join " + user + ".tenvien c1 on c1.mabv=c.mabv   left join " + user + ".d_dmnx c2 on c2.id=c.madv   left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id  left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id  left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id  left join " + user + ".btdkp_bv bv on a.makp=bv.makp  inner join " + user + ".benhandt badt on badt.maql=ab.maql  and badt.mabn=ab.mabn left join (select distinct mabv,maql,malk from (" + this.f_get_sql_theBHYT(1, mmyy) + ")) bh on badt.maql=bh.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on badt.mabs=b21.ma") : ""));
                builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and a.loaibn in(" + loaibenhan + ") and ab.id not in (" + str6 + ") ");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" and nhombh.idnhombhytmedisoft not in(" + str + ")" + str5 + "");
                builder.Append(" group by a.ngay,a.loaibn,nhombh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a1.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c1.mabv,d.thuocyhct,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat ,ab.maql,bh.malk,badt.mavaovien" : ""));
                builder.Append(" union all ");
                builder.Append(" select a.loaibn as loaiba,d.id_nhom as manhom, nhombh.idnhombhytmedisoft as nhombhyt,0 as stt,soft.ten as tennhombhyt,1 as loai,0 as congkham ,e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet, sum(b.soluong) as soluong, b.dongia as dongia,   sum(b.soluong*b.dongia)as sotien ,case when soft.id=10 then sum(b.soluong*b.dongia) else sum(b.bhyttra) end as bhyttra,to_char(c.ma) as mavp1,a.makp,bv.tenkp,bh.mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bh.malk,badt.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append("  from " + str4 + ".v_ttrvds ab inner join " + str4 + ".v_ttrvll a on ab.id=a.id    inner join " + str4 + ".v_ttrvct b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id  inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma  left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id  left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft = soft.id  left join " + user + ".btdkp_bv bv on a.makp=bv.makp  inner join " + user + ".benhandt badt on badt.maql=ab.maql and badt.mabn=ab.mabn left join (select distinct mabv,maql,malk from (" + this.f_get_sql_theBHYT(1, mmyy) + ")) bh on badt.maql=bh.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on badt.mabs=b21.ma") : ""));
                builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and a.loaibn in (" + loaibenhan + ") and ab.id not in (" + str6 + ") ");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and nhombh.idnhombhytmedisoft not in(" + str + ")" + str5 + "");
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" group by a.ngay,a.loaibn,c.ma,nhombh.idnhombhytmedisoft,soft.ten,soft.id ,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat ,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,ab.maql,bh.malk,badt.mavaovien" : ""));
            }
            return builder.ToString();
        }

        public string f_NoiTru_getSql_thuoc(string mmyy, string tungay, string denngay, string madoituong, string makp, string nhomvp_bhyt_medi, string loaibenhan, bool LaySLVPKhoa, string userid, string mabv, string notmabv, bool LaySLTheoBenhNhan, bool LayTheoTT2348, string bc_mabn, string bc_quyenso)
        {
            StringBuilder builder = new StringBuilder();
            string user = this._lib.user;
            if (bc_mabn != "")
            {
                bc_mabn = "'" + bc_mabn.Trim(new char[] { ',' }).Replace(",", "','") + "'";
            }
            string str2 = user + "d" + mmyy;
            string str3 = user + mmyy;
            string str4 = "select e.id from " + str3 + ".v_ttrvll e  left join " + str3 + ".v_hoantra f on e.quyenso=f.quyenso and e.sobienlai=f.sobienlai inner join " + str3 + ".v_ttrvds g on e.id=g.id where  g.mabn=f.mabn and e.ngay between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')";
            string str5 = " and nhombh.idnhombhytmedisoft in(" + nhomvp_bhyt_medi + ")";
            if (LaySLVPKhoa)
            {
                builder.Append(" select 0 as stt,bh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,xv.loaiba as loaiba, c.manhom, d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, ");
                builder.Append(" c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx, ");
                builder.Append(" c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,  sum(a.soluong) as soluong, round(a.dongia,2) as dongia, ");
                builder.Append(" sum(a.soluong*a.dongia) sotien , sum(a.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp ,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                builder.Append(" from " + str3 + ".v_thvpll ab inner join " + str3 + ".v_thvpct a on ab.id=a.id ");
                builder.Append(" inner join " + str3 + ".v_thvpbhyt a1 on a.id=a1.id  ");
                builder.Append(" inner join " + user + ".d_dmbd c on a.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id ");
                builder.Append(" inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma ");
                builder.Append(" left join " + user + ".tenvien c1 on c1.mabv=c.mabv  ");
                builder.Append(" left join " + user + ".d_dmnx c2 on c2.id=c.madv  ");
                builder.Append(" left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id ");
                builder.Append(" left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                builder.Append(" left join " + user + ".benhandt xv on ab.maql=xv.maql ");
                builder.Append(" left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on bh.maql=ab.maql ");
                builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and xv.loaiba in(" + loaibenhan + ") and a1.sothe is not null   ");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(str5);
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and ab.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" group by ab.ngay,xv.loaiba,bh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn" : ""));
                builder.Append(" union all ");
                builder.Append(" select 0 as stt,bh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,xv.loaiba as loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append("to_char(ab.ngay,'dd/mm/yyyy') ngayduyet,  sum(a.soluong) as soluong,  round(a.dongia,2) as dongia, sum(a.soluong*a.dongia)as sotien , sum(a.bhyttra)as bhyttra ,to_char(c.ma) as mavp1,a.makp,bv.tenkp,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn"));
                builder.Append("  from " + str3 + ".v_thvpll ab  inner join " + str3 + ".v_thvpct a on ab.id=a.id   inner join " + str3 + ".v_thvpbhyt b on a.id=b.id  inner join " + user + ".v_giavp c on a.mavp=c.id  inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma  left join " + user + ".v_nhombhyt bh on e.idnhombhyt=bh.id  left join " + user + ".v_nhombhyt_medisoft soft on bh.idnhombhytmedisoft=soft.id  left join " + user + ".btdkp_bv bv on a.makp=bv.makp  left join " + user + ".benhandt xv on ab.maql=xv.maql  left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on ab.maql=bh.maql ");
                builder.Append(" where to_date(to_char(ab.ngayra,'dd/mm/yyyy'),'dd/mm/yyyy')  between to_date('" + tungay + "','dd/mm/yyyy')  and to_date('" + denngay + "','dd/mm/yyyy')");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" and xv.loaiba in(" + loaibenhan + ") and b.sothe is not null ");
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(str5);
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and ab.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                builder.Append(" group by ab.ngay,xv.loaiba,c.ma ,bh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, a.dongia, c.kythuat ,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn" : ""));
            }
            else
            {
                builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,a.loaibn as loaiba, c.manhom, d.ten as tennhom, c.maloai,o.ten as tenloai, d.nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.tenhc, c.hamluong, c.dang as dvt, c.sodk, c.masobyt, c.tenbyt,0 as nhomcc, null as nhacc, c.manuoc, m.ten nuocsx,c.mahang, h.ten hangsx, c.kythuat,c.duongdung,d.stt sttnhom,c.stt31,c2.ten tencty,c.donvi,c.madv,c1.tenbv tenbvthau,c.mabv,d.thuocyhct,to_char(a.ngay,'dd/mm/yyyy') ngayduyet, sum(a1.soluong) as soluong,  round(a1.dongia,2) as dongia,  sum(a1.soluong*a1.dongia) sotien , sum(a1.bhyttra)as bhyttra,c.ma as mavp1,a.makp,bv.tenkp,bh.mabv " + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bh.malk,badt.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append(" from " + str3 + ".v_ttrvds ab  inner join " + str3 + ".v_ttrvll a on ab.id=a.id  inner join " + user + ".benhandt badt on badt.maql=ab.maql and badt.mabn=ab.mabn  left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on badt.maql=bh.maql  inner join " + str3 + ".v_ttrvct a1 on a.id=a1.id   inner join " + user + ".d_dmbd c on a1.mavp=c.id  inner join " + user + ".d_dmloai o on c.maloai=o.id  inner join " + user + ".d_dmnhom d on c.manhom= d.id  left join " + user + ".v_nhomvp e on d.nhomvp=e.ma  left join " + user + ".tenvien c1 on c1.mabv=c.mabv   left join " + user + ".d_dmnx c2 on c2.id=c.madv   left join " + user + ".d_dmnuoc m on c.manuoc=m.id  left join " + user + ".d_dmhang h on c.mahang=h.id  left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id  left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id  left join " + user + ".btdkp_bv bv on a.makp=bv.makp " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b21.ma=badt.mabs") : ""));
                builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                builder.Append(str5);
                builder.Append(" and a.loaibn in(" + loaibenhan + ") and ab.id not in (" + str4 + ") ");
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a1.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" group by a.ngay,a.loaibn,nhombh.idnhombhytmedisoft,soft.ten,c.ma ,a.makp,bv.tenkp,c.manhom, d.ten, c.maloai,  o.ten, d.nhomvp, e.ten, c.id, ");
                builder.Append(" c.ten, c.tenhc, c.hamluong, c.dang , c.sodk, c.masobyt, c.tenbyt , a1.dongia,c.manuoc, m.ten, c.mahang, h.ten, c.kythuat ,c.duongdung,d.stt,c.stt31,c2.ten,c.donvi,c.madv,c1.tenbv,c.mabv,d.thuocyhct,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(a1.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,ab.maql,bh.malk,badt.mavaovien" : ""));
                builder.Append(" union all ");
                builder.Append(" select 0 as stt,nhombh.idnhombhytmedisoft as nhombhyt,soft.ten as tennhombhyt,a.loaibn as loaiba, d.id_nhom as manhom, e.ten as tennhom, c.id_loai as maloai, d.ten as tenloai, ");
                builder.Append(" d.id_nhom as nhomvp, e.ten tennhomvp, c.id as mavp,  c.ten, c.ten as tenhc, null as hamluong, c.dvt, ");
                builder.Append(" null as sodk, c.masobyt, c.tenbyt, 0 as nhomcc, null as nhacc, 0 as manuoc, null as nuocsx, 0 as mahang, null as hangsx, c.kythuat,");
                builder.Append("null as duongdung,null as sttnhom,null as stt31,null as tencty,null as donvi,null as madv,null as tenbvthau,null as mabv,null as thuocyhct,");
                builder.Append("to_char(a.ngay,'dd/mm/yyyy') ngayduyet,   sum(b.soluong) as soluong,round(b.dongia,2) as dongia,  sum(b.soluong*b.dongia)as sotien , sum(b.bhyttra)as bhyttra ,to_char(c.ma) as mavp1,a.makp,bv.tenkp,bh.mabv" + (LaySLTheoBenhNhan ? ",ab.mabn as mabn" : ",'' as mabn") + (LayTheoTT2348 ? ",b21.viettat as kihieubs,ab.maicd as maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi') as ngayylenh,bv.viettat as kihieukp,ab.maql as maql,bh.malk,badt.mavaovien" : ",'' as kihieubs,'' as maicd, '' as ngayylenh,'' as kihieukp,0 as maql,0 as malk,0 as mavaovien"));
                builder.Append("  from " + str3 + ".v_ttrvds ab inner join " + str3 + ".v_ttrvll a on ab.id=a.id  ");
                builder.Append("  inner join " + str3 + ".v_ttrvct b on a.id=b.id inner join " + user + ".v_giavp c on b.mavp=c.id ");
                builder.Append(" inner join " + user + ".v_loaivp d on c.id_loai=d.id  inner join " + user + ".v_nhomvp e on d.id_nhom=e.ma ");
                builder.Append(" left join " + user + ".v_nhombhyt nhombh on e.idnhombhyt=nhombh.id ");
                builder.Append(" left join " + user + ".v_nhombhyt_medisoft soft on nhombh.idnhombhytmedisoft=soft.id ");
                builder.Append(" left join " + user + ".btdkp_bv bv on a.makp=bv.makp ");
                builder.Append(" inner join " + user + ".benhandt badt on badt.maql=ab.maql  and badt.mabn=ab.mabn");
                builder.Append(" left join (" + this.f_get_sql_theBHYT(1, mmyy) + ") bh on badt.maql=bh.maql " + (LayTheoTT2348 ? (" left join " + user + ".dmbs b21 on b21.ma=badt.mabs") : ""));
                builder.Append(" where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay + "','dd/mm/yyyy') and to_date('" + denngay + "','dd/mm/yyyy')");
                builder.Append(" and a.loaibn in (" + loaibenhan + ") and ab.id not in (" + str4 + ") ");
                builder.Append(str5);
                if (bc_mabn != "")
                {
                    builder.Append(" and ab.mabn in (" + bc_mabn + ")");
                }
                if (makp.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.makp in (" + makp.Trim(new char[] { ',' }) + ")");
                }
                if (madoituong.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and b.madoituong in (" + madoituong.Trim(new char[] { ',' }) + ")");
                }
                if (userid.Trim(new char[] { ',' }) != "")
                {
                    builder.Append(" and a.userid in (" + userid.Trim(new char[] { ',' }) + ")");
                }
                if (bc_quyenso != "")
                {
                    builder.Append(" and a.quyenso in (" + bc_quyenso + ")");
                }
                if (mabv != "")
                {
                    builder.Append(" and bh.mabv in('" + mabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                if (notmabv != "")
                {
                    builder.Append(" and bh.mabv not in('" + notmabv.Trim(new char[] { ',' }).Replace(",", "','") + "')");
                }
                builder.Append(" group by a.ngay,a.loaibn,c.ma ,nhombh.idnhombhytmedisoft,soft.ten,a.makp,bv.tenkp, d.id_nhom, e.ten, c.id_loai, d.ten, d.id_nhom, e.ten, c.id, c.masobyt, c.tenbyt,  c.ten, c.ten, c.dvt, b.dongia, c.kythuat,bh.mabv " + (LaySLTheoBenhNhan ? ",ab.mabn" : "") + (LayTheoTT2348 ? ",b21.viettat,ab.maicd,to_char(b.ngay,'dd/mm/yyyy hh24:mi'),bv.viettat,ab.maql,bh.malk,badt.mavaovien" : ""));
            }
            return builder.ToString();
        }

        public DataSet f_Noitru_tao_dataset(System.Data.DataTable dtnhomvp)
        {
            DataSet set = new DataSet();
            set = this._lib.get_data("select 0 as sobienlai, null as quyenso,null ngaythu, 0 stt,null as stt2,null sothe1,null sothe2,null sothe3,0 id,null mabn,null hoten,null ngaysinh,null phai,null sothe,null manoidk,null noidk,null chandoan,null maicd,null maicdkt,null chandoankt,null tngt,null as sophieu,null ngayvao,null ngayra,null songay from dual");
            set.Clear();
            foreach (DataRow row in dtnhomvp.Select("true", "stt"))
            {
                set.Tables[0].Columns.Add(new DataColumn("ST_" + row["id"].ToString().Trim().Trim(new char[] { '_' }), typeof(decimal)));
            }
            set.Tables[0].Columns.Add(new DataColumn("SOPHIEUTHANHTOANRAVIEN", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("ST_102", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("TONGCONG", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("BNTRA", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("BHYTTRA", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("CPDS", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("MAKP", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("KHOA", typeof(string)));
            set.Tables[0].Columns.Add(new DataColumn("TRAITUYEN", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("capcuu", typeof(decimal)));
            set.Tables[0].Columns.Add(new DataColumn("TYLEBHYT", typeof(decimal)));
            DataColumn column = new DataColumn();
            column.ColumnName = "NHOM_DT_BHYT";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "TEN_NHOMDT_BHYT";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "madk";
            column.DataType = System.Type.GetType("System.Decimal");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "Lydo";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "benhkhac";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "noidkkcb";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "namqt";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "thangqt";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "tungay";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "denngay";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "diachi";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "madoituong";
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "giatritu";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            column = new DataColumn();
            column.ColumnName = "giatriden";
            column.DataType = System.Type.GetType("System.String");
            set.Tables[0].Columns.Add(column);
            set.Tables[0].Columns.Add("malk", typeof(string));
            set.Tables[0].Columns.Add("loaiba", typeof(int));
            set.Tables[0].Columns.Add("nguoithu");
            set.Tables[0].Columns.Add("mavaovien", typeof(long));
            set.Tables[0].Columns.Add("ngayrv", typeof(string));
            set.Tables[0].Columns.Add("ngayvv", typeof(string));
            return set;
        }

        public void f_Noitru_xuatExcel_mau01_TT2348(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dset = new DataSet();
            dset.Tables.Add("Table");
            dset.Tables[0].Columns.Add("ma_lk", typeof(string));
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
            dset.Tables[0].Columns.Add("ten_benh", typeof(string));
            dset.Tables[0].Columns.Add("ma_lydo_vvien", typeof(int));
            dset.Tables[0].Columns.Add("ma_noi_chuyen", typeof(string));
            dset.Tables[0].Columns.Add("ma_tai_nan", typeof(int));
            dset.Tables[0].Columns.Add("ngay_vao", typeof(string));
            dset.Tables[0].Columns.Add("ngay_ra", typeof(string));
            dset.Tables[0].Columns.Add("so_ngay_dtri", typeof(int));
            dset.Tables[0].Columns.Add("ket_qua_dtri", typeof(int));
            dset.Tables[0].Columns.Add("tinh_trang_rv", typeof(int));
            dset.Tables[0].Columns.Add("muc_huong", typeof(decimal));
            dset.Tables[0].Columns.Add("t_tongchi", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bntt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_bhtt", typeof(decimal));
            dset.Tables[0].Columns.Add("t_nguonkhac", typeof(decimal));
            dset.Tables[0].Columns.Add("t_ngoaids", typeof(decimal));
            dset.Tables[0].Columns.Add("nam_qt", typeof(int));
            dset.Tables[0].Columns.Add("thang_qt", typeof(int));
            dset.Tables[0].Columns.Add("ma_loaikcb", typeof(int));
            dset.Tables[0].Columns.Add("ma_cskcb", typeof(string));
            dset.Tables[0].Columns.Add("ma_khuvuc", typeof(string));
            dset.Tables[0].Columns.Add("ma_PTTT_QT", typeof(string));
            decimal d = 0M;
            decimal num2 = this._lib.themoi15_sotien();
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                try
                {
                    DataRow row2 = dset.Tables[0].NewRow();
                    try
                    {
                        row2["ma_lk"] = this._lib.Mabv + row["malk"].ToString();
                    }
                    catch
                    {
                        row2["ma_lk"] = row["malk"].ToString();
                    }
                    row2["STT"] = d = decimal.op_Increment(d);
                    row2["ma_bn"] = row["mabn"].ToString();
                    row2["ho_ten"] = row["HOTEN"].ToString();
                    row2["ngay_sinh"] = row["ngaysinh"].ToString();
                    row2["gioi_tinh"] = (row["phai"].ToString() == "0") ? 1 : 2;
                    row2["dia_chi"] = row["diachi"].ToString();
                    try
                    {
                        row2["ma_the"] = row["sothe"].ToString().Substring(0, 15);
                    }
                    catch
                    {
                        row2["ma_the"] = row["sothe"].ToString();
                    }
                    row2["ma_dkbd"] = row["MANOIDK"].ToString();
                    try
                    {
                        row2["gt_the_tu"] = row["giatritu"].ToString().Substring(6, 4) + row["giatritu"].ToString().Substring(3, 2) + row["giatritu"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_tu"] = "";
                    }
                    try
                    {
                        row2["gt_the_den"] = row["giatriden"].ToString().Substring(6, 4) + row["giatriden"].ToString().Substring(3, 2) + row["giatriden"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["gt_the_den"] = "";
                    }
                    row2["ma_benh"] = row["MAICD"].ToString();
                    row2["ma_benhkhac"] = row["MAICDKT"].ToString();
                    row2["ten_benh"] = row["chandoan"].ToString();
                    if (row["lydo"].ToString() == "2")
                    {
                        row2["ma_lydo_vvien"] = 2;
                    }
                    else if (row["traituyen"].ToString() != "0")
                    {
                        row2["ma_lydo_vvien"] = 3;
                    }
                    else if (row["traituyen"].ToString() == "0")
                    {
                        row2["ma_lydo_vvien"] = 1;
                    }
                    row2["ma_noi_chuyen"] = "";
                    row2["ma_tai_nan"] = 0;
                    try
                    {
                        row2["ngay_vao"] = row["NGAYVAO"].ToString().Substring(6, 4) + row["NGAYVAO"].ToString().Substring(3, 2) + row["NGAYVAO"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_vao"] = "";
                    }
                    try
                    {
                        row2["ngay_ra"] = row["NGAYRA"].ToString().Substring(6, 4) + row["NGAYRA"].ToString().Substring(3, 2) + row["NGAYRA"].ToString().Substring(0, 2);
                    }
                    catch
                    {
                        row2["ngay_ra"] = "";
                    }
                    try
                    {
                        row2["so_ngay_dtri"] = Convert.ToInt16(row["SONGAY"].ToString());
                    }
                    catch
                    {
                        row2["so_ngay_dtri"] = 0;
                    }
                    row2["ket_qua_dtri"] = 0;
                    row2["tinh_trang_rv"] = 0;
                    try
                    {
                        row2["muc_huong"] = Convert.ToDecimal(row["tylebhyt"].ToString());
                    }
                    catch
                    {
                    }
                    row2["t_tongchi"] = row["tongcong"].ToString();
                    row2["t_bntt"] = row["bntra"].ToString();
                    row2["t_bhtt"] = row["bhyttra"].ToString();
                    row2["t_nguonkhac"] = 0;
                    row2["t_ngoaids"] = 0;
                    row2["nam_qt"] = denngay.Substring(6, 4);
                    row2["thang_qt"] = denngay.Substring(3, 2);
                    row2["ma_loaikcb"] = 3;
                    row2["ma_cskcb"] = this._lib.MABHXH;
                    row2["ma_khuvuc"] = "";
                    row2["ma_PTTT_QT"] = "";
                    dset.Tables[0].Rows.Add(row2);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, this._lib.Msg);
                }
            }
            dset.WriteXml("bang01_noi.xml", XmlWriteMode.WriteSchema);
            int num3 = 0;
            int num4 = 0;
            int num5 = 0;
            int i = 0;
            int num7 = 0;
            num3 = 3;
            num4 = 5;
            num5 = dset.Tables[0].Rows.Count + 5;
            i = dset.Tables[0].Columns.Count - 1;
            num7 = num5;
            this.tenfile = this._lib.Export_Excel(dset, "bccpkcb01_n");
            try
            {
                this._lib.check_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num3; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(num3 + 8) + num4.ToString(), this._lib.getIndex(i - 0x12) + num5.ToString()).NumberFormat = "#,##";
                this.osheet.get_Range(this._lib.getIndex(0) + "4", this._lib.getIndex(i) + num7.ToString()).Borders.LineStyle = XlBorderWeight.xlHairline;
                string[] strArray = new string[] { "ma_lk", "ma_bn", "ma_dkbd" };
                for (int k = 0; k < strArray.Length; k++)
                {
                    try
                    {
                        int ordinal = dset.Tables[0].Columns[strArray[k]].Ordinal;
                        this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + num4, this._lib.getIndex(ordinal) + num5);
                        this.orange.NumberFormat = "@";
                    }
                    catch
                    {
                    }
                }
                for (int m = 0; m < dset.Tables[0].Rows.Count; m++)
                {
                    for (int n = 0; n < strArray.Length; n++)
                    {
                        int num13 = dset.Tables[0].Columns[strArray[n]].Ordinal;
                        try
                        {
                            this.osheet.Cells[num4 + m, num13 + 1] = dset.Tables[0].Rows[m][num13].ToString();
                        }
                        catch
                        {
                        }
                    }
                }
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(i + 2) + num5.ToString());
                this.orange.Font.Name = "Arial";
                this.orange.Font.Size = 8;
                this.orange.EntireColumn.AutoFit();
                this.oxl.ActiveWindow.DisplayZeros = true;
                this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                this.osheet.PageSetup.LeftMargin = 20.0;
                this.osheet.PageSetup.RightMargin = 20.0;
                this.osheet.PageSetup.TopMargin = 30.0;
                this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + "1", this._lib.getIndex(3) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.osheet.Cells[1, 4] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NỘI TR\x00da ";
                this.osheet.Cells[2, 4] = (tungay == denngay) ? ("Ng\x00e0y : " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay);
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "1", this._lib.getIndex(i) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Size = 12;
                this.orange.Font.Bold = true;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception2)
            {
                MessageBox.Show("Kh\x00f4ng c\x00f3 số liệu\n\n" + exception2.Message, this._lib.Msg);
            }
        }

        public void f_Noitru_xuatExcel_mau26th(bool print, DataSet dsdulieu, string tungay, string denngay)
        {
            DataSet dstam = new DataSet();
            dstam.Tables.Add("tam");
            dstam.Tables[0].Columns.Add("stt");
            dstam.Tables[0].Columns.Add("hoten");
            dstam.Tables[0].Columns.Add("soluot");
            dstam.Tables[0].Columns.Add("songay");
            dstam.Tables[0].Columns.Add("st_1");
            dstam.Tables[0].Columns.Add("st_2");
            dstam.Tables[0].Columns.Add("st_3");
            dstam.Tables[0].Columns.Add("st_4");
            dstam.Tables[0].Columns.Add("st_5");
            dstam.Tables[0].Columns.Add("st_6");
            dstam.Tables[0].Columns.Add("st_7");
            dstam.Tables[0].Columns.Add("st_8");
            dstam.Tables[0].Columns.Add("st_11");
            dstam.Tables[0].Columns.Add("st_10");
            dstam.Tables[0].Columns.Add("tongcong");
            dstam.Tables[0].Columns.Add("bntra");
            dstam.Tables[0].Columns.Add("bhyttra");
            dstam.Tables[0].Columns.Add("chiphids");
            dstam.Tables[0].Columns.Add("madk");
            DataSet set2 = new DataSet();
            set2 = dsdulieu.Copy();
            try
            {
                set2.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            set2.Tables[0].Columns.Add("yyyymmdd");
            set2.Tables[0].Columns.Add("soluot");
            for (int i = 0; i < set2.Tables[0].Rows.Count; i++)
            {
                try
                {
                    set2.Tables[0].Rows[i]["madk"] = int.Parse(set2.Tables[0].Rows[i]["MADK"].ToString());
                }
                catch
                {
                }
                if (set2.Tables[0].Rows[i]["traituyen"].ToString() == "1")
                {
                    set2.Tables[0].Rows[i]["traituyen"] = 1;
                }
                else
                {
                    set2.Tables[0].Rows[i]["traituyen"] = 2;
                }
                try
                {
                    set2.Tables[0].Rows[i]["yyyymmdd"] = set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(6, 4) + set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(3, 2) + set2.Tables[0].Rows[i]["ngayra"].ToString().Substring(0, 2);
                }
                catch
                {
                    set2.Tables[0].Rows[i]["yyyymmdd"] = 0;
                }
                set2.Tables[0].Rows[i]["soluot"] = 1;
            }
            DataSet set3 = new DataSet();
            set3 = set2.Copy();
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "st_1", "st_2", "st_3", "st_4", "st_5", "st_6", "st_7", "st_8", "st_11", "st_10", "TONGCONG", "bntra", "bhyttra", "chiphids", "soluot", "songay" };
            string str = "";
            DataRow row = dstam.Tables[0].NewRow();
            DataRow row2 = dstam.Tables[0].NewRow();
            for (int j = 1; j <= 3; j++)
            {
                switch (j)
                {
                    case 1:
                        str = "ABỆNH nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str = "BBỆNH nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str = "CBỆNH nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row3 = dstam.Tables[0].NewRow();
                row3["stt"] = str.Substring(0, 1);
                row3["hoten"] = str.Substring(1);
                dstam.Tables[0].Rows.Add(row3);
                DataRow row4 = dstam.Tables[0].NewRow();
                for (int k = 1; k <= 2; k++)
                {
                    DataRow row5 = dstam.Tables[0].NewRow();
                    row5["stt"] = (k == 1) ? "I" : "II";
                    row5["hoten"] = (k == 1) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN";
                    foreach (DataRow row6 in set3.Tables[0].Select("madk=" + j.ToString() + " and traituyen=" + k.ToString(), "yyyymmdd"))
                    {
                        for (int m = 0; m < strArray2.Length; m++)
                        {
                            try
                            {
                                if (row5[strArray2[m]].ToString() == "")
                                {
                                    row5[strArray2[m]] = 0;
                                }
                                if (row[strArray2[m]].ToString() == "")
                                {
                                    row[strArray2[m]] = 0;
                                }
                                if (row4[strArray2[m]].ToString() == "")
                                {
                                    row4[strArray2[m]] = 0;
                                }
                                if (row6[strArray2[m]].ToString() == "")
                                {
                                    row6[strArray2[m]] = 0;
                                }
                                row5[strArray2[m]] = decimal.Parse(row5[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                                row[strArray2[m]] = decimal.Parse(row[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                                row4[strArray2[m]] = decimal.Parse(row4[strArray2[m]].ToString()) + decimal.Parse(row6[strArray2[m]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    if ((row5["tongcong"].ToString() != "") && (row5["tongcong"].ToString() != "0"))
                    {
                        dstam.Tables[0].Rows.Add(row5);
                    }
                }
                row4["stt"] = str.Substring(0, 1);
                row4["hoten"] = "cộng " + row4["stt"].ToString();
                if ((row4["tongcong"].ToString() != "") && (row4["tongcong"].ToString() != "0"))
                {
                    dstam.Tables[0].Rows.Add(row4);
                }
                else
                {
                    dstam.Tables[0].Rows.RemoveAt(dstam.Tables[0].Rows.Count - 1);
                }
            }
            row["hoten"] = "Tổng cộng A+B+C";
            dstam.Tables[0].Rows.Add(row);
            dstam.Tables[0].Columns.Remove("madk");
            dstam.Tables[0].Columns["st_1"].ColumnName = "X\x00e9t nghiệm";
            dstam.Tables[0].Columns["st_2"].ColumnName = "CĐHA TDCN";
            dstam.Tables[0].Columns["st_3"].ColumnName = "Thuốc dịch";
            dstam.Tables[0].Columns["st_4"].ColumnName = "M\x00e1u";
            dstam.Tables[0].Columns["st_5"].ColumnName = "Thủ thuật phẫu thuật";
            dstam.Tables[0].Columns["st_6"].ColumnName = "V?t tu y t?";
            dstam.Tables[0].Columns["st_7"].ColumnName = "DVKT cao";
            dstam.Tables[0].Columns["st_8"].ColumnName = "Thu?c K, CTG";
            dstam.Tables[0].Columns["st_11"].ColumnName = "Tiền giường";
            dstam.Tables[0].Columns["st_10"].ColumnName = "CP Vận chuyển";
            this.f_Noitru_xuatExcel_mau26th_run(print, dstam, tungay, denngay);
        }

        private void f_Noitru_xuatExcel_mau26th_run(bool print, DataSet dstam, string tungay, string denngay)
        {
            this._lib.check_process_Excel();
            try
            {
                DataRow row = dstam.Tables[0].NewRow();
                for (int i = 0; i < dstam.Tables[0].Columns.Count; i++)
                {
                    row[i] = i + 1;
                }
                dstam.Tables[0].Rows.InsertAt(row, 0);
                int num2 = 6;
                int num3 = 5;
                int count = dstam.Tables[0].Columns.Count;
                this.tenfile = this._lib.Export_Excel(dstam, "bccpkcb_26th");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num2; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).Font.Name = "Arial";
                this.osheet.get_Range(this._lib.getIndex(2) + num3, this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).NumberFormat = "#,##0";
                this.osheet.Cells[1, 1] = this._lib.Tenbv;
                this.osheet.Cells[1, count] = "Mẫu số: 26A-TH/BHYT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "1", this._lib.getIndex((count - 1) - 2) + "1");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Size = 8;
                this.orange.MergeCells = true;
                this.osheet.Cells[2, 1] = "B\x00c1O C\x00c1O CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT NỮI TR\x00da ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "2", this._lib.getIndex(count - 1) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 0x10;
                this.osheet.Cells[3, 1] = "Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "3", this._lib.getIndex(count - 1) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num2 + 2));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                int num6 = -1;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "TT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.ColumnWidth = 5;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Ch? ti\x00eau";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "S? lu?t";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Số ng\x00e0y di?u tr?";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num3 + 1, num6 + 1] = "CHI PH\x00cd PH\x00c1T SINH T?I CO S? KH\x00c1M CHỮA BỆNH".ToUpper();
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num3 + 1), this._lib.getIndex((count - 1) - 3) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 2] = "Người bệnh c\x00f9ng chi trả";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 2) + (num2 + 1), this._lib.getIndex((count - 1) - 2) + (num3 + 1));
                this.orange.ColumnWidth = 10;
                this.orange.MergeCells = true;
                this.osheet.Cells[num3 + 1, count - 1] = "Chi ph\x00ed đề nghị BHXH thanh to\x00e1n";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num3 + 1), this._lib.getIndex(count - 1) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 1] = "Số tiền";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num2 + 1), this._lib.getIndex((count - 1) - 1) + (num2 + 1));
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, count] = "Trong đ\x00f3 chi ph\x00ed ngo\x00e0i quỹ định suất";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, count - 3] = "Tổng cộng";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 4) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                int num7 = num2 + 2;
                for (int k = 1; k < dstam.Tables[0].Rows.Count; k++)
                {
                    num7++;
                    this.orange = this.osheet.get_Range("A" + num7.ToString(), this._lib.getIndex(count - 1) + num7.ToString());
                    if (((dstam.Tables[0].Rows[k]["stt"].ToString() == "A") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((dstam.Tables[0].Rows[k]["stt"].ToString() == "C") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                        if ((dstam.Tables[0].Rows[k][1].ToString().Substring(0, 1) != "C") && (dstam.Tables[0].Rows[k][1].ToString().Substring(0, 1) != "T"))
                        {
                            this.orange = this.osheet.get_Range("B" + num7.ToString(), this._lib.getIndex(count - 1) + num7.ToString());
                            this.orange.MergeCells = true;
                        }
                    }
                }
                this.oxl.ActiveWindow.DisplayZeros = false;
                try
                {
                    this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                    this.osheet.PageSetup.LeftMargin = 20.0;
                    this.osheet.PageSetup.RightMargin = 20.0;
                    this.osheet.PageSetup.TopMargin = 30.0;
                    this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                }
                catch
                {
                }
                decimal num9 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["tongcong"].ToString()), 0);
                string str = new numbertotext().doiraso(num9.ToString());
                this.osheet.Cells[((num2 + 1) + dstam.Tables[0].Rows.Count) + 1, 2] = "Số tiền d? ngh? thanh to\x00e1n (viết bằng chữ): " + str.Substring(0, 1).ToUpper() + str.Substring(1) + " đồng chẵn.";
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public void f_Noitru_xuatExcel_mau80(bool print, DataSet dsxml, string tungay, string denngay, string fontchu)
        {
            DataSet set = new DataSet();
            set = dsxml.Copy();
            set.Tables[0].Columns.Add("ngaysinh_nam");
            set.Tables[0].Columns.Add("ngaysinh_nu");
            try
            {
                set.Tables[0].Columns.Add("madk");
            }
            catch
            {
            }
            set.Tables[0].Columns.Add("yyyymmdd");
            try
            {
                set.Tables[0].Columns.Add("chiphids");
            }
            catch
            {
            }
            for (int i = 0; i < set.Tables[0].Rows.Count; i++)
            {
                if (set.Tables[0].Rows[i]["phai"].ToString() == "0")
                {
                    set.Tables[0].Rows[i]["phai"] = "Nam";
                }
                else
                {
                    set.Tables[0].Rows[i]["phai"] = "NỮ";
                }
                if (set.Tables[0].Rows[i]["phai"].ToString().ToUpper() == "NAM")
                {
                    set.Tables[0].Rows[i]["ngaysinh_nam"] = set.Tables[0].Rows[i]["ngaysinh"].ToString();
                }
                else
                {
                    set.Tables[0].Rows[i]["ngaysinh_nu"] = set.Tables[0].Rows[i]["ngaysinh"].ToString();
                }
                try
                {
                    set.Tables[0].Rows[i]["st_11"] = Convert.ToDecimal(set.Tables[0].Rows[i]["st_11"].ToString()) + Convert.ToDecimal(set.Tables[0].Rows[i]["st_9"].ToString());
                    set.Tables[0].Rows[i]["st_9"] = 0;
                }
                catch
                {
                }
                try
                {
                    set.Tables[0].Rows[i]["chiphids"] = Convert.ToDecimal(set.Tables[0].Rows[i]["cpds"].ToString());
                }
                catch
                {
                    set.Tables[0].Rows[i]["chiphids"] = 0;
                }
                if (set.Tables[0].Rows[i]["traituyen"].ToString() == "0")
                {
                    set.Tables[0].Rows[i]["traituyen"] = 1;
                }
                else
                {
                    set.Tables[0].Rows[i]["traituyen"] = 2;
                }
                set.Tables[0].Rows[i]["yyyymmdd"] = set.Tables[0].Rows[i]["ngayra"].ToString().Substring(6, 4) + set.Tables[0].Rows[i]["ngayra"].ToString().Substring(3, 2) + set.Tables[0].Rows[i]["ngayra"].ToString().Substring(0, 2);
            }
            DataSet set2 = set.Copy();
            string[] strArray = new string[] { "", "I", "II", "III", "IV", "V", "VI" };
            string[] strArray2 = new string[] { "songay", "st_1", "st_2", "st_3", "st_4", "st_5", "st_6", "st_7", "st_14", "st_10", "st_11", "st_8", "tongcong", "bntra", "bhyttra", "chiphids" };
            DataSet dstam = new DataSet();
            dstam.Tables.Add("tam");
            dstam.Tables[0].Columns.Add("stt");
            dstam.Tables[0].Columns.Add("hoten");
            dstam.Tables[0].Columns.Add("ngaysinh_nam");
            dstam.Tables[0].Columns.Add("ngaysinh_nu");
            dstam.Tables[0].Columns.Add("sothe");
            dstam.Tables[0].Columns.Add("manoidk");
            dstam.Tables[0].Columns.Add("maicd");
            dstam.Tables[0].Columns.Add("ngayvao");
            dstam.Tables[0].Columns.Add("ngayra");
            dstam.Tables[0].Columns.Add("songay");
            dstam.Tables[0].Columns.Add("tongcong");
            dstam.Tables[0].Columns.Add("st_1");
            dstam.Tables[0].Columns.Add("st_2");
            dstam.Tables[0].Columns.Add("st_3");
            dstam.Tables[0].Columns.Add("st_4");
            dstam.Tables[0].Columns.Add("st_5");
            dstam.Tables[0].Columns.Add("st_6");
            dstam.Tables[0].Columns.Add("st_14");
            dstam.Tables[0].Columns.Add("st_7");
            dstam.Tables[0].Columns.Add("st_8");
            dstam.Tables[0].Columns.Add("st_11");
            dstam.Tables[0].Columns.Add("st_10");
            dstam.Tables[0].Columns.Add("bntra");
            dstam.Tables[0].Columns.Add("bhyttra");
            dstam.Tables[0].Columns.Add("chiphids");
            dstam.Tables[0].Columns.Add("madk");
            dstam.Tables[0].Columns.Add("quyenso");
            dstam.Tables[0].Columns.Add("sobienlai", typeof(decimal));
            dstam.Tables[0].Columns.Add("khoa");
            dstam.Tables[0].Columns.Add("giatritu");
            dstam.Tables[0].Columns.Add("giatriden");
            dstam.Tables[0].Columns.Add("nguoithu");
            dstam.Tables[0].Columns.Add("ngaythu");
            string str = "";
            DataRow row = dstam.Tables[0].NewRow();
            DataRow row2 = dstam.Tables[0].NewRow();
            int num2 = 0;
            for (int j = 1; j <= 3; j++)
            {
                switch (j)
                {
                    case 1:
                        str = "ABỆNH nh\x00e2n nội tỉnh KCB ban đầu".ToUpper();
                        break;

                    case 2:
                        str = "BBệnh nh\x00e2n nội tỉnh đến".ToUpper();
                        break;

                    case 3:
                        str = "CBệnh nh\x00e2n ngoại tỉnh đến".ToUpper();
                        break;
                }
                DataRow row3 = dstam.Tables[0].NewRow();
                row3["stt"] = str.Substring(0, 1);
                row3["hoten"] = str.Substring(1);
                dstam.Tables[0].Rows.Add(row3);
                DataRow row4 = dstam.Tables[0].NewRow();
                for (int k = 1; k <= 2; k++)
                {
                    DataRow row5 = dstam.Tables[0].NewRow();
                    row5["stt"] = row3["stt"].ToString() + ((k == 1) ? "I" : "II");
                    row5["hoten"] = (k == 1) ? "BỆNH NH\x00c2N KCB Đ\x00daNG TUYẾN (C\x00d3 GIẤY CHUYỂN VIỆN HOẶC TH CẤP CỨU)" : "BỆNH NH\x00c2N KCB TR\x00c1I TUYẾN (KH\x00d4NG C\x00d3 GIẤY CHUYỂN VIỆN HOẶC KH\x00d4NG PHẢI TRƯỜNG HỢP CẤP CỨU)";
                    dstam.Tables[0].Rows.Add(row5);
                    DataRow row6 = dstam.Tables[0].NewRow();
                    foreach (DataRow row7 in set2.Tables[0].Select("madk=" + j.ToString() + " and traituyen=" + k.ToString(), "yyyymmdd"))
                    {
                        DataRow row8 = dstam.Tables[0].NewRow();
                        for (int m = 0; m < dstam.Tables[0].Columns.Count; m++)
                        {
                            try
                            {
                                row8[m] = row7[dstam.Tables[0].Columns[m].ColumnName].ToString();
                            }
                            catch
                            {
                            }
                        }
                        row8["stt"] = ++num2;
                        dstam.Tables[0].Rows.Add(row8);
                        for (int n = 0; n < strArray2.Length; n++)
                        {
                            try
                            {
                                if (row6[strArray2[n]].ToString() == "")
                                {
                                    row6[strArray2[n]] = 0;
                                }
                                if (row[strArray2[n]].ToString() == "")
                                {
                                    row[strArray2[n]] = 0;
                                }
                                if (row4[strArray2[n]].ToString() == "")
                                {
                                    row4[strArray2[n]] = 0;
                                }
                                if (row7[strArray2[n]].ToString() == "")
                                {
                                    row7[strArray2[n]] = 0;
                                }
                                row6[strArray2[n]] = decimal.Parse(row6[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                                row[strArray2[n]] = decimal.Parse(row[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                                row4[strArray2[n]] = decimal.Parse(row4[strArray2[n]].ToString()) + decimal.Parse(row7[strArray2[n]].ToString());
                            }
                            catch
                            {
                            }
                        }
                    }
                    row6["stt"] = (k == 1) ? "I" : "II";
                    row6["hoten"] = "TỔNG " + ((k == 1) ? "Đ\x00daNG TUYẾN" : "TR\x00c1I TUYẾN");
                    if ((row6["tongcong"].ToString() != "") && (row6["tongcong"].ToString() != "0"))
                    {
                        dstam.Tables[0].Rows.Add(row6);
                    }
                    else
                    {
                        dstam.Tables[0].Rows.Remove(row5);
                    }
                }
                row4["stt"] = str.Substring(0, 1);
                row4["hoten"] = "cộng " + row4["stt"].ToString();
                if ((row4["tongcong"].ToString() != "") && (row4["tongcong"].ToString() != "0"))
                {
                    dstam.Tables[0].Rows.Add(row4);
                }
                else
                {
                    dstam.Tables[0].Rows.RemoveAt(dstam.Tables[0].Rows.Count - 1);
                }
            }
            row["hoten"] = "Tổng cộng A+B+C";
            dstam.Tables[0].Rows.Add(row);
            dstam.Tables[0].Columns.Remove("madk");
            dstam.Tables[0].Columns["ngaysinh_nam"].ColumnName = "Nam";
            dstam.Tables[0].Columns["ngaysinh_nu"].ColumnName = "NỮ";
            dstam.Tables[0].Columns["st_1"].ColumnName = "X\x00e9t nghiệm";
            dstam.Tables[0].Columns["st_2"].ColumnName = "CĐHA TDCN";
            dstam.Tables[0].Columns["st_3"].ColumnName = "Thuốc, dịch";
            dstam.Tables[0].Columns["st_4"].ColumnName = "M\x00e1u";
            dstam.Tables[0].Columns["st_5"].ColumnName = "Thủ thuật phẫu thuật";
            dstam.Tables[0].Columns["st_6"].ColumnName = "Vật tư y tế ti\x00eau hao";
            dstam.Tables[0].Columns["st_14"].ColumnName = "Vật tư y tế thay thế";
            dstam.Tables[0].Columns["st_7"].ColumnName = "DVKT cao";
            dstam.Tables[0].Columns["st_11"].ColumnName = "Tiền giường";
            dstam.Tables[0].Columns["st_10"].ColumnName = "Vận chuyển";
            try
            {
                dstam.Tables[0].Columns["st_8"].ColumnName = "Thuốc K,thải gh\x00e9p";
            }
            catch
            {
            }
            dstam.Tables[0].AcceptChanges();
            this.f_Noitru_xuatExcel_maumoi_80_run2(print, dstam, tungay, denngay, fontchu);
        }

        public void f_Noitru_xuatExcel_maumoi_41(bool print, DataSet dsdulieu, string tungay, string denngay, string fontchu)
        {
            DataSet set = this.f_Noitru_excel_mau41_getdata(dsdulieu, int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)));
            this.f_Noitru_exp_excel_mau41_run(print, set, tungay, denngay, fontchu);
        }

        private void f_Noitru_xuatExcel_maumoi_80_run(bool print, DataSet dstam, string tungay, string denngay, string fontchu)
        {
            this._lib.check_process_Excel();
            try
            {
                DataRow row = dstam.Tables[0].NewRow();
                for (int i = 0; i < dstam.Tables[0].Columns.Count; i++)
                {
                    if (i <= dstam.Tables[0].Columns["songay"].Ordinal)
                    {
                        row[i] = Convert.ToChar((int) (i + 0x41));
                    }
                    else
                    {
                        row[i] = i - dstam.Tables[0].Columns["songay"].Ordinal;
                    }
                }
                dstam.Tables[0].Rows.InsertAt(row, 0);
                int num2 = 9;
                int num3 = 7;
                int count = dstam.Tables[0].Columns.Count;
                this.tenfile = this._lib.Export_Excel(dstam, "bccpkcb");
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(this.tenfile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                for (int j = 0; j < num2; j++)
                {
                    this.osheet.get_Range(this._lib.getIndex(j) + "1", this._lib.getIndex(j) + "1").EntireRow.Insert(Missing.Value);
                }
                this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 1)).Borders.LineStyle = XlBorderWeight.xlHairline;
                this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).Font.Name = "Arial";
                this.osheet.get_Range(this._lib.getIndex(dstam.Tables[0].Columns["tongcong"].Ordinal) + num3, this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 10)).NumberFormat = "#,##0";
                this.osheet.Cells[1, 1] = this._lib.Syte;
                this.osheet.Cells[2, 1] = this._lib.Tenbv;
                this.osheet.Cells[1, count] = "Mẫu số: C80a-HD\n(Ban h\x00e0nh theo Th\x00f4ng tư số 178/TT \nng\x00e0y 23/10/2012 của Bộ T\x00e0i Ch\x00ednh)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "1", this._lib.getIndex((count - 1) - 4) + "2");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.Font.Size = 8;
                this.orange.MergeCells = true;
                this.osheet.Cells[3, 3] = "DANH S\x00c1CH NGƯỜI BỆNH BẢO HIỂM Y TẾ KH\x00c1M CHỮA BỆNH NỮI TR\x00da ĐỀ NGHỊ THANH TO\x00c1N";
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "3", this._lib.getIndex((count - 1) - 2) + "3");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 0x10;
                this.osheet.Cells[5, 3] = "TỪ NG\x00c0Y " + tungay + " ĐẾN " + denngay;
                this.orange = this.osheet.get_Range(this._lib.getIndex(2) + "5", this._lib.getIndex((count - 1) - 2) + "5");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.Font.Bold = true;
                this.orange.MergeCells = true;
                this.orange.Font.Size = 12;
                this.osheet.Cells[6, count] = "ĐVT: Đồng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + "6", this._lib.getIndex((count - 1) - 2) + "6");
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                this.orange.MergeCells = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 2), this._lib.getIndex(count - 1) + (num2 + 2));
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
                this.orange.WrapText = true;
                this.orange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num3 + 1), this._lib.getIndex(count - 1) + (num3 + 1));
                this.orange.RowHeight = 0x2d;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + num2, this._lib.getIndex(count - 1) + num2);
                this.orange.RowHeight = 30;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + (num2 + 1), this._lib.getIndex(count - 1) + (num2 + 1));
                this.orange.RowHeight = 0x37;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + 1, this._lib.getIndex(count - 1) + 1);
                this.orange.RowHeight = 20;
                int num6 = -1;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "STT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.ColumnWidth = 5;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Họ v\x00e0 t\x00ean".ToUpper();
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num3 + 1, num6 + 1] = "Năm sinh";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num3 + 1), this._lib.getIndex(num6 + 1) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Nam";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + num2);
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "NỮ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + num2);
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "M\x00e3 thẻ BHYT";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "M\x00e3 ĐK BĐ";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + ((num2 + 1) + 1), this._lib.getIndex(num6) + (((num2 + 1) + 1) + dstam.Tables[0].Rows.Count));
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "M\x00e3 bệnh";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Ng\x00e0y v\x00e0o";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Ng\x00e0y ra";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Số ng\x00e0y di?u tr?";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + (num3 + 1));
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num3 + 1, num6 + 1] = "TỔNG chi phi kh\x00e1m CHỮA BỆNH bhyt".ToUpper();
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num3 + 1), this._lib.getIndex((count - 1) - 3) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, num6 + 1] = "Tổng cộng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + (num2 + 1), this._lib.getIndex(num6) + num2);
                this.orange.MergeCells = true;
                num6++;
                this.osheet.Cells[num2, num6 + 1] = "Trong đ\x00f3";
                this.orange = this.osheet.get_Range(this._lib.getIndex(num6) + num2, this._lib.getIndex((count - 1) - 3) + num2);
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 2] = "Người bệnh c\x00f9ng chi trả";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 2) + (num2 + 1), this._lib.getIndex((count - 1) - 2) + (num3 + 1));
                this.orange.ColumnWidth = 10;
                this.orange.MergeCells = true;
                this.osheet.Cells[num3 + 1, count - 1] = "Chi ph\x00ed đề nghị cơ quan BHYT thanh to\x00e1n";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num3 + 1), this._lib.getIndex(count - 1) + (num3 + 1));
                this.orange.MergeCells = true;
                this.osheet.Cells[num2 + 1, count - 1] = "Số tiền";
                this.orange = this.osheet.get_Range(this._lib.getIndex((count - 1) - 1) + (num2 + 1), this._lib.getIndex((count - 1) - 1) + num2);
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                this.osheet.Cells[num2 + 1, count] = "Trong đ\x00f3 chi ph\x00ed ngo\x00e0i quỹ định suất";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + (num2 + 1), this._lib.getIndex(count - 1) + num2);
                this.orange.MergeCells = true;
                this.orange.ColumnWidth = 10;
                int num7 = num2 + 2;
                for (int k = 1; k < dstam.Tables[0].Rows.Count; k++)
                {
                    num7++;
                    this.orange = this.osheet.get_Range("A" + num7.ToString(), this._lib.getIndex(count - 1) + num7.ToString());
                    if (((dstam.Tables[0].Rows[k]["stt"].ToString() == "A") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "B")) || ((dstam.Tables[0].Rows[k]["stt"].ToString() == "C") || (dstam.Tables[0].Rows[k]["stt"].ToString() == "")))
                    {
                        this.orange.Font.ColorIndex = 5;
                        this.orange.Font.Bold = true;
                    }
                    else if (((dstam.Tables[0].Rows[k]["hoten"].ToString().IndexOf("Đ\x00daNG TUYẾN") > -1) || (dstam.Tables[0].Rows[k]["hoten"].ToString().IndexOf("TR\x00c1I TUYẾN") > -1)) || ((dstam.Tables[0].Rows[k]["hoten"].ToString().IndexOf("Đ\x00daNG TUYẾN") > -1) || (dstam.Tables[0].Rows[k]["hoten"].ToString().IndexOf("TR\x00c1I TUYẾN") > -1)))
                    {
                        this.orange.Font.ColorIndex = 10;
                        this.orange.Font.Bold = true;
                    }
                    if (dstam.Tables[0].Rows[k]["tongcong"].ToString() == "")
                    {
                        this.orange = this.osheet.get_Range("B" + num7, this._lib.getIndex(count - 1) + num7);
                        this.orange.MergeCells = true;
                    }
                    else if ((dstam.Tables[0].Rows[k]["stt"].ToString() == "") || !char.IsDigit(dstam.Tables[0].Rows[k]["stt"].ToString(), 0))
                    {
                        this.orange = this.osheet.get_Range("B" + num7, this._lib.getIndex(dstam.Tables[0].Columns["ngayra"].Ordinal) + num7);
                        this.orange.MergeCells = true;
                    }
                }
                this.orange.MergeCells = true;
                this.oxl.ActiveWindow.DisplayZeros = false;
                try
                {
                    this.osheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    this.osheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
                    this.osheet.PageSetup.LeftMargin = 20.0;
                    this.osheet.PageSetup.RightMargin = 20.0;
                    this.osheet.PageSetup.TopMargin = 30.0;
                    this.osheet.PageSetup.CenterFooter = "Trang : &P/&N";
                }
                catch
                {
                }
                decimal num9 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["bhyttra"].ToString()), 0);
                string str = new numbertotext().doiraso(num9.ToString());
                this.osheet.Cells[((num2 + 1) + dstam.Tables[0].Rows.Count) + 1, 2] = "Số tiền d? ngh? thanh to\x00e1n (viết bằng chữ): " + str.Substring(0, 1).ToUpper() + str.Substring(1) + " đồng chẵn.";
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 2] = "Người lập";
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 2] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(1) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 6] = "Trưởng ph\x00f2ng KHTH";
                this.orange = this.osheet.get_Range(this._lib.getIndex(5) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(7) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 6] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(5) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(7) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, 15] = "Kế to\x00e1n trưởng";
                this.orange = this.osheet.get_Range(this._lib.getIndex(14) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex(15) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, 15] = "(K\x00fd, họ t\x00ean)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(14) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex(15) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 3, count] = "Ng\x00e0y " + DateTime.Now.ToString("dd") + " th\x00e1ng " + DateTime.Now.ToString("MM") + " nam " + DateTime.Now.ToString("yyyy");
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 3), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 3));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 4, count] = "Thủ trưởng đơn vị";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 4), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 4));
                this.orange.MergeCells = true;
                this.orange.Font.Bold = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.osheet.Cells[(num2 + dstam.Tables[0].Rows.Count) + 5, count] = "(K\x00fd, họ t\x00ean, đ\x00f3ng dấu)";
                this.orange = this.osheet.get_Range(this._lib.getIndex(count - 1) + ((num2 + dstam.Tables[0].Rows.Count) + 5), this._lib.getIndex((count - 1) - 3) + ((num2 + dstam.Tables[0].Rows.Count) + 5));
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                this.orange = this.osheet.get_Range(this._lib.getIndex(0) + "1", this._lib.getIndex(dstam.Tables[0].Columns.Count + 5) + (dstam.Tables[0].Rows.Count + 50));
                this.orange.Font.Name = fontchu;
                if (print)
                {
                    this.osheet.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
                else
                {
                    this.oxl.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void f_Noitru_xuatExcel_maumoi_80_run2(bool print, DataSet dstam, string tungay, string denngay, string fontchu)
        {
            string path = Environment.CurrentDirectory + "//Excel//excelmau80hd.xls";
            string str2 = fontchu;
            StringBuilder builder = new StringBuilder();
            StreamWriter writer = new StreamWriter(path, false, Encoding.Unicode);
            builder.Append("<table>");
            builder.Append("<tr>");
            builder.Append("<td colspan=3 style=\"font-family:" + str2 + ";align=left\">" + this._lib.Syte + "</td>");
            for (int i = 3; i < (dstam.Tables[0].Columns.Count - 1); i++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<td colspan=1 align=right style=\"font-family:" + str2 + ";font-size:8pt\">Mẫu số: C80a-HD(Ban h\x00e0nh theo Th\x00f4ng tư số 178/TT ng\x00e0y 23/10/2012 của Bộ T\x00e0i Ch\x00ednh)</td>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td colspan=3 style=\"font-family:" + str2 + ";align=left\">" + this._lib.Tenbv + "</td>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<th colspan=", dstam.Tables[0].Columns.Count, " align=centre style=\"font-family: ", str2, "; font-size: 16pt\">DANH S\x00c1CH NGƯỜI BỆNH BẢO HIỂM Y TẾ KH\x00c1M CHỮA BỆNH NỮI TR\x00da ĐỀ NGHỊ THANH TO\x00c1N</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<th colspan=", dstam.Tables[0].Columns.Count, " align=centre style=\"font-family: ", str2, "; font-size: 12pt\">", (tungay == denngay) ? ("Ng\x00e0y " + tungay) : ("Từ ng\x00e0y " + tungay + " đến ng\x00e0y " + denngay), "</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count, " align=right style=\"font-family: ", str2, "; font-size: 10pt\">ĐVT: Đồng</td>" }));
            builder.Append("</tr>");
            builder.Append("<tr></tr>");
            builder.Append("</table>");
            builder.Append("<table border=1>");
            builder.Append("<tr>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">STT</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Họ v\x00e0 t\x00ean</th>");
            builder.Append("<th rowspan=1 colspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Năm sinh</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 thẻ BHYT</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 ĐK BĐ</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">M\x00e3 bệnh</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Ng\x00e0y v\x00e0o</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Ng\x00e0y ra</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Số ng\x00e0y</th>");
            builder.Append("<th rowspan=1 colspan=12 align=centre style=\"font-family: " + str2 + "\">TỔNG CHI PH\x00cd KH\x00c1M CHỮA BỆNH BHYT</th>");
            builder.Append("<th rowspan=3 align=centre style=\"font-family: " + str2 + ";height:30px\">Người bệnh \nc\x00f9ng chi trả</th>");
            builder.Append("<th rowspan=1 colspan=2 align=centre style=\"font-family: " + str2 + ";height:30pt\">Chi ph\x00ed đề nghị\n cơ quan BHYT \nthanh to\x00e1n</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Quyển sổ</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Số bi\x00ean lai</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Khoa ph\x00f2ng</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Ng\x00e0y bắt đầu thẻ</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Hạn thẻ</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Người thu</th>");
            builder.Append("<th rowspan=3 height=30 align=centre style=\"font-family: " + str2 + "\">Ng\x00e0y thu</th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Nam</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">NỮ</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Tổng cộng</th>");
            builder.Append("<th rowspan=1 colspan=11 height=30 align=centre style=\"font-family: " + str2 + "\">Trong đ\x00f3</th>");
            builder.Append("<th rowspan=2 height=30 align=centre style=\"font-family: " + str2 + "\">Số tiền</th>");
            builder.Append("<th rowspan=2 align=centre style=\"font-family: " + str2 + ";height:30px; width:10px\">Trong đ\x00f3 \n\rchi ph\x00ed ngo\x00e0i quỹ \n\rđịnh suất</th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">X\x00e9t nghiệm</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">CĐHA TDCN</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thuốc dịch</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">M\x00e1u</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thủ thuật phẫu thuật</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vật tư y tế ti\x00eau hao</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vật tư thay thế</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">DVKT cao</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Thuốc k, thải gh\x00e9p</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Tiền giường</th>");
            builder.Append("<th align=centre style=\"font-family: " + str2 + ";width:15px\">Vận chuyển</th>");
            builder.Append("</tr>");
            builder.Append("<tr >");
            for (int j = 0; j < dstam.Tables[0].Columns.Count; j++)
            {
                if (j <= dstam.Tables[0].Columns["songay"].Ordinal)
                {
                    builder.Append(string.Concat(new object[] { "<th style=\"font-family: ", str2, ";font-size: 11pt\">", Convert.ToChar((int) (j + 0x41)), "</th>" }));
                }
                else
                {
                    builder.Append(string.Concat(new object[] { "<th style=\"font-family: ", str2, ";font-size: 11pt\">", j - dstam.Tables[0].Columns["songay"].Ordinal, "</th>" }));
                }
            }
            builder.Append("</tr>\n");
            builder.Append("</table>");
            writer.Write(builder);
            builder = new StringBuilder();
            builder.Append("<table border=1 style=\"font-family:" + str2 + "\">");
            decimal num3 = 0M;
            decimal num4 = 0M;
            for (int k = 0; k < dstam.Tables[0].Rows.Count; k++)
            {
                num3 = 0M;
                num4 = 0M;
                try
                {
                    num3 = decimal.Parse(dstam.Tables[0].Rows[k]["tongcong"].ToString());
                }
                catch
                {
                }
                try
                {
                    num4 = decimal.Parse(dstam.Tables[0].Rows[k]["stt"].ToString());
                }
                catch
                {
                }
                builder.Append("<tr>");
                for (int num6 = 0; num6 < dstam.Tables[0].Columns.Count; num6++)
                {
                    if ((num4 == 0M) && (num3 == 0M))
                    {
                        string str3 = "style=\"font-family:" + str2 + ";font-size:10pt; font-weight: bold";
                        if (dstam.Tables[0].Rows[k][0].ToString().Length == 1)
                        {
                            str3 = str3 + ";color:blue";
                        }
                        else
                        {
                            str3 = str3 + ";color:DarkGreen";
                        }
                        str3 = str3 + "\"";
                        if (num6 == 0)
                        {
                            builder.Append("<td " + str3 + ">" + dstam.Tables[0].Rows[k][num6].ToString() + "</td>");
                            continue;
                        }
                        builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count - 1, " ", str3, ">", dstam.Tables[0].Rows[k][num6].ToString(), "</td>" }));
                        break;
                    }
                    if (num4 == 0M)
                    {
                        string str4 = "style=\"font-family:" + str2 + ";font-size:10pt; font-weight: bold";
                        if ((dstam.Tables[0].Rows[k][0].ToString().Length < 1) || ((dstam.Tables[0].Rows[k][0].ToString().Length >= 1) && ("A,B,C".IndexOf(dstam.Tables[0].Rows[k][0].ToString().Substring(0, 1)) > -1)))
                        {
                            str4 = str4 + ";color:blue";
                        }
                        else
                        {
                            str4 = str4 + ";color:DarkGreen";
                        }
                        str4 = str4 + "\"";
                        switch (num6)
                        {
                            case 0:
                            {
                                builder.Append("<td " + str4 + ">" + dstam.Tables[0].Rows[k][num6].ToString() + "</td>");
                                continue;
                            }
                            case 1:
                            {
                                builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns["tongcong"].Ordinal - 1, " ", str4, ">", dstam.Tables[0].Rows[k][num6].ToString(), "</td>" }));
                                num6 = dstam.Tables[0].Columns["tongcong"].Ordinal - 1;
                                continue;
                            }
                        }
                        if ((num6 >= dstam.Tables[0].Columns["tongcong"].Ordinal) && (num6 <= dstam.Tables[0].Columns["chiphids"].Ordinal))
                        {
                            try
                            {
                                decimal num7 = decimal.Parse(dstam.Tables[0].Rows[k][num6].ToString());
                                builder.Append("<td " + str4 + ">" + num7.ToString("###,###,###") + "</td>");
                            }
                            catch
                            {
                                builder.Append("<td></td>");
                            }
                        }
                        else
                        {
                            builder.Append("<td " + str4 + ">" + dstam.Tables[0].Rows[k][num6].ToString() + "</td>");
                        }
                    }
                    else if ((num6 >= dstam.Tables[0].Columns["tongcong"].Ordinal) && (num6 <= dstam.Tables[0].Columns["chiphids"].Ordinal))
                    {
                        try
                        {
                            decimal num8 = decimal.Parse(dstam.Tables[0].Rows[k][num6].ToString());
                            builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt\">" + num8.ToString("###,###,###") + "</td>");
                        }
                        catch
                        {
                            builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt\"></td>");
                        }
                    }
                    else
                    {
                        builder.Append("<td style=\"font-family:" + str2 + ";font-size:10pt\">" + dstam.Tables[0].Rows[k][num6].ToString() + "</td>");
                    }
                }
                builder.Append("</tr>");
            }
            builder.Append("</table>");
            writer.Write(builder);
            builder = new StringBuilder();
            builder.Append("<table>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            decimal num9 = decimal.Round(decimal.Parse(dstam.Tables[0].Rows[dstam.Tables[0].Rows.Count - 1]["bhyttra"].ToString()), 0);
            string str5 = new numbertotext().doiraso(num9.ToString());
            builder.Append(string.Concat(new object[] { "<td colspan=", dstam.Tables[0].Columns.Count - 1, " style=\"font-family:", str2, ";font-size:11pt\">Số tiền d? ngh? thanh to\x00e1n (viết bằng chữ): ", str5.Substring(0, 1).ToUpper(), str5.Substring(1), " đồng chẵn.</td>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            for (int m = 0; m < (dstam.Tables[0].Columns.Count - 3); m++)
            {
                builder.Append("<td></td>");
            }
            builder.Append(string.Concat(new object[] { "<th colspan=3 style=\"font-family:", str2, "\">Ng\x00e0y ", DateTime.Now.ToString("dd"), " th\x00e1ng ", DateTime.Now.ToString("MM"), " nam ", DateTime.Now.Year, "</th>" }));
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            builder.Append("<th style=\"font-family:" + str2 + "\">Người lập</th>");
            builder.Append("<td></td><td></td><td></td>");
            builder.Append("<th colspan=3 style=\"font-family:" + str2 + "\">Trưởng ph\x00f2ng KHTH</th>");
            builder.Append("<td></td><td></td><td></td><td></td><td></td><td></td>");
            builder.Append("<th colspan=2 style=\"font-family:" + str2 + "\">Kế to\x00e1n trưởng</th>");
            for (int n = 0x10; n < (dstam.Tables[0].Columns.Count - 3); n++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<th colspan=3 style=\"font-family:" + str2 + "\">Thủ trưởng đơn vị</th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td></td>");
            builder.Append("<td align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("<td></td><td></td><td></td>");
            builder.Append("<td colspan=3 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("<td></td><td></td><td></td><td></td><td></td><td></td>");
            builder.Append("<td colspan=2 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            for (int num12 = 0x10; num12 < (dstam.Tables[0].Columns.Count - 3); num12++)
            {
                builder.Append("<td></td>");
            }
            builder.Append("<td colspan=3 align=center style=\"font-family:" + str2 + "\">(K\x00fd, họ t\x00ean)</td>");
            builder.Append("</tr>");
            builder.Append("</table>");
            writer.Write(builder);
            writer.Close();
            try
            {
                this._sIDProcessExcelCurrent = this._lib.getid_process_Excel();
                this.oxl = new ApplicationClass();
                this.owb = this.oxl.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                this.osheet = (_Worksheet) this.owb.ActiveSheet;
                this.oxl.ActiveWindow.DisplayGridlines = true;
                this.orange = this.osheet.get_Range(this._lib.getIndex(dstam.Tables[0].Columns.Count - 3) + "1", this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "2");
                this.orange.MergeCells = true;
                this.orange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                int ordinal = dstam.Tables[0].Columns["bntra"].Ordinal;
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "8", this._lib.getIndex(ordinal) + "8");
                this.orange.ColumnWidth = 14;
                ordinal = dstam.Tables[0].Columns["bhyttra"].Ordinal + 1;
                this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "9", this._lib.getIndex(ordinal) + "9");
                this.orange.ColumnWidth = 14;
                this.orange.RowHeight = 60;
                this.orange = this.osheet.get_Range(this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "7", this._lib.getIndex(dstam.Tables[0].Columns.Count - 1) + "7");
                this.orange.ColumnWidth = 12;
                ordinal = dstam.Tables[0].Columns["tongcong"].Ordinal;
                int num14 = dstam.Tables[0].Columns["bntra"].Ordinal - 1;
                for (ordinal = ordinal; ordinal <= num14; ordinal++)
                {
                    this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + "9", this._lib.getIndex(ordinal) + "9");
                    this.orange.ColumnWidth = 13;
                }
                try
                {
                    int num15 = 10;
                    ordinal = dstam.Tables[0].Columns["manoidk"].Ordinal;
                    this.orange = this.osheet.get_Range(this._lib.getIndex(ordinal) + num15, this._lib.getIndex(ordinal) + (num15 + dstam.Tables[0].Rows.Count));
                    this.orange.NumberFormat = "@";
                    for (int num16 = 0; num16 < dstam.Tables[0].Rows.Count; num16++)
                    {
                        try
                        {
                            this.osheet.Cells[num15, ordinal + 1] = dstam.Tables[0].Rows[num16]["manoidk"].ToString();
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                }
                this.owb.Save();
                this.oxl.Quit();
                string idprocesslast = this._lib.getid_process_Excel();
                this._lib.f_end_process_Excel(this._sIDProcessExcelCurrent, idprocesslast);
                Process.Start(path);
            }
            catch
            {
                Process.Start(path);
            }
        }

        public void f_Noitru_xuatExcel_maumoi_808(bool print, DataSet dsdulieu, string tungay, string denngay, string fontchu)
        {
            DataSet set = this.f_Noitru_excel_mau41_getdata(dsdulieu, int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)));
            string str = "stt#hoten#namsinh#gioitinh#mathe#ma_dkbd#mabenh#ngay_vao#ngay_ra#ngaydtr#t_tongchi#t_xn#t_cdha#t_thuoc#t_mau#t_pttt#t_vtytth#t_vtyttt#t_dvktc#t_ktg#t_kham#t_vchuyen#t_bnct#t_bhtt#t_ngoaids#lydo_vv#benhkhac#noikcb#nam_qt#thang_qt#gt_tu#gt_den#diachi#tt_tngt#";
            for (int i = 0; i < set.Tables[0].Columns.Count; i++)
            {
                if (str.IndexOf(set.Tables[0].Columns[i].ColumnName + "#") == -1)
                {
                    set.Tables[0].Columns.RemoveAt(i--);
                }
            }
            this.f_Noitru_exp_excel_mau808_run(print, set, tungay, denngay, fontchu);
        }

        public void f_Noitru_xuatExcel_mauxp_38(bool print, DataSet dsdulieu, string tungay, string denngay, string fontchu)
        {
            DataSet set = this.f_Noitru_excel_mau38_getdata(dsdulieu, int.Parse(denngay.Substring(6, 4)), int.Parse(denngay.Substring(3, 2)));
            this.f_Noitru_exp_excel_mau38_run(print, set, tungay, denngay, fontchu);
        }

        private DataSet f_setSapXepCotTheoThuTu(DataSet dsdulieu, string dscotHienThi, char kitu)
        {
            string[] strArray = dscotHienThi.Trim(new char[] { kitu }).Split(new char[] { kitu });
            DataSet set = new DataSet();
            set.Tables.Add("mau41");
            for (int i = 0; i < strArray.Length; i++)
            {
                try
                {
                    set.Tables[0].Columns.Add(strArray[i]);
                }
                catch
                {
                }
            }
            foreach (DataRow row in dsdulieu.Tables[0].Rows)
            {
                DataRow row2 = set.Tables[0].NewRow();
                for (int j = 0; j < set.Tables[0].Columns.Count; j++)
                {
                    row2[j] = row[set.Tables[0].Columns[j].ColumnName].ToString();
                }
                set.Tables[0].Rows.Add(row2);
            }
            return set;
        }

        public decimal f_TinhTien_ChiPhiNgoaiDS(DataSet dschiphingoaids, string s_maicd, decimal TongTienBhytTra, decimal TienChiPhiVanChuyen, string s_madoituong)
        {
            decimal num = 0M;
            string str = "";
            string str2 = "";
            string str3 = "";
            string str4 = "";
            foreach (DataRow row in dschiphingoaids.Tables[0].Rows)
            {
                str = row["CONGTHUC_1"].ToString();
                str2 = row["CONGTHUC_2"].ToString();
                str3 = row["CONGTHUC_3"].ToString();
                str4 = "";
                if (str.Trim() != "")
                {
                    str4 = "( " + str + " ) ";
                }
                if (str2.Trim() != "")
                {
                    str4 = str4 + "( " + str2 + " )";
                }
                if (str3.Trim() != "")
                {
                    str4 = str4 + "( " + str3 + " )";
                }
                if (str2.Trim() != "")
                {
                    str4 = str4.Substring(0, str4.Length - 2);
                }
                try
                {
                    string str5 = s_maicd.Substring(0, 3);
                    if (str2.Trim() != "")
                    {
                        if ((s_maicd.IndexOf(str2.ToString()) != -1) && (row["loai"].ToString() == "2"))
                        {
                            num = TongTienBhytTra;
                        }
                        else if (str2.Trim() != "")
                        {
                            if ((s_maicd.IndexOf(str2.ToString()) != -1) && (row["loai"].ToString() != "2"))
                            {
                                num = TongTienBhytTra;
                            }
                            else if (str2.Trim() != "")
                            {
                                if ((s_maicd.IndexOf(str2.ToString()) != -1) && (row["loai"].ToString() == "7"))
                                {
                                    num = TongTienBhytTra;
                                }
                                else if (row["loai"].ToString() == "1")
                                {
                                    num = TienChiPhiVanChuyen;
                                }
                            }
                        }
                    }
                    if (s_madoituong.IndexOf("6") > -1)
                    {
                        num = TongTienBhytTra;
                    }
                }
                catch
                {
                }
            }
            return num;
        }

        public DataSet f_TinhTien_theoSoTheBHYT(DataSet dsxml, bool bQLTraiTuyen, string s_sothemoi_80, string s_sothemoi_18_95, string s_vitri_themoi15, string s_sothemoi15_80, string s_sothemoi15_95, int dodaithe15, string s_madoituong, bool b_tinhtienCPDS)
        {
            decimal num = this._lib.tile_traituyen();
            decimal iMuctinhbhytmoi = this._lib.iMuctinhbhytmoi;
            string s = "";
            s = this._lib.sNgaytinhmoibhytmoi;
            decimal num3 = iMuctinhbhytmoi;
            decimal num4 = 0M;
            decimal num5 = 100M;
            decimal tongTienBhytTra = 0M;
            decimal num7 = 0M;
            bool flag = false;
            string str2 = this._lib.s_sothe_80_09;
            BHYT1314 bhyt = new BHYT1314();
            DataSet dschiphingoaids = new DataSet();
            try
            {
                dschiphingoaids.ReadXml(@"..\..\..\xml\congthuc_ds.xml", XmlReadMode.ReadSchema);
            }
            catch
            {
                MessageBox.Show("Thiếu file congthuc_ds.xml trong thư mục xml", "Th\x00f4ng b\x00e1o");
                return dsxml;
            }
            foreach (DataRow row in dsxml.Tables[0].Rows)
            {
                flag = this._lib.StringToDate(row["ngayvao"].ToString().Substring(0, 10)) > this._lib.StringToDate(s);
                num4 = decimal.Parse(row["tongcong"].ToString());
                tongTienBhytTra = num4;
                num7 = 0M;
                num5 = 100M;
                if ((row["traituyen"].ToString() != "0") && bQLTraiTuyen)
                {
                    num5 = num;
                }
                try
                {
                    num5 = Convert.ToDecimal(bhyt.f_get_TyLeThe(row["sothe"].ToString(), row["ngayra"].ToString()));
                    if (row["traituyen"].ToString() != "0")
                    {
                        num5 = (num5 * num) / 100M;
                    }
                    if (num4 <= iMuctinhbhytmoi)
                    {
                        num5 = 100M;
                    }
                }
                catch
                {
                }
                tongTienBhytTra = (num4 * num5) / 100M;
                num7 = num4 - tongTienBhytTra;
                row["bhyttra"] = tongTienBhytTra;
                row["bntra"] = num7;
                row["tylebhyt"] = num5;
                if (b_tinhtienCPDS)
                {
                    row["CHIPHIDS"] = this.f_TinhTien_ChiPhiNgoaiDS(dschiphingoaids, row["MAICD"].ToString(), tongTienBhytTra, decimal.Parse(row["ST_10"].ToString()), s_madoituong);
                }
                else if (row["madoituong"].ToString().IndexOf("6") > -1)
                {
                    row["CHIPHIDS"] = decimal.Parse(row["bhyttra"].ToString());
                }
                else
                {
                    row["CHIPHIDS"] = 0;
                }
            }
            return dsxml;
        }

        private decimal f_tinhtong_cttt9324(DataSet vdsct, string tencot, string exp)
        {
            decimal num = 0M;
            for (int i = 0; i < vdsct.Tables.Count; i++)
            {
                try
                {
                    num += Convert.ToDecimal(vdsct.Tables[i].Compute("sum(" + tencot + ")", exp).ToString());
                }
                catch
                {
                }
            }
            return num;
        }

        private decimal f_tinhtong_cttt9324(DataSet vdsthuoc, DataSet vdscls, string tencot, string exp)
        {
            decimal num = 0M;
            try
            {
                num += Convert.ToDecimal(vdsthuoc.Tables[0].Compute("sum(" + tencot + ")", exp).ToString());
            }
            catch
            {
            }
            try
            {
                num += Convert.ToDecimal(vdscls.Tables[0].Compute("sum(" + tencot + ")", exp).ToString());
            }
            catch
            {
            }
            return num;
        }

        private decimal f_tinhtong_cttt9324(DataSet vdsthuoc, DataSet vdscls, string tencot, string exp, bool roundchitiet)
        {
            decimal num = 0M;
            try
            {
                foreach (DataRow row in vdsthuoc.Tables[0].Select(exp))
                {
                    num += Math.Round(Convert.ToDecimal(row[tencot].ToString()));
                }
            }
            catch
            {
            }
            try
            {
                foreach (DataRow row2 in vdscls.Tables[0].Select(exp))
                {
                    num += Math.Round(Convert.ToDecimal(row2[tencot].ToString()));
                }
            }
            catch
            {
            }
            return num;
        }

        public void f_Update_MaVaoVien(int loaibn, string mmyy)
        {
            try
            {
                if (mmyy == "")
                {
                    mmyy = DateTime.Now.ToString("MMyy");
                }
                if (loaibn == 1)
                {
                    this._lib.execute_data("update " + this._lib.user + ".benhandt  set mavaovien = maql where mavaovien=0");
                }
                else
                {
                    this._lib.execute_data("update " + this._lib.user + mmyy + ".benhandt  set mavaovien = maql where mavaovien=0");
                    string str = "select mavaovien from " + this._lib.user + mmyy + ".vi_benhandt" + ((loaibn == 2) ? "ngtr" : "");
                    this._lib.execute_data("update " + this._lib.user + "d" + mmyy + ".bhytkb a set mavaovien =(" + str.Trim(new char[] { '_' }) + " b where a.maql=b.maql and a.mabn=b.mabn)");
                }
            }
            catch
            {
            }
        }

        private void f_update_mavaovien_324moi(string mavaovien, string expdieukien, ref DataSet dsdulieu)
        {
            try
            {
                DataRow[] rowArray = dsdulieu.Tables[0].Select(expdieukien);
                if (rowArray.Length > 0)
                {
                    for (int i = 0; i < rowArray.Length; i++)
                    {
                        rowArray[i]["mavaovien"] = mavaovien;
                    }
                    dsdulieu.AcceptChanges();
                }
            }
            catch
            {
            }
        }

        public void f_Update_NgayYLenh(string mmyy, string mmyytruoc)
        {
            if (mmyy == "")
            {
                mmyy = DateTime.Now.ToString("MMyy");
            }
            string user = this._lib.user;
            this._lib.execute_data("alter table " + user + mmyy + ".v_thvpct  add updngayylenh numeric(1) default 0");
            this._lib.execute_data("alter table " + user + mmyytruoc + ".v_thvpct  add updngayylenh numeric(1) default 0");
            DataSet set = this._lib.get_data("select  distinct b.mabn,b.maql,a.id,a.stt,a.makp,a.madoituong,a.mavp,a.soluong,a.dongia from " + user + mmyy + ".v_ttrvct a inner join " + user + mmyy + ".v_ttrvds b on a.id=b.id where a.ngay is null order by a.id,a.stt,a.mavp");
            DataSet set2 = this._lib.get_data("select  distinct '" + mmyy + "' as mmyy, b.mabn,b.maql,a.id,a.makp,a.madoituong,a.mavp,a.soluong,a.dongia,a.idttrv,updngayylenh,to_char(ngay,'yyyymmddhh24mi') as ngay from " + user + mmyy + ".v_thvpct a inner join " + user + mmyy + ".v_thvpll b on a.id=b.id where updngayylenh<>1  order by a.id,a.mavp,ngay");
            try
            {
                set2.Merge(this._lib.get_data("select  distinct '" + mmyytruoc + "' as mmyy, b.mabn,b.maql,a.id,a.makp,a.madoituong,a.mavp,a.soluong,a.dongia,a.idttrv,updngayylenh,to_char(ngay,'yyyymmddhh24mi') as ngay from " + user + mmyytruoc + ".v_thvpct a inner join " + user + mmyytruoc + ".v_thvpll b on a.id=b.id where updngayylenh<>1 order by a.id,a.mavp,ngay"));
            }
            catch
            {
            }
            foreach (DataRow row in set.Tables[0].Rows)
            {
                DataRow[] rowArray = set2.Tables[0].Select("maql=" + row["maql"].ToString() + " and madoituong=" + row["madoituong"].ToString() + " and mavp=" + row["mavp"].ToString() + " and soluong=" + row["soluong"].ToString() + " and updngayylenh<>1 and dongia=" + row["dongia"].ToString() + " and makp='" + row["makp"].ToString() + "'");
                if (rowArray.Length > 0)
                {
                    this._lib.execute_data("update " + user + mmyy + ".v_ttrvct set ngay=to_date('" + rowArray[0]["ngay"].ToString() + "','yyyymmddhh24mi') where id=" + row["id"].ToString() + " and madoituong=" + rowArray[0]["madoituong"].ToString() + " and mavp=" + rowArray[0]["mavp"].ToString() + " and soluong=" + rowArray[0]["soluong"].ToString() + " and dongia=" + rowArray[0]["dongia"].ToString() + " and makp='" + rowArray[0]["makp"].ToString() + "'");
                    this._lib.execute_data("update " + user + rowArray[0]["mmyy"].ToString() + ".v_thvpct set updngayylenh=1 where id=" + rowArray[0]["id"].ToString() + " and madoituong=" + row["madoituong"].ToString() + " and mavp=" + row["mavp"].ToString() + " and soluong=" + row["soluong"].ToString() + " and dongia=" + row["dongia"].ToString() + " and makp='" + row["makp"].ToString() + "'");
                    rowArray[0]["updngayylenh"] = "1";
                }
                else
                {
                    rowArray = set2.Tables[0].Select("maql=" + row["maql"].ToString() + " and madoituong=" + row["madoituong"].ToString() + " and mavp=" + row["mavp"].ToString() + " and updngayylenh<>1 and dongia=" + row["dongia"].ToString() + " and makp='" + row["makp"].ToString() + "'");
                    if (rowArray.Length > 0)
                    {
                        this._lib.execute_data("update " + user + mmyy + ".v_ttrvct set ngay=to_date('" + rowArray[0]["ngay"].ToString() + "','yyyymmddhh24mi') where id=" + row["id"].ToString() + " and madoituong=" + rowArray[0]["madoituong"].ToString() + " and mavp=" + rowArray[0]["mavp"].ToString() + " and dongia=" + rowArray[0]["dongia"].ToString() + " and makp='" + rowArray[0]["makp"].ToString() + "'");
                        this._lib.execute_data("update " + user + rowArray[0]["mmyy"].ToString() + ".v_thvpct set updngayylenh=1 where id=" + rowArray[0]["id"].ToString() + " and madoituong=" + row["madoituong"].ToString() + " and mavp=" + row["mavp"].ToString() + " and dongia=" + row["dongia"].ToString() + " and makp='" + row["makp"].ToString() + "'");
                        rowArray[0]["updngayylenh"] = "1";
                    }
                }
            }
        }

        public void f_update_quyenso_324moi(ref DataRow vdr)
        {
            try
            {
                string str = "";
                str = vdr["idttrv"].ToString().Substring(2, 2) + vdr["idttrv"].ToString().Substring(0, 2);
                DataSet set = this._lib.get_data("select a.sobienlai,b.sohieu,c.id useridvp,c.hoten as nguoithu from " + this._lib.user + str + ".v_ttrvll a join v_quyenso b on a.quyenso=b.id join v_dlogin c on c.id=a.userid where a.id=" + vdr["idttrv"].ToString());
                vdr["sobienlai"] = set.Tables[0].Rows[0]["sobienlai"].ToString();
                vdr["ten_quyenso"] = set.Tables[0].Rows[0]["sohieu"].ToString();
                vdr["userid_thuvp"] = set.Tables[0].Rows[0]["useridvp"].ToString();
                vdr["ten_thuvp"] = set.Tables[0].Rows[0]["nguoithu"].ToString();
                vdr.AcceptChanges();
            }
            catch
            {
            }
        }

        private void f_update_soam_324moi(string soluongam, string expdieukien, ref DataSet dsdulieu)
        {
            try
            {
                decimal num = Math.Abs(decimal.Parse(soluongam));
                DataRow[] rowArray = dsdulieu.Tables[0].Select(expdieukien + " and soluong=" + num, "soluong desc");
                if (rowArray.Length > 0)
                {
                    rowArray[0]["soluong"] = 0;
                    rowArray[0]["sotien"] = 0;
                    dsdulieu.AcceptChanges();
                }
                else
                {
                    rowArray = dsdulieu.Tables[0].Select(expdieukien, "soluong desc");
                    if (rowArray.Length > 0)
                    {
                        for (int i = 0; i < rowArray.Length; i++)
                        {
                            if (num <= 0M)
                            {
                                break;
                            }
                            if (decimal.Parse(rowArray[i]["soluong"].ToString()) > num)
                            {
                                num = decimal.Parse(rowArray[i]["soluong"].ToString()) - num;
                                rowArray[i]["soluong"] = num;
                                rowArray[i]["sotien"] = decimal.Parse(rowArray[i]["soluong"].ToString()) * decimal.Parse(rowArray[i]["dongia"].ToString());
                                break;
                            }
                            num -= decimal.Parse(rowArray[i]["soluong"].ToString());
                            rowArray[i]["soluong"] = 0;
                            rowArray[i]["sotien"] = 0;
                        }
                        dsdulieu.AcceptChanges();
                    }
                }
            }
            catch
            {
            }
        }

        public void f_Update_TienKham(string mmyy)
        {
            if (mmyy == "")
            {
                mmyy = DateTime.Now.ToString("MMyy");
            }
            string user = this._lib.user;
            try
            {
                int count = this._lib.get_data("select t_kham from " + user + "d" + mmyy + ".bhytkb  where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this._lib.get_data("alter table " + user + "d" + mmyy + ".bhytkb add t_kham number(15,2) default 0 ");
            }
            try
            {
                int num2 = this._lib.get_data("select miencongkham from " + user + "d" + mmyy + ".bhytkb  where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this._lib.get_data("alter table " + user + "d" + mmyy + ".bhytkb add miencongkham number(1) default 0 ");
            }
            try
            {
                int num3 = this._lib.get_data("select solieubhyt from " + user + "d" + mmyy + ".bhytkb  where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this._lib.get_data("alter table " + user + "d" + mmyy + ".bhytkb add solieubhyt number(1) default 0 ");
            }
            try
            {
                int num4 = this._lib.get_data("select solieubhyt from " + user + mmyy + ".v_ttrvll  where 1=0 ").Tables[0].Rows.Count;
            }
            catch
            {
                this._lib.get_data("alter table " + user + mmyy + ".v_ttrvll add solieubhyt number(1) default 0 ");
            }
            DataSet set = this._lib.get_data("select mabn,maql,id,makp,to_char(ngay,'dd/mm/yyyy') ngay,congkham from " + user + "d" + mmyy + ".bhytkb  where t_kham=0 and miencongkham=0  order by ngay,mabn,to_char(maql)");
            for (int i = 0; i < set.Tables[0].Rows.Count; i++)
            {
                this._lib.execute_data("update " + user + "d" + mmyy + ".bhytkb set t_kham=congkham where id=" + set.Tables[0].Rows[i]["id"].ToString());
                this._lib.execute_data("update " + user + "d" + mmyy + ".bhytkb set t_kham=cls where congkham=0 and id=" + set.Tables[0].Rows[i]["id"].ToString());
                int num6 = 0;
                int num7 = 2;
                num6 = i + 1;
                while (num6 < set.Tables[0].Rows.Count)
                {
                    if ((set.Tables[0].Rows[i]["ngay"].ToString() != set.Tables[0].Rows[num6]["ngay"].ToString()) || (set.Tables[0].Rows[i]["mabn"].ToString() != set.Tables[0].Rows[num6]["mabn"].ToString()))
                    {
                        break;
                    }
                    if (num7 < 6)
                    {
                        this._lib.execute_data("update " + user + "d" + mmyy + ".bhytkb set t_kham=congkham*" + ((num7 < 5) ? "0.3" : ((num7 == 5) ? "0.1" : "0")) + " where id=" + set.Tables[0].Rows[num6]["id"].ToString());
                    }
                    else
                    {
                        this._lib.execute_data("update " + user + "d" + mmyy + ".bhytkb set t_kham=0,miencongkham=1 where id=" + set.Tables[0].Rows[num6]["id"].ToString());
                    }
                    num7++;
                    num6++;
                }
                i = num6 - 1;
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception exception)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + exception.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public enum enMauBaoCaoBHYT
        {
            Ngoai_25act = 9,
            Ngoai_25ath = 7,
            Ngoai_41cot = 2,
            Ngoai_79_mau2 = 20,
            Ngoai_79aHD = 3,
            Ngoai_CV808 = 1,
            Ngoai_TT2348 = 11,
            Ngoai_TT324 = 0x10,
            Ngoai_TT324moi = 0x12,
            Ngoai_TT4210 = 0x17,
            Ngoai_TT9324 = 14,
            Noi_26act = 10,
            Noi_26ath = 8,
            Noi_38cot_xp = 13,
            Noi_41cot = 6,
            Noi_80_mau2 = 0x15,
            Noi_80aHD = 4,
            Noi_CV808 = 5,
            Noi_TT2348 = 12,
            Noi_TT324 = 0x11,
            Noi_TT324moi = 0x13,
            Noi_TT4210 = 0x16,
            Noi_TT9324 = 15
        }
    }
}

