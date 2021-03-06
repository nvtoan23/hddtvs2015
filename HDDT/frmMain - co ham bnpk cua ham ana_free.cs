using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using LibMedi;
using libHDDT;
using System.Linq;
using dichso;
using dllbhyt;
using dlltinhcp;

namespace HDDT
{
    public partial class frmMain : Form
    {
        #region Khai báo
        private AccessDataApi m = new AccessDataApi();
        private AccessDataAPI hddt = new AccessDataAPI();
        private LibClass _cacc;
        private BHYT1314 _ccpbhyt;
        private LibMaubaocao _cmaubhyt;
        //private libacc _lib;
        private DataSet ds = new DataSet();
        private DataSet dsHDDT = new DataSet();
        private DataSet dsAna = new DataSet();
        private DataSet hMasterTT_temp;
        private DataSet hMasterTTRV_temp;
        private DataTable dtth_TTRV = new DataTable();
        private DataTable dtct_TTRV = new DataTable();
        private DataTable dtmaxSttAna = new DataTable();
        private DataTable ds_temp = new DataTable();
        private DataTable dtth = new DataTable();
        private DataTable dtct = new DataTable();
        private DataTable dttongtien = new DataTable();
        private int _iKyGD = 1;
        private int _iLoaiGD = 1;
        private bool bAna, bHDDT;
        private string sql = "", id_temp = "", user = "", usermmyy = "", s_giavp_dmbd = "", s_hanhchinh = "", s_dulieu = "", paratt = "";
        private string[] _s_arrCap = new string[] { "", "\r\n  ", "\r\n    ", "\r\n      ", "\r\n        " };
        private string _s_xmlpath = @"..\xml\";
        private string _sGlivecThuMucGoc = "";
        private string _sGlivecThuMucXoa = "";
        private string _sMabv = "";
        private string _sNamGD = "";
        private string _sNgayLap = "";
        private string _sTenbv = "";
        private string _sUser = "";
        public string ipApi = "";
        const int TamUng = 1, ThuTrucTiep = 2, TTRV = 3, bhNgoaiTru = 4, NhaThuoc = 5;
        private decimal l_tongtien = 0;
        private DataGridView dtg;
        private Button bStart;
        private Timer timer_insert;
        private IContainer components;
        private Button bClose;
        private Label label1;
        private Label label2;
        private DateTimePicker tungay;
        private DateTimePicker denngay;
        private NumericUpDown nbPhut;
        private Label label3;
        private Label label4;
        private Button bPause;
        private Label label5;
        private ComboBox cbMMYY;
        private Label record;
        private ComboBox cbLoaiVP;
        private Label label6;
        private TextBox tim;
        private DataGridViewTextBoxColumn c_sohoso;
        private DataGridViewTextBoxColumn c_loaivp;
        private DataGridViewTextBoxColumn c_makhachhang;
        private DataGridViewTextBoxColumn c_tenkhachhang;
        private DataGridViewTextBoxColumn c_phai;
        private DataGridViewTextBoxColumn c_namsinh;
        private DataGridViewTextBoxColumn c_diachi;
        private DataGridViewTextBoxColumn quyenso;
        private DataGridViewTextBoxColumn sobienlai;
        private DataGridViewTextBoxColumn c_ngaylap;
        private GroupBox gGetData;
        private TabControl TabHDDT;
        private TabPage tab_hsba;
        private TabPage tab_hddt;
        private ProgressBar progressBar1;
        public int interval = 60000;
        private CheckBox chkAna;
        private string s_bc1 = "", s_bc2 = "";
        private bool m_giavpbangdongiacongvattu = false;
        private DataSet m_ds = new DataSet();
        private DataSet m_ds1 = new DataSet();
        private DataSet m_ds2 = new DataSet();
        private DataSet _dsChiPhi = new DataSet();
        private DataSet _dsChiPhiCT = new DataSet();
        private DataSet _dsDuLieu = new DataSet();
        private DataSet _dtketqua = new DataSet();
        private decimal sttAna = 0;
        private CheckBox chkAll1;
        private RadioButton rd2;
        private RadioButton rd1;
        private RadioButton rd5;
        private RadioButton rd4;
        private RadioButton rd3;
        private CheckBox chkTheokhoa;
        private CheckBox chkChitiet;
        private CheckBox chkkhongtinhchenhlech;
        private CheckBox chkAll;
        private TreeView tree_Field;
        private TreeView tree_Loai;
        private TreeView tree_Loaibn;
        private CheckBox chkLoaibn;
        private Timer timer1;
        private TabPage tabPage1;
        private DataGridView dataGridView1;
        private TabPage tabPage2;
        private DataGridView dataGridView2;
        private GroupBox groupBox1;
        private TabPage tabPage3;
        private Button bKetxuat;

        private System.Threading.Thread tMediSoft_HDDT_TT; //thu trực tiếp
        private System.Threading.Thread tMediSoft_HDDT_TTRV; //TT ra viện
        private System.Threading.Thread tHDDT_Ana; //Đẩy sang table THANHTOANVIENPHI những hđơn đã đc xuất trên hddt
        private System.Threading.Thread tHDDT_Ana_free; //Đẩy sang table THANHTOANVIENPHI những hđơn ko phải xuất HĐ
        private System.Threading.Thread tHDDT_Ana_KyQuy;

        private GroupBox groupBox3;
        private GroupBox groupBox2;
        private RadioButton rdb_noitru;
        private RadioButton rdb_ngoaitru;
        private ComboBox cmbmaubaocao;
        private TextBox txtmabn;
        private Label label7;
        private Label label8;
        private CheckBox chkxuatkhoa;
        private CheckBox chklaythuocglivec;
        private Label lblrefesh;
        public TextBox txtIP;
        private CheckBox chkExcel;
        private Button bHDDTStart;
        private Label label9;
        #endregion
        public frmMain()
        {
            InitializeComponent();
            if (m.Mabv != "101.1.05")
            {
                MessageBox.Show("Bệnh viên chưa đăng ký sử dụng !");
                this.Close();
            }
        }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.dtg = new System.Windows.Forms.DataGridView();
            this.c_sohoso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_loaivp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_makhachhang = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_tenkhachhang = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_phai = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_namsinh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_diachi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.quyenso = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sobienlai = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.c_ngaylap = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bStart = new System.Windows.Forms.Button();
            this.bClose = new System.Windows.Forms.Button();
            this.timer_insert = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tungay = new System.Windows.Forms.DateTimePicker();
            this.denngay = new System.Windows.Forms.DateTimePicker();
            this.nbPhut = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.bPause = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cbMMYY = new System.Windows.Forms.ComboBox();
            this.record = new System.Windows.Forms.Label();
            this.cbLoaiVP = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tim = new System.Windows.Forms.TextBox();
            this.gGetData = new System.Windows.Forms.GroupBox();
            this.txtIP = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.TabHDDT = new System.Windows.Forms.TabControl();
            this.tab_hsba = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkExcel = new System.Windows.Forms.CheckBox();
            this.lblrefesh = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdb_noitru = new System.Windows.Forms.RadioButton();
            this.rdb_ngoaitru = new System.Windows.Forms.RadioButton();
            this.cmbmaubaocao = new System.Windows.Forms.ComboBox();
            this.txtmabn = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.chkxuatkhoa = new System.Windows.Forms.CheckBox();
            this.chklaythuocglivec = new System.Windows.Forms.CheckBox();
            this.rd3 = new System.Windows.Forms.RadioButton();
            this.rd4 = new System.Windows.Forms.RadioButton();
            this.rd5 = new System.Windows.Forms.RadioButton();
            this.chkChitiet = new System.Windows.Forms.CheckBox();
            this.chkkhongtinhchenhlech = new System.Windows.Forms.CheckBox();
            this.chkTheokhoa = new System.Windows.Forms.CheckBox();
            this.chkAna = new System.Windows.Forms.CheckBox();
            this.tree_Loaibn = new System.Windows.Forms.TreeView();
            this.chkLoaibn = new System.Windows.Forms.CheckBox();
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.tree_Field = new System.Windows.Forms.TreeView();
            this.tree_Loai = new System.Windows.Forms.TreeView();
            this.chkAll1 = new System.Windows.Forms.CheckBox();
            this.rd2 = new System.Windows.Forms.RadioButton();
            this.rd1 = new System.Windows.Forms.RadioButton();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tab_hddt = new System.Windows.Forms.TabPage();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.bKetxuat = new System.Windows.Forms.Button();
            this.bHDDTStart = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dtg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbPhut)).BeginInit();
            this.gGetData.SuspendLayout();
            this.TabHDDT.SuspendLayout();
            this.tab_hsba.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tab_hddt.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtg
            // 
            this.dtg.AllowUserToAddRows = false;
            this.dtg.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Bisque;
            this.dtg.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dtg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtg.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dtg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dtg.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.c_sohoso,
            this.c_loaivp,
            this.c_makhachhang,
            this.c_tenkhachhang,
            this.c_phai,
            this.c_namsinh,
            this.c_diachi,
            this.quyenso,
            this.sobienlai,
            this.c_ngaylap});
            this.dtg.Cursor = System.Windows.Forms.Cursors.Default;
            this.dtg.Location = new System.Drawing.Point(3, 29);
            this.dtg.Name = "dtg";
            this.dtg.ReadOnly = true;
            this.dtg.RowHeadersVisible = false;
            this.dtg.Size = new System.Drawing.Size(837, 356);
            this.dtg.TabIndex = 0;
            // 
            // c_sohoso
            // 
            this.c_sohoso.DataPropertyName = "sohoso";
            this.c_sohoso.HeaderText = "Số hồ sơ";
            this.c_sohoso.Name = "c_sohoso";
            this.c_sohoso.ReadOnly = true;
            this.c_sohoso.Width = 110;
            // 
            // c_loaivp
            // 
            this.c_loaivp.DataPropertyName = "loaivp";
            this.c_loaivp.HeaderText = "Loại VP";
            this.c_loaivp.Name = "c_loaivp";
            this.c_loaivp.ReadOnly = true;
            // 
            // c_makhachhang
            // 
            this.c_makhachhang.DataPropertyName = "makhachhang";
            this.c_makhachhang.HeaderText = "Mã KH";
            this.c_makhachhang.Name = "c_makhachhang";
            this.c_makhachhang.ReadOnly = true;
            this.c_makhachhang.Width = 70;
            // 
            // c_tenkhachhang
            // 
            this.c_tenkhachhang.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.c_tenkhachhang.DataPropertyName = "tenkhachhang";
            this.c_tenkhachhang.HeaderText = "Tên khách hàng";
            this.c_tenkhachhang.Name = "c_tenkhachhang";
            this.c_tenkhachhang.ReadOnly = true;
            // 
            // c_phai
            // 
            this.c_phai.DataPropertyName = "phai";
            this.c_phai.HeaderText = "Phái";
            this.c_phai.Name = "c_phai";
            this.c_phai.ReadOnly = true;
            this.c_phai.Width = 50;
            // 
            // c_namsinh
            // 
            this.c_namsinh.DataPropertyName = "namsinh";
            this.c_namsinh.HeaderText = "Năm sinh";
            this.c_namsinh.Name = "c_namsinh";
            this.c_namsinh.ReadOnly = true;
            this.c_namsinh.Width = 70;
            // 
            // c_diachi
            // 
            this.c_diachi.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.c_diachi.DataPropertyName = "diachi";
            this.c_diachi.HeaderText = "Địa chỉ";
            this.c_diachi.Name = "c_diachi";
            this.c_diachi.ReadOnly = true;
            // 
            // quyenso
            // 
            this.quyenso.DataPropertyName = "quyenso";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.quyenso.DefaultCellStyle = dataGridViewCellStyle2;
            this.quyenso.HeaderText = "Quyển sổ";
            this.quyenso.Name = "quyenso";
            this.quyenso.ReadOnly = true;
            this.quyenso.Width = 70;
            // 
            // sobienlai
            // 
            this.sobienlai.DataPropertyName = "sobienlai";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.sobienlai.DefaultCellStyle = dataGridViewCellStyle3;
            this.sobienlai.HeaderText = "Số BL";
            this.sobienlai.Name = "sobienlai";
            this.sobienlai.ReadOnly = true;
            this.sobienlai.Width = 60;
            // 
            // c_ngaylap
            // 
            this.c_ngaylap.DataPropertyName = "ngaylap";
            this.c_ngaylap.HeaderText = "Ngày ";
            this.c_ngaylap.Name = "c_ngaylap";
            this.c_ngaylap.ReadOnly = true;
            this.c_ngaylap.Width = 80;
            // 
            // bStart
            // 
            this.bStart.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.bStart.Location = new System.Drawing.Point(396, 476);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(110, 33);
            this.bStart.TabIndex = 1;
            this.bStart.Text = "Medi - Ana";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // bClose
            // 
            this.bClose.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.bClose.Location = new System.Drawing.Point(699, 476);
            this.bClose.Name = "bClose";
            this.bClose.Size = new System.Drawing.Size(94, 33);
            this.bClose.TabIndex = 2;
            this.bClose.Text = "Kết thúc";
            this.bClose.UseVisualStyleBackColor = true;
            this.bClose.Click += new System.EventHandler(this.bClose_Click);
            // 
            // timer_insert
            // 
            this.timer_insert.Interval = 3000;
            this.timer_insert.Tick += new System.EventHandler(this.timer_insert_Tick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Từ ngày :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(144, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Đến ngày :";
            // 
            // tungay
            // 
            this.tungay.CustomFormat = "dd/MM/yyyy";
            this.tungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.tungay.Location = new System.Drawing.Point(57, 14);
            this.tungay.Name = "tungay";
            this.tungay.Size = new System.Drawing.Size(85, 20);
            this.tungay.TabIndex = 7;
            this.tungay.ValueChanged += new System.EventHandler(this.tungay_ValueChanged);
            // 
            // denngay
            // 
            this.denngay.CustomFormat = "dd/MM/yyyy";
            this.denngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.denngay.Location = new System.Drawing.Point(201, 14);
            this.denngay.Name = "denngay";
            this.denngay.Size = new System.Drawing.Size(85, 20);
            this.denngay.TabIndex = 8;
            this.denngay.ValueChanged += new System.EventHandler(this.denngay_ValueChanged);
            // 
            // nbPhut
            // 
            this.nbPhut.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nbPhut.Location = new System.Drawing.Point(366, 14);
            this.nbPhut.Name = "nbPhut";
            this.nbPhut.Size = new System.Drawing.Size(41, 20);
            this.nbPhut.TabIndex = 9;
            this.nbPhut.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(292, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Lấy data sau :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(408, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(28, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "phút";
            // 
            // bPause
            // 
            this.bPause.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bPause.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.bPause.Location = new System.Drawing.Point(622, 476);
            this.bPause.Name = "bPause";
            this.bPause.Size = new System.Drawing.Size(74, 33);
            this.bPause.TabIndex = 12;
            this.bPause.Text = "Dừng";
            this.bPause.UseVisualStyleBackColor = true;
            this.bPause.Click += new System.EventHandler(this.bPause_Click);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(19, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Số liệu tháng :";
            this.label5.Visible = false;
            // 
            // cbMMYY
            // 
            this.cbMMYY.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbMMYY.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbMMYY.FormattingEnabled = true;
            this.cbMMYY.Location = new System.Drawing.Point(94, 151);
            this.cbMMYY.Name = "cbMMYY";
            this.cbMMYY.Size = new System.Drawing.Size(81, 21);
            this.cbMMYY.TabIndex = 14;
            this.cbMMYY.Visible = false;
            // 
            // record
            // 
            this.record.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.record.AutoSize = true;
            this.record.Location = new System.Drawing.Point(22, 481);
            this.record.Name = "record";
            this.record.Size = new System.Drawing.Size(65, 13);
            this.record.TabIndex = 15;
            this.record.Text = "Tổng cộng :";
            // 
            // cbLoaiVP
            // 
            this.cbLoaiVP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbLoaiVP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLoaiVP.FormattingEnabled = true;
            this.cbLoaiVP.Items.AddRange(new object[] {
            "Tất cả",
            "Tạm ứng",
            "Thu trực tiếp",
            "TTRV",
            "BHYT Ngoại trú",
            "Nhà Thuốc"});
            this.cbLoaiVP.Location = new System.Drawing.Point(212, 151);
            this.cbLoaiVP.Name = "cbLoaiVP";
            this.cbLoaiVP.Size = new System.Drawing.Size(107, 21);
            this.cbLoaiVP.TabIndex = 16;
            this.cbLoaiVP.Visible = false;
            this.cbLoaiVP.SelectedIndexChanged += new System.EventHandler(this.cbLoaiVP_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(181, 154);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Loại :";
            this.label6.Visible = false;
            // 
            // tim
            // 
            this.tim.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tim.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tim.Location = new System.Drawing.Point(6, 3);
            this.tim.Name = "tim";
            this.tim.Size = new System.Drawing.Size(834, 22);
            this.tim.TabIndex = 18;
            this.tim.Text = "Tìm kiếm";
            this.tim.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tim.MouseClick += new System.Windows.Forms.MouseEventHandler(this.tim_MouseClick);
            this.tim.TextChanged += new System.EventHandler(this.tim_TextChanged);
            this.tim.Validated += new System.EventHandler(this.tim_Validated);
            // 
            // gGetData
            // 
            this.gGetData.Controls.Add(this.bKetxuat);
            this.gGetData.Controls.Add(this.txtIP);
            this.gGetData.Controls.Add(this.tungay);
            this.gGetData.Controls.Add(this.nbPhut);
            this.gGetData.Controls.Add(this.label9);
            this.gGetData.Controls.Add(this.label3);
            this.gGetData.Controls.Add(this.denngay);
            this.gGetData.Controls.Add(this.label4);
            this.gGetData.Controls.Add(this.label2);
            this.gGetData.Controls.Add(this.label1);
            this.gGetData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gGetData.Location = new System.Drawing.Point(15, 11);
            this.gGetData.Name = "gGetData";
            this.gGetData.Size = new System.Drawing.Size(851, 44);
            this.gGetData.TabIndex = 19;
            this.gGetData.TabStop = false;
            this.gGetData.Text = " LẤY DỮ LIỆU ";
            // 
            // txtIP
            // 
            this.txtIP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtIP.Location = new System.Drawing.Point(629, 14);
            this.txtIP.Name = "txtIP";
            this.txtIP.Size = new System.Drawing.Size(209, 20);
            this.txtIP.TabIndex = 12;
            this.txtIP.Text = "103.3.253.54:81";
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(566, 17);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(57, 13);
            this.label9.TabIndex = 10;
            this.label9.Text = "IP Server :";
            // 
            // TabHDDT
            // 
            this.TabHDDT.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TabHDDT.Controls.Add(this.tab_hsba);
            this.TabHDDT.Controls.Add(this.tabPage3);
            this.TabHDDT.Controls.Add(this.tabPage1);
            this.TabHDDT.Controls.Add(this.tabPage2);
            this.TabHDDT.Controls.Add(this.tab_hddt);
            this.TabHDDT.Location = new System.Drawing.Point(15, 61);
            this.TabHDDT.Name = "TabHDDT";
            this.TabHDDT.SelectedIndex = 0;
            this.TabHDDT.Size = new System.Drawing.Size(851, 411);
            this.TabHDDT.TabIndex = 21;
            // 
            // tab_hsba
            // 
            this.tab_hsba.Controls.Add(this.groupBox1);
            this.tab_hsba.Controls.Add(this.chkAna);
            this.tab_hsba.Controls.Add(this.tree_Loaibn);
            this.tab_hsba.Controls.Add(this.chkLoaibn);
            this.tab_hsba.Controls.Add(this.chkAll);
            this.tab_hsba.Controls.Add(this.tree_Field);
            this.tab_hsba.Controls.Add(this.tree_Loai);
            this.tab_hsba.Controls.Add(this.chkAll1);
            this.tab_hsba.Controls.Add(this.rd2);
            this.tab_hsba.Controls.Add(this.rd1);
            this.tab_hsba.Location = new System.Drawing.Point(4, 22);
            this.tab_hsba.Name = "tab_hsba";
            this.tab_hsba.Padding = new System.Windows.Forms.Padding(3);
            this.tab_hsba.Size = new System.Drawing.Size(843, 385);
            this.tab_hsba.TabIndex = 1;
            this.tab_hsba.Text = "Cấu hình Ana";
            this.tab_hsba.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.rd3);
            this.groupBox1.Controls.Add(this.cbLoaiVP);
            this.groupBox1.Controls.Add(this.rd4);
            this.groupBox1.Controls.Add(this.cbMMYY);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.rd5);
            this.groupBox1.Controls.Add(this.chkChitiet);
            this.groupBox1.Controls.Add(this.chkkhongtinhchenhlech);
            this.groupBox1.Controls.Add(this.chkTheokhoa);
            this.groupBox1.Location = new System.Drawing.Point(18, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(323, 348);
            this.groupBox1.TabIndex = 92;
            this.groupBox1.TabStop = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkExcel);
            this.groupBox3.Controls.Add(this.lblrefesh);
            this.groupBox3.Controls.Add(this.groupBox2);
            this.groupBox3.Controls.Add(this.cmbmaubaocao);
            this.groupBox3.Controls.Add(this.txtmabn);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.chkxuatkhoa);
            this.groupBox3.Controls.Add(this.chklaythuocglivec);
            this.groupBox3.Location = new System.Drawing.Point(2, 178);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(317, 167);
            this.groupBox3.TabIndex = 241;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "BHYT";
            // 
            // chkExcel
            // 
            this.chkExcel.AutoSize = true;
            this.chkExcel.Location = new System.Drawing.Point(210, 73);
            this.chkExcel.Name = "chkExcel";
            this.chkExcel.Size = new System.Drawing.Size(52, 17);
            this.chkExcel.TabIndex = 242;
            this.chkExcel.Text = "Excel";
            this.chkExcel.UseVisualStyleBackColor = true;
            // 
            // lblrefesh
            // 
            this.lblrefesh.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblrefesh.Location = new System.Drawing.Point(6, 138);
            this.lblrefesh.Name = "lblrefesh";
            this.lblrefesh.Size = new System.Drawing.Size(307, 23);
            this.lblrefesh.TabIndex = 241;
            this.lblrefesh.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdb_noitru);
            this.groupBox2.Controls.Add(this.rdb_ngoaitru);
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(161, 44);
            this.groupBox2.TabIndex = 234;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Loại bệnh nhân";
            // 
            // rdb_noitru
            // 
            this.rdb_noitru.Location = new System.Drawing.Point(8, 16);
            this.rdb_noitru.Name = "rdb_noitru";
            this.rdb_noitru.Size = new System.Drawing.Size(72, 24);
            this.rdb_noitru.TabIndex = 0;
            this.rdb_noitru.Text = "Nội trú";
            this.rdb_noitru.CheckedChanged += new System.EventHandler(this.rdb_ngoaitru_CheckedChanged);
            // 
            // rdb_ngoaitru
            // 
            this.rdb_ngoaitru.Location = new System.Drawing.Point(80, 16);
            this.rdb_ngoaitru.Name = "rdb_ngoaitru";
            this.rdb_ngoaitru.Size = new System.Drawing.Size(72, 24);
            this.rdb_ngoaitru.TabIndex = 0;
            this.rdb_ngoaitru.Text = "Ngoại trú";
            this.rdb_ngoaitru.CheckedChanged += new System.EventHandler(this.rdb_ngoaitru_CheckedChanged);
            // 
            // cmbmaubaocao
            // 
            this.cmbmaubaocao.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbmaubaocao.Location = new System.Drawing.Point(85, 114);
            this.cmbmaubaocao.Name = "cmbmaubaocao";
            this.cmbmaubaocao.Size = new System.Drawing.Size(232, 21);
            this.cmbmaubaocao.TabIndex = 240;
            // 
            // txtmabn
            // 
            this.txtmabn.Location = new System.Drawing.Point(212, 35);
            this.txtmabn.Name = "txtmabn";
            this.txtmabn.Size = new System.Drawing.Size(96, 20);
            this.txtmabn.TabIndex = 236;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(3, 117);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 16);
            this.label7.TabIndex = 239;
            this.label7.Text = "Mẫu báo cáo:";
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(172, 35);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 24);
            this.label8.TabIndex = 235;
            this.label8.Text = "Mã BN:";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkxuatkhoa
            // 
            this.chkxuatkhoa.Location = new System.Drawing.Point(6, 69);
            this.chkxuatkhoa.Name = "chkxuatkhoa";
            this.chkxuatkhoa.Size = new System.Drawing.Size(192, 24);
            this.chkxuatkhoa.TabIndex = 237;
            this.chkxuatkhoa.Text = "Lấy số liệu theo ngày xuất viện";
            // 
            // chklaythuocglivec
            // 
            this.chklaythuocglivec.Location = new System.Drawing.Point(6, 93);
            this.chklaythuocglivec.Name = "chklaythuocglivec";
            this.chklaythuocglivec.Size = new System.Drawing.Size(192, 24);
            this.chklaythuocglivec.TabIndex = 238;
            this.chklaythuocglivec.Text = "Lấy số liệu thuốc glivec";
            // 
            // rd3
            // 
            this.rd3.Checked = true;
            this.rd3.Location = new System.Drawing.Point(6, 19);
            this.rd3.Name = "rd3";
            this.rd3.Size = new System.Drawing.Size(104, 27);
            this.rd3.TabIndex = 80;
            this.rd3.TabStop = true;
            this.rd3.Text = "Theo biên lai";
            this.rd3.CheckedChanged += new System.EventHandler(this.rd3_CheckedChanged);
            this.rd3.Click += new System.EventHandler(this.rd3_Click);
            // 
            // rd4
            // 
            this.rd4.Location = new System.Drawing.Point(113, 19);
            this.rd4.Name = "rd4";
            this.rd4.Size = new System.Drawing.Size(83, 27);
            this.rd4.TabIndex = 81;
            this.rd4.Text = "Theo ngày";
            this.rd4.CheckedChanged += new System.EventHandler(this.rd4_CheckedChanged);
            this.rd4.Click += new System.EventHandler(this.rd4_Click);
            // 
            // rd5
            // 
            this.rd5.Location = new System.Drawing.Point(212, 19);
            this.rd5.Name = "rd5";
            this.rd5.Size = new System.Drawing.Size(83, 27);
            this.rd5.TabIndex = 82;
            this.rd5.Text = "Theo khoa";
            this.rd5.CheckedChanged += new System.EventHandler(this.rd3_CheckedChanged);
            this.rd5.Click += new System.EventHandler(this.rd5_Click);
            // 
            // chkChitiet
            // 
            this.chkChitiet.Location = new System.Drawing.Point(6, 52);
            this.chkChitiet.Name = "chkChitiet";
            this.chkChitiet.Size = new System.Drawing.Size(146, 20);
            this.chkChitiet.TabIndex = 86;
            this.chkChitiet.Text = "Chi tiết theo thực thu";
            // 
            // chkkhongtinhchenhlech
            // 
            this.chkkhongtinhchenhlech.Checked = true;
            this.chkkhongtinhchenhlech.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkkhongtinhchenhlech.Location = new System.Drawing.Point(6, 93);
            this.chkkhongtinhchenhlech.Name = "chkkhongtinhchenhlech";
            this.chkkhongtinhchenhlech.Size = new System.Drawing.Size(144, 25);
            this.chkkhongtinhchenhlech.TabIndex = 88;
            this.chkkhongtinhchenhlech.Text = "Không tính chênh lệch";
            this.chkkhongtinhchenhlech.Visible = false;
            // 
            // chkTheokhoa
            // 
            this.chkTheokhoa.Location = new System.Drawing.Point(6, 73);
            this.chkTheokhoa.Name = "chkTheokhoa";
            this.chkTheokhoa.Size = new System.Drawing.Size(144, 23);
            this.chkTheokhoa.TabIndex = 87;
            this.chkTheokhoa.Text = "Tách theo khoa điều trị";
            // 
            // chkAna
            // 
            this.chkAna.AutoSize = true;
            this.chkAna.Checked = true;
            this.chkAna.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAna.Location = new System.Drawing.Point(6, 8);
            this.chkAna.Name = "chkAna";
            this.chkAna.Size = new System.Drawing.Size(130, 17);
            this.chkAna.TabIndex = 23;
            this.chkAna.Text = "Xuất dữ liệu sang Ana";
            this.chkAna.UseVisualStyleBackColor = true;
            // 
            // tree_Loaibn
            // 
            this.tree_Loaibn.CheckBoxes = true;
            this.tree_Loaibn.ForeColor = System.Drawing.Color.DimGray;
            this.tree_Loaibn.FullRowSelect = true;
            this.tree_Loaibn.Location = new System.Drawing.Point(347, 293);
            this.tree_Loaibn.Name = "tree_Loaibn";
            this.tree_Loaibn.ShowLines = false;
            this.tree_Loaibn.ShowPlusMinus = false;
            this.tree_Loaibn.ShowRootLines = false;
            this.tree_Loaibn.Size = new System.Drawing.Size(259, 86);
            this.tree_Loaibn.Sorted = true;
            this.tree_Loaibn.TabIndex = 91;
            // 
            // chkLoaibn
            // 
            this.chkLoaibn.Location = new System.Drawing.Point(347, 272);
            this.chkLoaibn.Name = "chkLoaibn";
            this.chkLoaibn.Size = new System.Drawing.Size(114, 17);
            this.chkLoaibn.TabIndex = 90;
            this.chkLoaibn.TabStop = false;
            this.chkLoaibn.Text = "Loại bệnh nhân";
            this.chkLoaibn.CheckedChanged += new System.EventHandler(this.chkLoaibn_CheckedChanged);
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Location = new System.Drawing.Point(348, 7);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(57, 17);
            this.chkAll.TabIndex = 89;
            this.chkAll.Text = "Tất cả";
            this.chkAll.UseVisualStyleBackColor = true;
            this.chkAll.CheckedChanged += new System.EventHandler(this.chkAll_CheckedChanged);
            // 
            // tree_Field
            // 
            this.tree_Field.BackColor = System.Drawing.Color.White;
            this.tree_Field.CheckBoxes = true;
            this.tree_Field.ForeColor = System.Drawing.Color.DimGray;
            this.tree_Field.FullRowSelect = true;
            this.tree_Field.Location = new System.Drawing.Point(612, 28);
            this.tree_Field.Name = "tree_Field";
            this.tree_Field.ShowLines = false;
            this.tree_Field.ShowPlusMinus = false;
            this.tree_Field.ShowRootLines = false;
            this.tree_Field.Size = new System.Drawing.Size(222, 351);
            this.tree_Field.TabIndex = 25;
            // 
            // tree_Loai
            // 
            this.tree_Loai.BackColor = System.Drawing.Color.White;
            this.tree_Loai.CheckBoxes = true;
            this.tree_Loai.ForeColor = System.Drawing.Color.DimGray;
            this.tree_Loai.FullRowSelect = true;
            this.tree_Loai.Location = new System.Drawing.Point(347, 28);
            this.tree_Loai.Name = "tree_Loai";
            this.tree_Loai.ShowLines = false;
            this.tree_Loai.ShowPlusMinus = false;
            this.tree_Loai.ShowRootLines = false;
            this.tree_Loai.Size = new System.Drawing.Size(259, 236);
            this.tree_Loai.TabIndex = 24;
            // 
            // chkAll1
            // 
            this.chkAll1.Location = new System.Drawing.Point(612, 6);
            this.chkAll1.Name = "chkAll1";
            this.chkAll1.Size = new System.Drawing.Size(117, 16);
            this.chkAll1.TabIndex = 85;
            this.chkAll1.Text = "Thông tin hiển thị";
            this.chkAll1.CheckedChanged += new System.EventHandler(this.chkAll1_CheckedChanged);
            // 
            // rd2
            // 
            this.rd2.Location = new System.Drawing.Point(511, 7);
            this.rd2.Name = "rd2";
            this.rd2.Size = new System.Drawing.Size(92, 16);
            this.rd2.TabIndex = 84;
            this.rd2.Text = "Loại viện phí";
            this.rd2.CheckedChanged += new System.EventHandler(this.rd2_CheckedChanged);
            // 
            // rd1
            // 
            this.rd1.Checked = true;
            this.rd1.Location = new System.Drawing.Point(412, 7);
            this.rd1.Name = "rd1";
            this.rd1.Size = new System.Drawing.Size(97, 16);
            this.rd1.TabIndex = 83;
            this.rd1.TabStop = true;
            this.rd1.Text = "Nhóm viện phí";
            this.rd1.CheckedChanged += new System.EventHandler(this.rd1_CheckedChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(843, 385);
            this.tabPage3.TabIndex = 4;
            this.tabPage3.Text = "Cấu hình HDDT";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(843, 385);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(7, 7);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(827, 375);
            this.dataGridView1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataGridView2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(843, 385);
            this.tabPage2.TabIndex = 3;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(7, 7);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(830, 372);
            this.dataGridView2.TabIndex = 0;
            // 
            // tab_hddt
            // 
            this.tab_hddt.Controls.Add(this.tim);
            this.tab_hddt.Controls.Add(this.dtg);
            this.tab_hddt.Location = new System.Drawing.Point(4, 22);
            this.tab_hddt.Name = "tab_hddt";
            this.tab_hddt.Padding = new System.Windows.Forms.Padding(3);
            this.tab_hddt.Size = new System.Drawing.Size(843, 385);
            this.tab_hddt.TabIndex = 0;
            this.tab_hddt.Text = "Hóa đơn điện tử";
            this.tab_hddt.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(799, 476);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(63, 33);
            this.progressBar1.TabIndex = 22;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // bKetxuat
            // 
            this.bKetxuat.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bKetxuat.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.bKetxuat.Location = new System.Drawing.Point(450, 6);
            this.bKetxuat.Name = "bKetxuat";
            this.bKetxuat.Size = new System.Drawing.Size(99, 33);
            this.bKetxuat.TabIndex = 23;
            this.bKetxuat.Text = "Kết xuất";
            this.bKetxuat.UseVisualStyleBackColor = true;
            this.bKetxuat.Visible = false;
            this.bKetxuat.Click += new System.EventHandler(this.bKetxuat_Click);
            // 
            // bHDDTStart
            // 
            this.bHDDTStart.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bHDDTStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.bHDDTStart.Location = new System.Drawing.Point(509, 476);
            this.bHDDTStart.Name = "bHDDTStart";
            this.bHDDTStart.Size = new System.Drawing.Size(110, 33);
            this.bHDDTStart.TabIndex = 1;
            this.bHDDTStart.Text = "Medi - HDDT";
            this.bHDDTStart.UseVisualStyleBackColor = true;
            this.bHDDTStart.Click += new System.EventHandler(this.bHDDTStart_Click);
            // 
            // frmMain
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(878, 513);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.TabHDDT);
            this.Controls.Add(this.gGetData);
            this.Controls.Add(this.record);
            this.Controls.Add(this.bPause);
            this.Controls.Add(this.bClose);
            this.Controls.Add(this.bHDDTStart);
            this.Controls.Add(this.bStart);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Hoá đơn điện tử";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nbPhut)).EndInit();
            this.gGetData.ResumeLayout(false);
            this.gGetData.PerformLayout();
            this.TabHDDT.ResumeLayout(false);
            this.tab_hsba.ResumeLayout(false);
            this.tab_hsba.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tab_hddt.ResumeLayout(false);
            this.tab_hddt.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private void frmMain_Load(object sender, EventArgs e)
        {

            //this.Text = this.Text + " - Version: " + ProductVersion;
            ipApi = txtIP.Text;
            user = m.user;
            s_giavp_dmbd = "SELECT * FROM (SELECT giavp.id AS mavp, translate (giavp.ma using nchar_cs) AS ma, giavp.ten, giavp.dvt, nhomvp.ma as manhomdv,nhomvp.ten AS nhomdv,'DVKT' AS loaidv";
            s_giavp_dmbd += " FROM v_giavp giavp INNER JOIN v_loaivp loaivp ON giavp.id_loai = loaivp.id INNER JOIN v_nhomvp nhomvp ON loaivp.id_nhom = nhomvp.ma WHERE giavp.hide = 0";
            s_giavp_dmbd += " UNION ALL SELECT id AS mavp, translate (d_dmbd.ma using nchar_cs) AS ma, d_dmbd.ten, d_dmbd.dang AS dvt, 0 as manhomdv,translate ('Thuốc - VTYT' using nchar_cs) AS nhom,'THUOC' AS loaidv FROM d_dmbd d_dmbd)";

            s_hanhchinh = "SELECT * FROM (SELECT b.mabn, b.hoten, b.namsinh, b.phai, (b.sonha || ' ' || b.thon ||  ' ' || c.tenpxa || ' ' || d.tenquan || ' ' || e.tentt) AS diachi";
            s_hanhchinh += " FROM btdbn b INNER JOIN btdpxa c ON c.maphuongxa = b.maphuongxa INNER JOIN btdquan d ON d.maqu = b.maqu INNER JOIN btdtt e ON e.matt = b.matt)";
            m = new AccessDataApi(ipApi);
            //cbMMYY.DataSource = m.get_data("select mmyy from m_table order by mmyy").Tables[0];
            //cbMMYY.DisplayMember = "mmyy";
            //cbMMYY.ValueMember = "mmyy";
            //
            cbMMYY.Text = denngay.Text.Substring(3, 2) + denngay.Text.Substring(8, 2);
            //cbLoaiVP.SelectedIndex = 0;
            intAna();
            //load_grid();
        }
        private void intAna()
        {
            
            chkkhongtinhchenhlech.Checked = true;
            try
            {
                m_giavpbangdongiacongvattu = m.s_giavbangdongiacongvattu;
                f_Load_Tyle();
                f_Load_Tree_Loai();
                f_Load_Tree_Field();
                f_Load_Tree_Loaibn();
            }
            catch {}
            
            f_Load_OptionTree_Field();//12/10/2013
            f_Load_OptionTree_Loai();//21/10/2013
            rd3.Checked = true;
            chkAll.Checked = true;
            chkAll_CheckedChanged(null, null);

            chkAll1.Checked = true;
            chkAll1_CheckedChanged(null, null);

            chkLoaibn.Checked = true;
            chkLoaibn_CheckedChanged(null, null);

        }
        private void f_Load_Tyle()
        {
            try
            {
               s_bc1 = m.get_data("select mp from v_mpxxx where id=8 and stt=1").Tables[0].Rows[0]["mp"].ToString();
               s_bc2 = m.get_data("select mp from v_mpxxx where id=8 and stt=2").Tables[0].Rows[0]["mp"].ToString();
            }
            catch
            {
                s_bc1 = "100";
                s_bc2 = "100";
            }
        }
        private void f_Load_Tree_Field()
        {
            try
            {
                m_ds1 = new DataSet();
                m_ds1.Tables.Add("Table");
                m_ds1.Tables[0].Columns.Add("MA");
                m_ds1.Tables[0].Columns.Add("TEN");
                if (rd3.Checked)
                {
                    m_ds1.Tables[0].Rows.Add(new string[] { "MABN", "Mã số BN" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "HOTEN", "Họ và tên bệnh nhân" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "NAMSINH", "Năm sinh" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "NGAY", "Ngày thu" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "TENKP", "Tên kp" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "DOITUONG", "Đối tượng" });//28102013-bv Huyet Hoc ha noi.
                    m_ds1.Tables[0].Rows.Add(new string[] { "QUYENSO", "Quyển sổ" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOBIENLAI", "Số biên lai" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOCHUNGTU", "Số chứng từ" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOTIEN", "Tổng số tiền" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "BHYT", "BHYT trả" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "chenhlechdv", "Dịch vụ" });//ThanhCuong-18062011
                    m_ds1.Tables[0].Rows.Add(new string[] { "TAMUNG", "Tạm ứng" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "TONGTAMUNG", "Tổng tạm ứng" });//03102013-bv Huyet Hoc ha noi.
                    m_ds1.Tables[0].Rows.Add(new string[] { "TONGHOANTAMUNG", "Tổng hoàn ứng" });//11102013-bv Huyet Hoc ha noi.					
                    m_ds1.Tables[0].Rows.Add(new string[] { "MIEN", "Miễn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCTHU", "Thực thu" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCCHI", "Thực chi" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "LYDOMIEN", "Lý do miễn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "NGUOIKYMIEN", "Nhân viên nhập miễn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "GHICHU", "Ghi chú miễn" });
                }
                else
                if (rd4.Checked)
                {
                    m_ds1.Tables[0].Rows.Add(new string[] { "NGAY", "Ngày thu" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOHOADON", "Tổng hoá đơn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOTIEN", "Tổng số tiền" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "BHYT", "Tổng BHYT" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "TAMUNG", "Tổng tạm ứng" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "MIEN", "Tổng miễn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCTHU", "Thực thu" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCCHI", "Thực chi" });
                }
                else
                    if (rd5.Checked)
                {
                    m_ds1.Tables[0].Rows.Add(new string[] { "NGAY", "Mã KP" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "VIETTAT", "Khoa" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "TENKP", "Tên khoa" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOHOADON", "Tổng hoá đơn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "SOTIEN", "Tổng số tiền" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "BHYT", "Tổng BHYT" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "TAMUNG", "Tổng tạm ứng" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "MIEN", "Tổng miễn" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCTHU", "Thực thu" });
                    m_ds1.Tables[0].Rows.Add(new string[] { "THUCCHI", "Thực chi" });
                }
                f_Load_Tree(tree_Field, m_ds1);
                f_Set_CheckID(tree_Field, chkAll1.Checked);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void f_Load_Tree_Loai()
        {
            try
            {
                string asql = "select to_char(ma) ma, ten from v_nhomvp order by stt asc";
                if (rd2.Checked)
                {
                    asql = "select to_char(id) ma, ten from v_loaivp order by stt asc";
                }
                m_ds = m.get_data(asql);
                DataRow r = m_ds.Tables[0].NewRow();
                r["ma"] = "0";
                r["ten"] = "Thuốc khoa dược";
                m_ds.Tables[0].Rows.Add(r);
                if (rd1.Checked)
                {
                    r = m_ds.Tables[0].NewRow();
                    r["ten"] = "Giường dịch vụ";
                    string sq = "select a.ma,b.idnhombhytmedisoft idmedi from v_nhomvp a"
                        + " inner join v_nhombhyt b on a.idnhombhyt=b.id"
                        + " where b.idnhombhytmedisoft=11";//tien giuong
                    DataTable dttam = m.get_data(sq).Tables[0];
                    for (int i = 0; i < m_ds.Tables[0].Rows.Count; i++)
                    {
                        if (m_ds.Tables[0].Rows[i]["ma"].ToString()
                            == dttam.Rows[0]["ma"].ToString())
                        {
                            r["ma"] = m_ds.Tables[0].Rows[i]["ma"].ToString() + "101";
                            m_ds.Tables[0].Rows.InsertAt(r, i + 1);
                            break;
                        }
                    }
                }
                f_Load_Tree(tree_Loai, m_ds);
                f_Set_CheckID(tree_Field, chkAll.Checked);
            }
            catch
            {
            }
        }
        private void f_Load_Tree_Loaibn()
        {
            try
            {
                string asql = "select to_char(id) ma, ten from v_loaibn order by id";
                f_Load_Tree(tree_Loaibn, m.get_data(asql));
                f_Set_CheckID(tree_Loaibn, chkLoaibn.Checked);
            }
            catch
            {
            }
        }
        private void f_Load_Tree(TreeView v_tree, DataSet v_ds)
        {
            try
            {
                TreeNode anode;
                v_tree.Nodes.Clear();
                v_tree.Sorted = false;
                for (int i = 0; i < v_ds.Tables[0].Rows.Count; i++)
                {
                    anode = new TreeNode(v_ds.Tables[0].Rows[i]["ten"].ToString());
                    anode.Tag = v_ds.Tables[0].Rows[i]["ma"].ToString();
                    v_tree.Nodes.Add(anode);
                }
            }
            catch
            {
            }
        }
        private void f_Set_CheckID(TreeView v_tree, bool v_b)
        {
            try
            {
                for (int i = 0; i < v_tree.Nodes.Count; i++)
                {
                    v_tree.Nodes[i].Checked = v_b;
                }
            }
            catch
            {
            }
        }
        private string f_Get_CheckID(TreeView v_tree)
        {
            try
            {
                string r = "";
                for (int i = 0; i < v_tree.Nodes.Count; i++)
                {
                    if (v_tree.Nodes[i].Checked) r = r.Trim().Trim(',') + "," + v_tree.Nodes[i].Tag.ToString();
                }
                r = r.Trim().Trim(',');
                //MessageBox.Show(r);
                return r;
            }
            catch
            {
                return "";
            }
        }
        void f_Load_OptionTree_Field()
        {
            try
            {
                for (int i = 0; i < tree_Field.Nodes.Count; i++)
                {
                    if (tree_Field.Nodes[i].Tag.ToString() == "tongtamung".ToUpper())
                    {
                        tree_Field.Nodes[i].Checked = optionBaocaoTTRV_LayTongTamUng;
                    }
                    else if (tree_Field.Nodes[i].Tag.ToString() == "tonghoantamung".ToUpper())
                    {
                        tree_Field.Nodes[i].Checked = optionBaocaoTTRV_LayTongTamUng_HoanTra;
                    }
                    else if (tree_Field.Nodes[i].Tag.ToString() == "doituong".ToUpper())
                    {
                        tree_Field.Nodes[i].Checked = optionBaocaoTTRV_LayCotDoiTuongBN;
                    }
                    else continue;
                }
            }
            catch { }
        }
        bool optionBaocaoTTRV_LayTongTamUng
        {

            get
            {
                try
                {
                    return f_get_optionBaocaoTTRV(2) == "1";
                }
                catch { return false; }
            }
            set
            {
                f_set_optionBaocaoTTRV(2, "Lấy thêm cột tổng tiền tạm ứng.", (value) ? "1" : "0");
            }
        }
        bool optionBaocaoTTRV_LayTongTamUng_HoanTra
        {

            get
            {
                try
                {
                    return f_get_optionBaocaoTTRV(3) == "1";
                }
                catch { return false; }
            }
            set
            {
                f_set_optionBaocaoTTRV(3, "Lấy thêm cột tổng tiền tạm ứng đã hoàn trả.", (value) ? "1" : "0");
            }
        }
        bool optionBaocaoTTRV_LayCotDoiTuongBN
        {

            get
            {
                try
                {
                    return f_get_optionBaocaoTTRV(5) == "1";
                }
                catch { return false; }
            }
            set
            {
                f_set_optionBaocaoTTRV(5, "Lấy thêm cột đối tượng bệnh nhân.", (value) ? "1" : "0");
            }
        }
        bool optionBaocaoTTRV_LayCotGiuongDichVu_NhomVP
        {

            get
            {
                try
                {
                    return f_get_optionBaocaoTTRV(4) == "1";
                }
                catch { return false; }
            }
            set
            {
                f_set_optionBaocaoTTRV(4, "Lấy thêm cột giường dịch vụ.", (value) ? "1" : "0");
            }
        }
        void f_set_optionBaocaoTTRV(int id, string ten, string giatri)
        {
            string fileName = "..\\..\\..\\xml\\optionBaocaoTTRV.xml";
            DataSet dstam = new DataSet();
            try
            {
                dstam.ReadXml(fileName, XmlReadMode.ReadSchema);
                string test = "";
                test = dstam.Tables[0].Rows[0]["id"].ToString();
                test = dstam.Tables[0].Rows[0]["giatri"].ToString();
                test = dstam.Tables[0].Rows[0]["ten"].ToString();
            }
            catch
            {
                dstam = new DataSet();
                dstam.Tables.Add("option");
                dstam.Tables[0].Columns.Add("id");
                dstam.Tables[0].Columns.Add("giatri");
                dstam.Tables[0].Columns.Add("ten");

            }
            DataRow[] arrdr = dstam.Tables[0].Select("id=" + id.ToString());

            if (arrdr.Length > 0)
            {
                arrdr[0]["giatri"] = giatri;
                arrdr[0]["ten"] = ten;
            }
            else
            {
                dstam.Tables[0].Rows.Add(new object[] { id, giatri, ten });
            }
            dstam.WriteXml(fileName, XmlWriteMode.WriteSchema);

        }
        string f_get_optionBaocaoTTRV(int id)
        {
            string sgiatri = "0";
            try
            {
                string fileName = "..\\..\\..\\xml\\optionBaocaoTTRV.xml";
                DataSet dstam = new DataSet();
                dstam.ReadXml(fileName, XmlReadMode.ReadSchema);
                DataRow[] arrdr = dstam.Tables[0].Select("id=" + id.ToString());
                sgiatri = arrdr[0]["giatri"].ToString();
            }
            catch
            {
            }
            return sgiatri;

        }
        void f_Load_OptionTree_Loai()
        {
            try
            {
                for (int i = 0; i < tree_Loai.Nodes.Count; i++)
                {
                    if (tree_Loai.Nodes[i].Tag.ToString() == "8101".ToUpper())
                    {
                        tree_Loai.Nodes[i].Checked = optionBaocaoTTRV_LayCotGiuongDichVu_NhomVP;
                    }
                    else continue;
                }
            }
            catch { }
        }

        private void f_KetxuatAna()
        {
            try
            {

                string hddtID = "";
                try
                {
                    dsHDDT = hddt.dataSetFromSql("mysql", "select sochungtu,sohoadon,tencn,mach_cn from tb_master where tinhtranghoadon=1 and ngaylap BETWEEN str_to_date('" + tungay.Text + " 00:00:00','%d/%m/%Y %H:%i:%s') AND str_to_date('" + denngay.Text + " 23:59:59','%d/%m/%Y %H:%i:%s')");
                }
                catch { }
                
                if (dsHDDT.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dHDDT in dsHDDT.Tables[0].Rows)
                    {
                        dsAna = hddt.dataSetFromSql("mssql", "select sochungtu from  BANGDULIEUVIENPHI_TAM where sochungtu='" + dHDDT["sochungtu"].ToString() + "'"); //Lọc những Sochungtu chưa có trong Ana
                        if (dsAna.Tables[0].Rows.Count == 0)
                            hddtID += "'" + dHDDT["sochungtu"].ToString() + "',";
                    }
                    hddtID = hddtID.TrimEnd(',');
                    if (hddtID.Length == 0) //Không tồn tại hóa đơn nào đã có số HĐ điện tử nhưng ko có trong Ana
                        return;
                }
                else //Không có hóa đơn của Medi nào thì return
                    return;
                string asql = "", asql1 = "", aexp = "";
                string atmp = "";
                string aloaibn = f_Get_CheckID(tree_Loaibn);
                string asqlht = "select * from v_hoantra where to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy')";

                DataSet ads = new DataSet();
                string atable = "", s_tyle = "";
                atable = "h.ma";
                s_tyle = s_bc1;
                m_ds2 = m_ds1.Clone();
                if (tungay.Text.Substring(3, 7) != "09/2008") s_tyle = "100";//Benh vien my phuoc 
                for (int i = 0; i < tree_Field.Nodes.Count; i++)
                {
                    if (tree_Field.Nodes[i].Checked)
                    {
                        try
                        {
                            m_ds2.Tables[0].Rows.Add(new string[] { tree_Field.Nodes[i].Tag.ToString().Trim(), tree_Field.Nodes[i].Text.Trim() });
                        }
                        catch
                        {
                        }
                    }
                }
                //

                bool ok = false;
                for (int i = 0; i < tree_Loai.Nodes.Count; i++)
                {
                    if (tree_Loai.Nodes[i].Checked)
                    {
                        ok = true;
                        break;
                    }
                }

                atmp = "";
                string athucthu = "";//24/09/2013 chỉ lấy số tiền thực thu của BN.
                string aphongkham = "";//14/10/2013 lấy số liệu của BN phòng khám.
                if (ok)
                {
                    for (int i = 0; i < tree_Loai.Nodes.Count; i++)
                    {
                        if (tree_Loai.Nodes[i].Checked)
                        {
                            try
                            {
                                atmp = atmp + ",sum(decode(" + atable + ","
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString()
                                    + ",b.soluong*b.dongia "
                                    + "+ b.soluong*b.dongia*nvl(b.vat,0)/100,0))*"
                                    + s_tyle + "/100 LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                athucthu = athucthu + ",sum(decode(" + atable + ","
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString()
                                    + ",b.sotien-b.bhyttra-b.mien "
                                    + "+ b.soluong*b.dongia*nvl(b.vat,0)/100,0))"
                                    + " LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                aphongkham = aphongkham + ",sum(decode(" + atable + ",'"
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString().Trim()
                                    + "',b.soluong*b.dongia-b.thieu,0))*"
                                    + s_tyle + "/100 LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                m_ds2.Tables[0].Rows.Add(new string[] { "LLLL" + tree_Loai.Nodes[i].Tag.ToString(), tree_Loai.Nodes[i].Text.Trim() });
                            }
                            catch (Exception exx)
                            {
                                MessageBox.Show(exx.ToString());
                            }
                        }
                    }
                    atmp = atmp.Trim().Trim(',').Trim();
                    if (chkChitiet.Checked)
                    {
                        //lay cau truy van lay so tien thuc thu theo nhom(hoac loai) VP.
                        athucthu = athucthu.Trim().Trim(',').Trim();
                        atmp = athucthu;
                    }
                    if (atmp.Length > 0)
                    {
                        atmp = "," + atmp.Trim(',');
                    }
                }

                string asqldmbd = "select a1.id id, c1.id_loai id_loai from d_dmbd a1, d_dmnhom b1, (select a0.ma, min(nvl(b0.id,0)) id_loai from v_nhomvp a0, v_loaivp b0 where a0.ma=b0.id_nhom(+) group by a0.ma) c1 where a1.manhom=b1.id(+) and b1.nhomvp=c1.ma(+)";

                if (rd5.Checked)//Nhom theo khoa
                {
                    //Nếu chkTheokhoa có chọn, sẽ in theo khoa và tách theo tung khoa điều trị.
                    if (chkTheokhoa.Checked == true)
                    {
                        asql = "select g.makp ngay, g.tenkp, g.viettat, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0) thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and b.makp=g.makp(+) " + aexp + " group by g.makp,g.tenkp,g.viettat order by g.tenkp asc";
                        asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, g.makp ngay,g.tenkp,g.viettat, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and b.makp=g.makp(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')', g.makp,g.tenkp,g.viettat ";
                        asql1 += "order by d.sohieu asc, a.sobienlai asc, g.tenkp asc";
                    }
                    else
                    {
                        asql = "select g.makp ngay, g.tenkp, g.viettat, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0) thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.makp=g.makp(+) " + aexp + " group by g.makp,g.tenkp,g.viettat order by g.tenkp asc";
                        asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, g.makp ngay,g.tenkp,g.viettat, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and a.makp=g.makp(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')', g.makp,g.tenkp,g.viettat ";
                        asql1 += "order by d.sohieu asc, a.sobienlai asc, g.tenkp asc";
                    }
                }
                else if (rd4.Checked)//Nhom theo ngay
                {
                    asql = "select to_char(a.ngay,'dd/mm/yyyy') ngay, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null " + aexp + " group by to_char(a.ngay,'dd/mm/yyyy') order by to_char(a.ngay,'dd/mm/yyyy') asc";
                    asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)*" + s_tyle + "/100) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, a.ngay, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')' order by d.sohieu asc, a.sobienlai asc, a.ngay asc";
                }
                else if (rd3.Checked)//Nhóm theo biên lai
                {
                    //nếu chkTheokhoa có chọn, sẽ in theo biên lai và theo khoa.
                    if (chkTheokhoa.Checked == true)
                    {
                        asql = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu, ccc.ten lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)*" + s_tyle + "/100) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid,g.makp,g.tenkp " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_lydomien ccc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h,btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.id=c.id(+) and c.maduyet=cc.ma(+) and c.lydo=ccc.id(+) and a.quyenso=d.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and b.makp=g.makp " + aexp;
                        asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, a.ngay, c.ghichu, cc.ten, ccc.ten, e.hoten ||' ('||to_char(e.userid)||')',g.makp,g.tenkp order by d.sohieu asc, a.sobienlai asc, a.ngay asc";
                    }
                    else
                    {
                        #region Query lấy dữ liệu

                        #region bn noi tru + ngoai tru
                        #region select
                        asql = "select a.loaibn,to_char(a.id) id, to_char(a.quyenso) quyensoid"
                            + " , to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay"
                            + " , d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu"
                            + ", c.ghichu, ccc.ten lydomien, cc.ten nguoikymien, aaa.mabn" 
                            + ", aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien"
                            + ", sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt"
                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",i.chenhlechdv")
                            + ", nvl(a.tamung,0) tamung"
                            + ", nvl(case when sum(nvl(b.sotien,0))"
                            + "-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0))>=0 "
                            + " then sum(nvl(b.sotien,0))-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))"
                            + "-sum(nvl(b.mien,0)) end,0)*" + s_tyle + "/100 thucthu"
                            + ", nvl(case when sum(nvl(b.sotien,0))-nvl(a.tamung,0)"
                            + "-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0))<0 then (sum(nvl(b.sotien,0))"
                            + "-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0)))*-1 end,0)*"
                            + s_tyle + "/100 thucchi, e.hoten nguoithu"
                            + ",0.0 tongtamung"//03/10/2013
                            + ",0.0 tonghoantamung"//11/10/2013
                            + ",to_char(aa.ngayvao,'dd/mm/yyyy') ngayvao"//12/10/2013
                            + ",to_char(aa.ngayra,'dd/mm/yyyy') ngayra"//12/10/2013
                            + ",to_char(aa.maql) maql,'' doituong"//12/10/2013
                            + ", e.hoten ||' ('||to_char(e.userid)||')' userid,g2.makp"
                            + ",g2.tenkp " + atmp;
                        #endregion
                        #region from
                        asql += " from  v_ttrvds aa"
                            + " inner join v_ttrvll a on aa.id=a.id"
                            + " left join btdbn aaa on aa.mabn=aaa.mabn"
                            + " left join (" + asqlht + ") aht on a.quyenso=aht.quyenso and a.sobienlai=aht.sobienlai"
                            + " left join v_ttrvct b on a.id=b.id"
                            + " left join v_miennoitru c on a.id=c.id"
                            + " left join v_dsduyet cc on c.maduyet=cc.ma"
                            + " left join v_lydomien ccc on c.lydo=ccc.id"
                            + " left join v_quyenso d on a.quyenso=d.id"
                            + " left join v_dlogin e on a.userid=e.id"
                            + " left join (select id, id_loai from v_giavp union all select id"
                                + ",id_loai from (" + asqldmbd + ")) f on b.mavp= f.id"
                            + " left join (select id, id_nhom from v_loaivp union all select 0 id"
                                + ",0 id_nhom from dual) g on f.id_loai=g.id"
                            + " left join (select ma from v_nhomvp union all "
                                + "select 0 from dual) h on g.id_nhom=h.ma"
                            + " inner join btdkp_bv g2 on a.makp=g2.makp"

                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",V_TTRVCT_chenhlech i ")

                        #endregion
                        #region where
                            + " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                            + " >=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                            + " and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                            + " <=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                            + " and aht.id is null "
                            + ((chkkhongtinhchenhlech.Checked) ? "" : " and a.id=i.id and a.makp=i.makp  ")
                            + aexp;
                        if (hddtID.Length >0)
                            asql += "  and lower(d.sohieu) ||' - '|| to_char(a.sobienlai ) in (" + hddtID + ")";
                        #endregion
                        #region group by
                        asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn"
                            + ", aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt"
                            + ", a.mien, e.hoten, a.ngay "
                            + ",c.ghichu, cc.ten, ccc.ten" //Toản đóng kiểm tra lại table v_miennoitru
                            + ", e.hoten ||' ('||to_char(e.userid)||')',g2.makp,g2.tenkp"
                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",i.chenhlechdv ")
                            + ",aa.ngayvao"
                            + ",aa.ngayra"
                            + ",aa.maql,a.loaibn";
                        #endregion
                        #endregion

                        #region bn phong kham(thu truc tiep)

                        if (aloaibn == "" || aloaibn.IndexOf("3") > -1)
                        {
                            asql += " union all ";
                            asql = asql + " select a.loaibn,to_char(a.id) id"
                                + ", to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai"
                                + ", to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso"
                                + ", d.sohieu||' - '||to_char(a.sobienlai) sochungtu"
                                + ", c.ghichu, ccc.ten lydomien, cc.ten nguoikymien" 
                                + ", a.mabn, a.hoten, a.namsinh"
                                + ", sum(nvl(b.soluong,0)*nvl(b.dongia,0))*" + s_tyle + "/100 sotien"
                                + ", sum(decode(b.madoituong,1,nvl(b.mien,0),0)) bhyt"
                                + ", nvl(c.sotien,0) mien"//, sum(nvl(b.thieu,0)) thieu"  //Toản đóng xem lại table v_mienngtru
                                + ", 0 as mien"
                                + ((chkkhongtinhchenhlech.Checked) ? "" : ",0 chenhlechdv")
                                + ",0 tamung, sum(nvl(b.soluong,0)*nvl(b.dongia,0)*" + s_tyle + "/100"
                                + "-nvl(decode(b.madoituong,1,b.mien,0),0)-nvl(b.thieu,0)) -nvl(c.sotien,0) thucthu"
                                + ",0 thucchi, e.hoten nguoithu"
                                + ",0.0 tongtamung"
                                + ",0.0 tonghoantamung"
                                + ",to_char(a.ngay,'dd/mm/yyyy') ngayvao"
                                + ",to_char(a.ngay,'dd/mm/yyyy') ngayra"
                                + ",to_char(a.maql) maql,'' doituong "
                                + ", e.hoten ||' ('||to_char(e.userid)||')' userid "
                                + ",i.makp,i.tenkp " + aphongkham

                                + " from v_vienphill a"
                                + " left join v_vienphict b on a.id=b.id"
                                + " left join (" + asqlht + ") aa on a.quyenso=aa.quyenso and a.sobienlai=aa.sobienlai"
                                + " left join v_mienngtru c on a.id=c.id"
                                + " left join v_dsduyet cc on c.maduyet=cc.ma "
                                + " left join v_lydomien ccc on c.lydo=ccc.id "
                                + " left join v_quyenso d on a.quyenso=d.id"
                                + " left join v_dlogin e on a.userid=e.id"
                                + " left join (select id, id_loai from v_giavp union all "
                                + "select id, nvl(id_loai,0) id_loai from (" + asqldmbd + "))"
                                + " f on b.mavp= f.id"
                                + " left join (select id, id_nhom from v_loaivp union all "
                                + " select 0 id, 0 id_nhom from dual) g on f.id_loai=g.id"
                                + " left join (select ma from v_nhomvp "
                                + " union all select 0 from dual) h on g.id_nhom=h.ma"
                                + " left join btdkp_bv i on b.makp=i.makp"


                                + " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                                + " >=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                                + " and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                                + " <=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                                + " and aa.id is null "
                                + aexp;
                            if (hddtID.Length > 0)
                                asql += "  and lower(d.sohieu) ||' - '|| to_char(a.sobienlai ) in (" + hddtID + ")";
                            asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu"
                            + ", a.mabn, a.hoten, a.namsinh, e.hoten"
                            + ",a.ngay,e.hoten ||' ('||to_char(e.userid)||')'"
                            + ",a.maql,a.ngay,i.tenkp,i.makp,a.loaibn"
                            ;
                        }
                        #endregion

                        asql = "select * from (" + asql + ") "
                            + " order by quyenso asc, sobienlai asc, ngay asc"
                            ;
                        #endregion
                    }
                }

                if (m_ds2.Tables[0].Rows.Count <= 0)
                {
                    progressBar1.Value = progressBar1.Maximum;
                    timer1.Enabled = false;
                    progressBar1.Value = 0;
                    MessageBox.Show(this, "Chọn thông tin báo cáo cần hiển thị", "Thông báo", MessageBoxButtons.OK);
                    return;
                }
                ads = m.get_data_all_vp(m.StringToDate(tungay.Text), m.StringToDate(denngay.Text), asql);
                if (rd4.Checked || rd5.Checked)
                {
                    #region
                    DataSet ads1 = new DataSet();
                    ads1 = m.get_data_mmyy(asql1, tungay.Text, denngay.Text);
                    for (int i = 0; i < ads.Tables[0].Rows.Count; i++)
                    {
                        ads.Tables[0].Rows[i]["sohoadon"] = 0;
                        ads.Tables[0].Rows[i]["sotien"] = 0;
                        ads.Tables[0].Rows[i]["bhyt"] = 0;
                        ads.Tables[0].Rows[i]["tamung"] = 0;
                        ads.Tables[0].Rows[i]["mien"] = 0;
                        ads.Tables[0].Rows[i]["thucthu"] = 0;
                        ads.Tables[0].Rows[i]["thucchi"] = 0;
                        try
                        {
                            foreach (DataRow r1 in ads1.Tables[0].Select("ngay='" + ads.Tables[0].Rows[i]["ngay"].ToString() + "'"))
                            {
                                ads.Tables[0].Rows[i]["sohoadon"] = decimal.Parse(ads.Tables[0].Rows[i]["sohoadon"].ToString()) + 1;
                                ads.Tables[0].Rows[i]["sotien"] = decimal.Parse(ads.Tables[0].Rows[i]["sotien"].ToString()) + decimal.Parse(r1["sotien"].ToString());
                                ads.Tables[0].Rows[i]["bhyt"] = decimal.Parse(ads.Tables[0].Rows[i]["bhyt"].ToString()) + decimal.Parse(r1["bhyt"].ToString());
                                ads.Tables[0].Rows[i]["tamung"] = decimal.Parse(ads.Tables[0].Rows[i]["tamung"].ToString()) + decimal.Parse(r1["tamung"].ToString());
                                ads.Tables[0].Rows[i]["mien"] = decimal.Parse(ads.Tables[0].Rows[i]["mien"].ToString()) + decimal.Parse(r1["mien"].ToString());
                                ads.Tables[0].Rows[i]["thucthu"] = decimal.Parse(ads.Tables[0].Rows[i]["thucthu"].ToString()) + decimal.Parse(r1["thucthu"].ToString());
                                ads.Tables[0].Rows[i]["thucchi"] = decimal.Parse(ads.Tables[0].Rows[i]["thucchi"].ToString()) + decimal.Parse(r1["thucchi"].ToString());
                            }
                        }
                        catch
                        {
                        }
                    }
                    #endregion
                }

                //bv huyet hoc ha noi muon lay tong tien tam ung(bao gom da hoan tra)
                //va tong tien da hoan tam ung.
                string stam = f_Get_CheckID(tree_Field);
                string sfieldtamung = (stam.IndexOf("TONGTAMUNG".ToUpper()) == -1) ? "" : "TONGTAMUNG".ToUpper();
                string sfieldhoantamung = (stam.IndexOf("TONGhoanTAMUNG".ToUpper()) == -1) ? "" : "TONGhoanTAMUNG".ToUpper();
                string sfielddoituong = (stam.IndexOf("doituong".ToUpper()) == -1) ? "" : "doituong".ToUpper();
                stam = f_Get_CheckID(tree_Loai);
                string sfieldtiengiuong = "";
                foreach (string st in stam.Split(','))
                {
                    if (st.IndexOf("101") > -1)//lay ten cot tien giuong dich vu
                    {
                        sfieldtiengiuong = "LLLL" + st;
                        break;
                    }
                }

                if (sfieldtamung != "" || sfieldhoantamung != ""
                    || sfieldtiengiuong != "" || sfielddoituong != "")
                {
                    f_set_tongtamung_tonggiuong(ads.Tables[0]
                        , sfieldhoantamung, sfieldtamung
                        , sfieldtiengiuong.Replace("101", "")
                        , sfieldtiengiuong, sfielddoituong);
                }

                if (ads.Tables[0].Rows.Count <= 0)
                {
                    progressBar1.Value = progressBar1.Maximum;
                    timer1.Enabled = false;
                    progressBar1.Value = 0;
                    MessageBox.Show(this, "Không có số liệu báo cáo", "Thông báo", MessageBoxButtons.OK);
                    return;
                }
                int column1 = 0;//lay index cot nhom(hoac loai) vien phi dau tien. 
                for (int i = 0; i < ads.Tables[0].Columns.Count; i++)
                {
                    if (ads.Tables[0].Columns[i].ToString().IndexOf("LLLL") == 0)
                    {
                        column1 = i; break;
                    }
                }
                //

                //khai bao dong du lieu de tinh TONG
                DataRow r = ads.Tables[0].NewRow();
                //set cac gia tri mac dinh.
                for (int i = 0; i < ads.Tables[0].Columns.Count; i++)
                {
                    if (ads.Tables[0].Columns[i].DataType.ToString() == "System.Decimal")
                    {
                        r[i] = "0";
                    }
                }
                //
                for (int i = 0; i < ads.Tables[0].Rows.Count; i++)
                {
                    for (int j = 0; j < ads.Tables[0].Columns.Count; j++)
                    {
                        if (ads.Tables[0].Columns[j].DataType.ToString() == "System.Decimal")
                        {
                            try
                            {
                                r[j] = decimal.Parse(r[j].ToString()) + decimal.Parse(ads.Tables[0].Rows[i][j].ToString());
                            }
                            catch
                            {
                                r[j] = 0;
                            }
                        }
                    }
                }
                try
                {
                    r["mabn"] = "Tổng: " + ads.Tables[0].Rows.Count.ToString();
                }
                catch
                {
                    try
                    {
                        r["ngay"] = "Tổng: " + ads.Tables[0].Rows.Count.ToString();
                    }
                    catch
                    {
                    }
                }

                ads.Tables[0].Rows.Add(r);//them dong TONG vao table.
                #region Export
                //try
                //{
                //    if (!System.IO.Directory.Exists("..//..//Export"))
                //    {
                //        System.IO.Directory.CreateDirectory("..//..//Export");
                //    }
                //}
                //catch
                //{
                //}
                //string apath = Application.ExecutablePath;
                //apath = apath.Substring(0, apath.LastIndexOf("\\"));
                //apath = apath.Substring(0, apath.LastIndexOf("\\"));
                //apath = apath.Substring(0, apath.LastIndexOf("\\"));
                //apath = apath.Replace("\\", "//");
                //switch (v_format)
                //{
                //    case "HTML":
                //        f_Export_HTML(ads, m_ds2, apath + "//Export//baocaothanhtoanravien");
                //        break;
                //    default:
                //        f_Export_Excel(ads, m_ds2, apath + "//Export//baocaothanhtoanravien");
                //        break;
                //}
                #endregion
                dataGridView1.DataSource = ads.Tables[0];
                dataGridView2.DataSource = m_ds2.Tables[0];

            } ////////////////////////
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void f_KetxuatAna_Free()
        {
            try
            {

                dsAna = hddt.dataSetFromSql("mssql", "select sochungtu from  BANGDULIEUVIENPHI_TAM where tongcong=bhxh and convert(varchar(10),giophut,103) between '" + tungay.Text + "' and '" + denngay.Text + "'"); //Lọc những chứng từ 0đ đã có trong db ana

                string asql = "", asql1 = "", aexp = "";
                string atmp = "";
                string aloaibn = f_Get_CheckID(tree_Loaibn);
                string asqlht = "select * from v_hoantra where to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy')";

                DataSet ads = new DataSet();
                string atable = "", s_tyle = "";
                atable = "h.ma";
                s_tyle = s_bc1;
                m_ds2 = m_ds1.Clone();
                if (tungay.Text.Substring(3, 7) != "09/2008") s_tyle = "100";//Benh vien my phuoc 
                for (int i = 0; i < tree_Field.Nodes.Count; i++)
                {
                    if (tree_Field.Nodes[i].Checked)
                    {
                        try
                        {
                            m_ds2.Tables[0].Rows.Add(new string[] { tree_Field.Nodes[i].Tag.ToString().Trim(), tree_Field.Nodes[i].Text.Trim() });
                        }
                        catch
                        {
                        }
                    }
                }
                //

                bool ok = false;
                for (int i = 0; i < tree_Loai.Nodes.Count; i++)
                {
                    if (tree_Loai.Nodes[i].Checked)
                    {
                        ok = true;
                        break;
                    }
                }

                atmp = "";
                string athucthu = "";//24/09/2013 chỉ lấy số tiền thực thu của BN.
                string aphongkham = "";//14/10/2013 lấy số liệu của BN phòng khám.
                if (ok)
                {
                    for (int i = 0; i < tree_Loai.Nodes.Count; i++)
                    {
                        if (tree_Loai.Nodes[i].Checked)
                        {
                            try
                            {
                                atmp = atmp + ",sum(decode(" + atable + ","
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString()
                                    + ",b.soluong*b.dongia "
                                    + "+ b.soluong*b.dongia*nvl(b.vat,0)/100,0))*"
                                    + s_tyle + "/100 LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                athucthu = athucthu + ",sum(decode(" + atable + ","
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString()
                                    + ",b.sotien-b.bhyttra-b.mien "
                                    + "+ b.soluong*b.dongia*nvl(b.vat,0)/100,0))"
                                    + " LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                aphongkham = aphongkham + ",sum(decode(" + atable + ",'"
                                    + m_ds.Tables[0].Rows[i]["MA"].ToString().Trim()
                                    + "',b.soluong*b.dongia-b.thieu,0))*"
                                    + s_tyle + "/100 LLLL" + tree_Loai.Nodes[i].Tag.ToString();
                                m_ds2.Tables[0].Rows.Add(new string[] { "LLLL" + tree_Loai.Nodes[i].Tag.ToString(), tree_Loai.Nodes[i].Text.Trim() });
                            }
                            catch (Exception exx)
                            {
                                MessageBox.Show(exx.ToString());
                            }
                        }
                    }
                    atmp = atmp.Trim().Trim(',').Trim();
                    if (chkChitiet.Checked)
                    {
                        //lay cau truy van lay so tien thuc thu theo nhom(hoac loai) VP.
                        athucthu = athucthu.Trim().Trim(',').Trim();
                        atmp = athucthu;
                    }
                    if (atmp.Length > 0)
                    {
                        atmp = "," + atmp.Trim(',');
                    }
                }

                string asqldmbd = "select a1.id id, c1.id_loai id_loai from d_dmbd a1, d_dmnhom b1, (select a0.ma, min(nvl(b0.id,0)) id_loai from v_nhomvp a0, v_loaivp b0 where a0.ma=b0.id_nhom(+) group by a0.ma) c1 where a1.manhom=b1.id(+) and b1.nhomvp=c1.ma(+)";

                if (rd5.Checked)//Nhom theo khoa
                {
                    //Nếu chkTheokhoa có chọn, sẽ in theo khoa và tách theo tung khoa điều trị.
                    if (chkTheokhoa.Checked == true)
                    {
                        asql = "select g.makp ngay, g.tenkp, g.viettat, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0) thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and b.makp=g.makp(+) " + aexp + " group by g.makp,g.tenkp,g.viettat order by g.tenkp asc";
                        asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, g.makp ngay,g.tenkp,g.viettat, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and b.makp=g.makp(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')', g.makp,g.tenkp,g.viettat ";
                        asql1 += "order by d.sohieu asc, a.sobienlai asc, g.tenkp asc";
                    }
                    else
                    {
                        asql = "select g.makp ngay, g.tenkp, g.viettat, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0) thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.makp=g.makp(+) " + aexp + " group by g.makp,g.tenkp,g.viettat order by g.tenkp asc";
                        asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, g.makp ngay,g.tenkp,g.viettat, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h, btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and a.makp=g.makp(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')', g.makp,g.tenkp,g.viettat ";
                        asql1 += "order by d.sohieu asc, a.sobienlai asc, g.tenkp asc";
                    }
                }
                else if (rd4.Checked)//Nhom theo ngay
                {
                    asql = "select to_char(a.ngay,'dd/mm/yyyy') ngay, count(a.id) sohoadon, sum(nvl(b.sotien,0))*" + s_tyle + "/100 sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucthu, nvl(sum(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then nvl(a.tamung,0)-nvl(b.sotien,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end),0)*" + s_tyle + "/100 thucchi " + atmp + " from v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 ma from dual) h where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.id=b.id and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and a.id=c.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null " + aexp + " group by to_char(a.ngay,'dd/mm/yyyy') order by to_char(a.ngay,'dd/mm/yyyy') asc";
                    asql1 = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)*" + s_tyle + "/100) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.id=c.id(+) and c.maduyet=cc.ma(+) and a.quyenso=d.id(+) and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) " + aexp + " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, a.ngay, c.ghichu, cc.ten, e.hoten ||' ('||to_char(e.userid)||')' order by d.sohieu asc, a.sobienlai asc, a.ngay asc";
                }
                else if (rd3.Checked)//Nhóm theo biên lai
                {
                    //nếu chkTheokhoa có chọn, sẽ in theo biên lai và theo khoa.
                    if (chkTheokhoa.Checked == true)
                    {
                        asql = "select to_char(a.id) id, to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu, c.ghichu, ccc.ten lydomien, cc.ten nguoikymien, aaa.mabn, aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien, sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt, sum(nvl(a.tamung,0)) tamung, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)>=0 then nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0) end,0)*" + s_tyle + "/100) thucthu, sum(nvl(case when nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0)<0 then (nvl(b.sotien,0)-nvl(a.tamung,0)-nvl(b.bhyttra,0)-nvl(b.mien,0))*-1 end,0)*" + s_tyle + "/100) thucchi, e.hoten nguoithu, e.hoten ||' ('||to_char(e.userid)||')' userid,g.makp,g.tenkp " + atmp + " from btdbn aaa, v_ttrvds aa, v_ttrvll a, (" + asqlht + ") aht, v_ttrvct b, v_miennoitru c, v_dsduyet cc, v_lydomien ccc, v_quyenso d, v_dlogin e, (select id, id_loai from v_giavp union all select id, id_loai from (" + asqldmbd + ")) f, (select id, id_nhom from v_loaivp union all select 0 id, 0 id_nhom from dual) g, (select ma from v_nhomvp union all select 0 from dual) h,btdkp_bv g where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')>=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')<=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') and aa.id=a.id and aa.mabn=aaa.mabn(+) and a.id=b.id(+) and a.id=c.id(+) and c.maduyet=cc.ma(+) and c.lydo=ccc.id(+) and a.quyenso=d.id(+) and a.quyenso=aht.quyenso(+) and a.sobienlai=aht.sobienlai(+) and aht.id is null and a.userid=e.id(+) and a.userid=e.id(+) and b.mavp= f.id(+) and f.id_loai=g.id(+) and g.id_nhom=h.ma(+) and c.maduyet=cc.ma(+) and b.makp=g.makp " + aexp;
                        asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn, aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt, a.mien, e.hoten, a.ngay, c.ghichu, cc.ten, ccc.ten, e.hoten ||' ('||to_char(e.userid)||')',g.makp,g.tenkp order by d.sohieu asc, a.sobienlai asc, a.ngay asc";
                    }
                    else
                    {
                        #region Query lấy dữ liệu

                        #region bn noi tru + ngoai tru
                        #region select
                        asql = "select a.loaibn,to_char(a.id) id, to_char(a.quyenso) quyensoid"
                            + " , to_char(a.sobienlai) sobienlai, to_char(a.ngay,'dd/mm/yyyy') ngay"
                            + " , d.sohieu quyenso, d.sohieu||' - '||to_char(a.sobienlai) sochungtu"
                            + ", c.ghichu, ccc.ten lydomien, cc.ten nguoikymien, aaa.mabn" 
                            + ", aaa.hoten, aaa.namsinh, sum(nvl(b.sotien,0)*" + s_tyle + "/100) sotien"
                            + ", sum(nvl(b.mien,0)) mien, sum(nvl(b.bhyttra,0)) bhyt"
                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",i.chenhlechdv")
                            + ", nvl(a.tamung,0) tamung"
                            + ", nvl(case when sum(nvl(b.sotien,0))"
                            + "-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0))>=0 "
                            + " then sum(nvl(b.sotien,0))-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))"
                            + "-sum(nvl(b.mien,0)) end,0)*" + s_tyle + "/100 thucthu"
                            + ", nvl(case when sum(nvl(b.sotien,0))-nvl(a.tamung,0)"
                            + "-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0))<0 then (sum(nvl(b.sotien,0))"
                            + "-nvl(a.tamung,0)-sum(nvl(b.bhyttra,0))-sum(nvl(b.mien,0)))*-1 end,0)*"
                            + s_tyle + "/100 thucchi, e.hoten nguoithu"
                            + ",0.0 tongtamung"
                            + ",0.0 tonghoantamung"
                            + ",to_char(aa.ngayvao,'dd/mm/yyyy') ngayvao"
                            + ",to_char(aa.ngayra,'dd/mm/yyyy') ngayra"
                            + ",to_char(aa.maql) maql,'' doituong"
                            + ", e.hoten ||' ('||to_char(e.userid)||')' userid,g2.makp"
                            + ",g2.tenkp " + atmp;
                        #endregion
                        #region from
                        asql += " from  v_ttrvds aa"
                            + " inner join v_ttrvll a on aa.id=a.id"
                            + " left join btdbn aaa on aa.mabn=aaa.mabn"
                            + " left join (" + asqlht + ") aht on a.quyenso=aht.quyenso and a.sobienlai=aht.sobienlai"
                            + " left join v_ttrvct b on a.id=b.id"
                            + " left join v_miennoitru c on a.id=c.id"
                            + " left join v_dsduyet cc on c.maduyet=cc.ma"
                            + " left join v_lydomien ccc on c.lydo=ccc.id"
                            + " left join v_quyenso d on a.quyenso=d.id"
                            + " left join v_dlogin e on a.userid=e.id"
                            + " left join (select id, id_loai from v_giavp union all select id"
                                + ",id_loai from (" + asqldmbd + ")) f on b.mavp= f.id"
                            + " left join (select id, id_nhom from v_loaivp union all select 0 id"
                                + ",0 id_nhom from dual) g on f.id_loai=g.id"
                            + " left join (select ma from v_nhomvp union all "
                                + "select 0 from dual) h on g.id_nhom=h.ma"
                            + " inner join btdkp_bv g2 on a.makp=g2.makp"

                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",V_TTRVCT_chenhlech i ")

                        #endregion
                        #region where
                            + " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                            + " >=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                            + " and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                            + " <=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                            + " and aht.id is null "
                            + ((chkkhongtinhchenhlech.Checked) ? "" : " and a.id=i.id and a.makp=i.makp  ")
                            + aexp;
                        //if (hddtID.Length > 0)
                        //    asql += "  and lower(d.sohieu) ||' - '|| to_char(a.sobienlai ) in (" + hddtID + ")";
                        #endregion
                        #region group by
                        asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu, aaa.mabn"
                            + ", aaa.hoten, aaa.namsinh, a.sotien, a.tamung, a.bhyt"
                            + ", a.mien, e.hoten, a.ngay "
                            + ",c.ghichu, cc.ten, ccc.ten" 
                            + ", e.hoten ||' ('||to_char(e.userid)||')',g2.makp,g2.tenkp"
                            + ((chkkhongtinhchenhlech.Checked) ? "" : ",i.chenhlechdv ")
                            + ",aa.ngayvao"
                            + ",aa.ngayra"
                            + ",aa.maql,a.loaibn";
                        #endregion
                        #endregion

                        #region bn phong kham(thu truc tiep)

                        if (aloaibn == "" || aloaibn.IndexOf("3") > -1)
                        {
                            asql += " union all ";
                            asql = asql + " select a.loaibn,to_char(a.id) id"
                                + ", to_char(a.quyenso) quyensoid, to_char(a.sobienlai) sobienlai"
                                + ", to_char(a.ngay,'dd/mm/yyyy') ngay, d.sohieu quyenso"
                                //+ ", d.sohieu||' - '||to_char(a.sobienlai,'0000000') sochungtu"
                                + ", d.sohieu||' - '||to_char(a.sobienlai) sochungtu"
                                //+ ", c.ghichu, ccc.ten lydomien, cc.ten nguoikymien" //Đóng kiểm tra miễn giảm sau
                                + ", '' ghichu, '' lydomien, '' nguoikymien"
                                + ", a.mabn, a.hoten, a.namsinh"
                                + ", sum(nvl(b.soluong,0)*nvl(b.dongia,0))*" + s_tyle + "/100 sotien"
                                + ", sum(decode(b.madoituong,1,nvl(b.mien,0),0)) bhyt"
                                + ", nvl(c.sotien,0) mien"//, sum(nvl(b.thieu,0)) thieu"  
                                + ", 0 as mien"
                                + ((chkkhongtinhchenhlech.Checked) ? "" : ",0 chenhlechdv")
                                + ",0 tamung, sum(nvl(b.soluong,0)*nvl(b.dongia,0)*" + s_tyle + "/100"
                                //+ "-nvl(decode(b.madoituong,1,b.mien,0),0)-nvl(b.thieu,0)) -nvl(c.sotien,0) thucthu" //Đóng kiểm tra miễn giảm sau
                                + "-nvl(decode(b.madoituong,1,b.mien,0),0)-nvl(b.thieu,0))  thucthu"
                                + ",0 thucchi, e.hoten nguoithu"
                                + ",0.0 tongtamung"
                                + ",0.0 tonghoantamung"
                                + ",to_char(a.ngay,'dd/mm/yyyy') ngayvao"
                                + ",to_char(a.ngay,'dd/mm/yyyy') ngayra"
                                + ",to_char(a.maql) maql,'' doituong "
                                + ", e.hoten ||' ('||to_char(e.userid)||')' userid "
                                + ",i.makp,i.tenkp " + aphongkham

                                + " from v_vienphill a"
                                + " left join v_vienphict b on a.id=b.id"
                                + " left join (" + asqlht + ") aa on a.quyenso=aa.quyenso and a.sobienlai=aa.sobienlai"
                                //+ " left join v_mienngtru c on a.id=c.id left join v_dsduyet cc on c.maduyet=cc.ma left join v_lydomien ccc on c.lydo=ccc.id " //Đóng kiểm tra miễn giảm sau
                                + " left join v_quyenso d on a.quyenso=d.id"
                                + " left join v_dlogin e on a.userid=e.id"
                                + " left join (select id, id_loai from v_giavp union all "
                                + "select id, nvl(id_loai,0) id_loai from (" + asqldmbd + "))"
                                + " f on b.mavp= f.id"
                                + " left join (select id, id_nhom from v_loaivp union all "
                                + " select 0 id, 0 id_nhom from dual) g on f.id_loai=g.id"
                                + " left join (select ma from v_nhomvp "
                                + " union all select 0 from dual) h on g.id_nhom=h.ma"
                                + " left join btdkp_bv i on b.makp=i.makp"


                                + " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                                + " >=to_date('" + tungay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                                + " and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"
                                + " <=to_date('" + denngay.Text.Trim().Substring(0, 10) + "','dd/mm/yyyy') "
                                + " and aa.id is null "
                                + aexp;
                            //if (hddtID.Length > 0)
                            //    asql += "  and lower(d.sohieu) ||' - '|| to_char(a.sobienlai ) in (" + hddtID + ")";
                            asql += " group by a.id, a.quyenso, a.sobienlai, d.sohieu"
                            + ", a.mabn, a.hoten, a.namsinh, e.hoten"
                            + ",a.ngay,e.hoten ||' ('||to_char(e.userid)||')'"
                            + ",a.maql,a.ngay,i.tenkp,i.makp,a.loaibn"
                            ;
                        }
                        #endregion

                        asql = "select * from (" + asql + ") "
                            + " where thucthu=0 order by quyenso asc, sobienlai asc, ngay asc"
                            ;
                        #endregion
                    }
                }

                if (m_ds2.Tables[0].Rows.Count <= 0)
                {
                    progressBar1.Value = progressBar1.Maximum;
                    timer1.Enabled = false;
                    progressBar1.Value = 0;
                    MessageBox.Show(this, "Chọn thông tin báo cáo cần hiển thị", "Thông báo", MessageBoxButtons.OK);
                    return;
                }
                ads = m.get_data_all_vp(m.StringToDate(tungay.Text), m.StringToDate(denngay.Text), asql);
                if (rd4.Checked || rd5.Checked)
                {
                    #region
                    DataSet ads1 = new DataSet();
                    ads1 = m.get_data_mmyy(asql1, tungay.Text, denngay.Text);
                    for (int i = 0; i < ads.Tables[0].Rows.Count; i++)
                    {
                        ads.Tables[0].Rows[i]["sohoadon"] = 0;
                        ads.Tables[0].Rows[i]["sotien"] = 0;
                        ads.Tables[0].Rows[i]["bhyt"] = 0;
                        ads.Tables[0].Rows[i]["tamung"] = 0;
                        ads.Tables[0].Rows[i]["mien"] = 0;
                        ads.Tables[0].Rows[i]["thucthu"] = 0;
                        ads.Tables[0].Rows[i]["thucchi"] = 0;
                        try
                        {
                            foreach (DataRow r1 in ads1.Tables[0].Select("ngay='" + ads.Tables[0].Rows[i]["ngay"].ToString() + "'"))
                            {
                                ads.Tables[0].Rows[i]["sohoadon"] = decimal.Parse(ads.Tables[0].Rows[i]["sohoadon"].ToString()) + 1;
                                ads.Tables[0].Rows[i]["sotien"] = decimal.Parse(ads.Tables[0].Rows[i]["sotien"].ToString()) + decimal.Parse(r1["sotien"].ToString());
                                ads.Tables[0].Rows[i]["bhyt"] = decimal.Parse(ads.Tables[0].Rows[i]["bhyt"].ToString()) + decimal.Parse(r1["bhyt"].ToString());
                                ads.Tables[0].Rows[i]["tamung"] = decimal.Parse(ads.Tables[0].Rows[i]["tamung"].ToString()) + decimal.Parse(r1["tamung"].ToString());
                                ads.Tables[0].Rows[i]["mien"] = decimal.Parse(ads.Tables[0].Rows[i]["mien"].ToString()) + decimal.Parse(r1["mien"].ToString());
                                ads.Tables[0].Rows[i]["thucthu"] = decimal.Parse(ads.Tables[0].Rows[i]["thucthu"].ToString()) + decimal.Parse(r1["thucthu"].ToString());
                                ads.Tables[0].Rows[i]["thucchi"] = decimal.Parse(ads.Tables[0].Rows[i]["thucchi"].ToString()) + decimal.Parse(r1["thucchi"].ToString());
                            }
                        }
                        catch
                        {
                        }
                    }
                    #endregion
                }

                //bv huyet hoc ha noi muon lay tong tien tam ung(bao gom da hoan tra)
                //va tong tien da hoan tam ung.
                string stam = f_Get_CheckID(tree_Field);
                string sfieldtamung = (stam.IndexOf("TONGTAMUNG".ToUpper()) == -1) ? "" : "TONGTAMUNG".ToUpper();
                string sfieldhoantamung = (stam.IndexOf("TONGhoanTAMUNG".ToUpper()) == -1) ? "" : "TONGhoanTAMUNG".ToUpper();
                string sfielddoituong = (stam.IndexOf("doituong".ToUpper()) == -1) ? "" : "doituong".ToUpper();
                stam = f_Get_CheckID(tree_Loai);
                string sfieldtiengiuong = "";
                foreach (string st in stam.Split(','))
                {
                    if (st.IndexOf("101") > -1)//lay ten cot tien giuong dich vu
                    {
                        sfieldtiengiuong = "LLLL" + st;
                        break;
                    }
                }

                if ((sfieldtamung != "" || sfieldhoantamung != ""
                    || sfieldtiengiuong != "" || sfielddoituong != "") && ads.Tables[0].Rows.Count > 0)
                {
                    f_set_tongtamung_tonggiuong(ads.Tables[0]
                        , dsAna.Tables[0], sfieldhoantamung, sfieldtamung
                        , sfieldtiengiuong.Replace("101", "")
                        , sfieldtiengiuong, sfielddoituong);
                }
                else { return; }
                if (ads.Tables[0].Rows.Count <= 0)
                {
                    progressBar1.Value = progressBar1.Maximum;
                    timer1.Enabled = false;
                    progressBar1.Value = 0;
                    MessageBox.Show(this, "Không có số liệu báo cáo", "Thông báo", MessageBoxButtons.OK);
                    return;
                }
                int column1 = 0;//lay index cot nhom(hoac loai) vien phi dau tien. 
                for (int i = 0; i < ads.Tables[0].Columns.Count; i++)
                {
                    if (ads.Tables[0].Columns[i].ToString().IndexOf("LLLL") == 0)
                    {
                        column1 = i; break;
                    }
                }
                //

                //khai bao dong du lieu de tinh TONG
                DataRow r = ads.Tables[0].NewRow();
                //set cac gia tri mac dinh.
                for (int i = 0; i < ads.Tables[0].Columns.Count; i++)
                {
                    if (ads.Tables[0].Columns[i].DataType.ToString() == "System.Decimal")
                    {
                        r[i] = "0";
                    }
                }
                //
                for (int i = 0; i < ads.Tables[0].Rows.Count; i++)
                {
                    for (int j = 0; j < ads.Tables[0].Columns.Count; j++)
                    {
                        if (ads.Tables[0].Columns[j].DataType.ToString() == "System.Decimal")
                        {
                            try
                            {
                                r[j] = decimal.Parse(r[j].ToString()) + decimal.Parse(ads.Tables[0].Rows[i][j].ToString());
                            }
                            catch
                            {
                                r[j] = 0;
                            }
                        }
                    }
                }
                try
                {
                    r["mabn"] = "Tổng: " + ads.Tables[0].Rows.Count.ToString();
                }
                catch
                {
                    try
                    {
                        r["ngay"] = "Tổng: " + ads.Tables[0].Rows.Count.ToString();
                    }
                    catch
                    {
                    }
                }

                ads.Tables[0].Rows.Add(r);//them dong TONG vao table.
                dataGridView1.DataSource = ads.Tables[0];
                dataGridView2.DataSource = m_ds2.Tables[0];

            } ////////////////////////
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void f_KetxuatAna_KyQuy()
        {
            DataSet ds_Kyquy_Med = new DataSet();
            DataSet ds_Kyquy_Ana = new DataSet();
            sql = "SELECT 'I. KÝ QUỸ' LOAIPHIEU,'01' MASO,to_char(a.ngayud,'dd/mm/yyyy') GIOPHUT,'' COMPUTERNAME,c.userid,C.HOTEN,A.MAKP MAKHOA, D.TENKP TENKHOA, A.MABN,B.HOTEN TENBN, B.NAMSINH,";
            sql += " E.DOITUONG,F.SOHIEU QUYENSO, A.SOBIENLAI,F.SOHIEU || ' - ' || A.SOBIENLAI as SOCHUNGTU,A.SOTIEN TONGCONG,0 SOTIENHUY,A.NGAY NGAYCT,A.NGAY NGAYHOADON,'' CUATHU,";
            sql += " c.hoten as NGUOITHANHTOAN,case when a.ngaytra is not null then 1 else 0 end as dahoantra ,a.ngaytra thoigianhoantra";
            sql += " FROM xxx.V_TAMUNG A INNER JOIN BTDBN B ON A.MABN = B.MABN INNER JOIN V_DLOGIN C ON C.ID = A.USERID INNER JOIN BTDKP_BV D ON D.MAKP = A.MAKP INNER JOIN DOITUONG E ON E.MADOITUONG = A.MADOITUONG INNER JOIN V_QUYENSO F ON F.ID = A.QUYENSO";
            sql += " WHERE to_date(to_char(a.ngay, 'dd/mm/yyyy'), 'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "', 'dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "', 'dd/mm/yyyy')";
            sql += " ORDER BY A.NGAY";
            ds_Kyquy_Med = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10));
            sql = "select  * from  BANGDULIEUKYQUY_TAM WHERE convert(varchar(10), ngayhoadon, 103) between '" + tungay.Text.Substring(0, 10) + "' and '" + denngay.Text.Substring(0, 10) + "'";
            ds_Kyquy_Ana = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10), "mssql");

            if (ds_Kyquy_Med.Tables[0].Rows.Count > 0)
            {
                DataRow[] arrdr = new DataRow[0];
                string kqHoten, kqTenkhoa, kqHotenBN, kqDoituong, loaiPhieu,maSoPhieu;
                sttAna = 0;
                foreach (DataRow rkq in ds_Kyquy_Med.Tables[0].Rows)
                {

                    try
                    {
                        arrdr = ds_Kyquy_Ana.Tables[0].Select("quyenso='" + rkq["QUYENSO"].ToString() + "' and sobienlai='" + rkq["SOBIENLAI"].ToString() + "'");
                    }
                    catch { }
                    if (arrdr.Length == 0)
                    {
                       
                        sttAna++;
                        if (sttAna == 1)
                            sttAna = sttVienphi(sttAna, "BANGDULIEUKYQUY_TAM", rkq["NGAYCT"].ToString());
                        //end get stt
                        kqHoten = m.Replace_ErrorFont(rkq["hoten"].ToString());
                        kqTenkhoa = m.Replace_ErrorFont(rkq["tenkhoa"].ToString());
                        kqHotenBN = m.Replace_ErrorFont(rkq["tenbn"].ToString());
                        kqDoituong = m.Replace_ErrorFont(rkq["doituong"].ToString());
                        if(rkq["userid"].ToString()== "trucvp" || rkq["userid"].ToString() == "truc")
                        {
                            maSoPhieu = "01.04";
                            loaiPhieu = "IV. DOANH THU TRỰC";
                        }
                        else
                        {
                            maSoPhieu = "01";
                            loaiPhieu = "I. KÝ QUỸ";
                        }
                        try
                        {
                            m.upd_ana_BANGDULIEUKYQUY_API(loaiPhieu, maSoPhieu, rkq["GIOPHUT"].ToString(), rkq["COMPUTERNAME"].ToString(), kqHoten, rkq["MAKHOA"].ToString(), kqTenkhoa,
                                rkq["MABN"].ToString(), kqHotenBN, rkq["NAMSINH"].ToString(), kqDoituong, rkq["QUYENSO"].ToString(), rkq["SOBIENLAI"].ToString(),
                                decimal.Parse(rkq["TONGCONG"].ToString()), decimal.Parse(rkq["SOTIENHUY"].ToString()), rkq["NGAYCT"].ToString(), sttAna, rkq["NGAYHOADON"].ToString(), rkq["CUATHU"].ToString(), rkq["thoigianhoantra"].ToString(), int.Parse(rkq["dahoantra"].ToString()));
                        }
                        catch
                        {

                        }
                    }
                }
            }
            else
                return;

        }
        private void f_KetxuatAna_BHYT()
        {
            this._dtketqua = this.f_taodataset();
            this._dtketqua.Tables[0].Clear();
            try
            {
                this._dtketqua.Tables[0].Columns.Remove("id");
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("t_thuoc", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("t_cls", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("tyle", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("makp_byt");
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("maql", typeof(decimal)).DefaultValue = 0;
            }
            catch
            {
            }
            try
            {
                this._dtketqua.Tables[0].Columns.Add("ngaythu", typeof(string));
            }
            catch
            {
            }
            string loaibn = "";
            if (this.rdb_ngoaitru.Checked)
            {
                loaibn = "2,3,4";
            }
            else if (this.rdb_noitru.Checked)
            {
                loaibn = "1";
            }
            this._dsChiPhi = new DataSet();
            this._dsChiPhiCT = new DataSet();
            this._dsChiPhiCT = this.m.f_get_chiphict(this.tungay.Text, this.denngay.Text, this.txtmabn.Text, loaibn, this.chkxuatkhoa.Checked, "", false, this.chklaythuocglivec.Checked);

            //_dsChiPhi = m.f_get_chiphi(this.tungay.Text, this.denngay.Text, this.txtmabn.Text, this.rdb_ngoaitru.Checked, this.rdb_noitru.Checked, this.chkxuatkhoa.Checked);
            _dsChiPhi = m.f_get_chiphi(this.tungay.Text, this.denngay.Text, this.txtmabn.Text, true,true, this.chkxuatkhoa.Checked);

            this.f_get_xuat_ngoaitru(this.tungay.Text, this.denngay.Text, this._dsChiPhi, this._dsChiPhiCT);

            if (chkExcel.Checked)
            {
                //this.cmbmaubaocao.SelectedValue.ToString() = "3";
                if (this.rdb_ngoaitru.Checked)
                {
                    if ((this.cmbmaubaocao.SelectedValue.ToString() != "2") && (this.cmbmaubaocao.SelectedValue.ToString() == "4"))
                    {
                        m.f_Ngoaitru_xuatExcel_mau79_Mau2(false, this._dtketqua, tungay.Text, denngay.Text);
                    }
                }
                else if (!(this.cmbmaubaocao.SelectedValue.ToString() == "1") && (this.cmbmaubaocao.SelectedValue.ToString() == "3"))
                {
                    m.f_Ngoaitru_xuatExcel_mau79_Mau2(false, this._dtketqua, tungay.Text, denngay.Text);
                }
            }

        }
         private void f_get_xuat_ngoaitru(string tungay, string denngay, DataSet dscp, DataSet dscpct)
        {
            //dscp.WriteXml("dt.xml", XmlWriteMode.WriteSchema);
           // dscpct.WriteXml("dtct.xml", XmlWriteMode.WriteSchema);
            string filterExpression = "";
            string ana_loaiphieu = "", ana_tenkp = "", ana_tenbn = "", ana_ngaysinhnam = "", ana_ngaysinhnu = "", ana_chandoan = "", ana_diachi = "", nam_qt="", thang_qt="";
            decimal ana_tyle = 0;
            foreach (DataRow row2 in dscp.Tables[0].Rows)
            {
                DataRow row;
                filterExpression = "maql=" + row2["maql"].ToString();
                if (row2["loaibn"].ToString() == "3")
                {
                    string str2 = m.f_get_ngayvao(row2["maql"].ToString(), row2["mabn"].ToString());
                    //string str2 = m.f_get_ngayvao_ana(row2["maql"].ToString(), row2["mabn"].ToString());
                    if (str2 != "")
                    {
                        row2["ngayvao"] = str2;
                    }
                    row2["ngayra"] = m.f_get_ngayvao_ra(row2["ngayvao"].ToString().Replace("0000", "0800"), 2.0);
                    
                }
                try
                {
                    row = this._dtketqua.Tables[0].Select(filterExpression)[0];
                }
                catch
                {
                    row = this._dtketqua.Tables[0].NewRow();
                    for (int j = 0; j < this._dtketqua.Tables[0].Columns.Count; j++)
                    {
                        try
                        {
                            row[j] = row2[this._dtketqua.Tables[0].Columns[j].ColumnName].ToString();
                        }
                        catch
                        {
                        }
                    }
                    row["phai"] = (row["phai"].ToString() == "2") ? 1 : 0;
                    try
                    {
                        row["ngaysinh"] = row2["namsinh"].ToString();
                    }
                    catch
                    {
                    }
                    try
                    {
                        row["lydo"] = 1;
                    }
                    catch
                    {
                    }
                    DataRow row3 = this.f_get_tungay(row2["maql"].ToString(), row2["loaibn"].ToString(), (row2["loaibn"].ToString() != "3") ? "" : row2["sothe"].ToString());
                    row["tungay"] = (row3["tungay"].ToString()).Trim(new char[] { ';' });
                    row["denngay"] = (row3["denngay"].ToString()).Trim(new char[] { ';' });
                    row["sothe"] = (row3["sothe"].ToString()).Trim(new char[] { ';' });
                    row["manoidk"] = (row3["mabv"].ToString()).Trim(new char[] { ';' });
                    this._dtketqua.Tables[0].Rows.Add(row);
                }
                if (long.Parse(row2["ngayvao"].ToString()) < long.Parse(row["ngayvao"].ToString()))
                {
                    row["ngayvao"] = row2["ngayvao"].ToString();
                }
                if (long.Parse(row2["ngayra"].ToString()) > long.Parse(row["ngayra"].ToString()))
                {
                    row["ngayra"] = row2["ngayra"].ToString();
                }
                try
                {
                    row["songay"] = int.Parse(row2["songaydt"].ToString());
                }
                catch
                {
                }
                try
                {
                    row["tongcong"] = decimal.Parse(row["tongcong"].ToString()) + decimal.Parse(dscpct.Tables[0].Compute("sum(sotien)", "id=" + row2["id"].ToString()).ToString());
                }
                catch
                {
                }
                try
                {
                    row["bhyttra"] = decimal.Parse(row["bhyttra"].ToString()) + decimal.Parse(dscpct.Tables[0].Compute("sum(bhyttra)", "id=" + row2["id"].ToString()).ToString());
                }
                catch
                {
                }
                for (int i = 1; i <= 14; i++)
                {
                    try
                    {
                        row["st_" + i.ToString()] = decimal.Parse(row["st_" + i.ToString()].ToString()) + decimal.Parse(dscpct.Tables[0].Compute("sum(sotien)", "id=" + row2["id"].ToString() + " and nhombhyt=" + i.ToString()).ToString());
                    }
                    catch
                    {
                    }
                }
                //Insert ana
                if(row2["loaibn"].ToString() == "1")
                {
                    ana_loaiphieu = "V.2 Doanh thu Bảo hiểm Nội trú";
                }
                else
                {
                    ana_loaiphieu = "V.1 Doanh thu Bảo hiểm Ngoại trú";
                }
                if (row2["phai"].ToString() == "0")
                    ana_ngaysinhnam = row2["namsinh"].ToString();
                else
                    ana_ngaysinhnu = row2["namsinh"].ToString();
                try { ana_tyle = decimal.Parse(row["bhyttra"].ToString()) / decimal.Parse(row["tongcong"].ToString()); }
                catch { }
                nam_qt = denngay.Substring(6, 4);
                thang_qt = denngay.Substring(3, 2);
                ana_chandoan = m.f_get_fix_maicd(row2["maicd"].ToString());
                ana_tenkp = m.Replace_ErrorFont(row2["tenkp"].ToString());
                ana_tenbn = m.Replace_ErrorFont(row2["hoten"].ToString());
                ana_diachi = m.Replace_ErrorFont(row2["diachi"].ToString());
                m.upd_ana_BANGDULIEUBAOHIEM_API(1, ana_loaiphieu, row2["ngaythu"].ToString(), row2["ngaythu"].ToString(), row2["ngaythu_ana"].ToString(), "aa", row2["makp_byt"].ToString(), ana_tenkp, row2["sothe"].ToString(),
                     row2["madkkcb"].ToString(), row2["tungay"].ToString(), row2["denngay"].ToString(), 0, ana_chandoan, row2["mabn"].ToString(), ana_tenbn, ana_ngaysinhnam, ana_ngaysinhnu, row2["namsinh"].ToString(),
                     ana_diachi, row2["ngaythu_ana"].ToString(), row2["ngayvao"].ToString(), row2["ngayra"].ToString(), decimal.Parse(row2["songaydt"].ToString()), decimal.Parse(row["tongcong"].ToString()), decimal.Parse(row2["bntra"].ToString()),
                    decimal.Parse(row["bhyttra"].ToString()), decimal.Parse(row["st_1"].ToString()), decimal.Parse(row["st_2"].ToString()), decimal.Parse(row["st_3"].ToString()), decimal.Parse(row["st_4"].ToString()), decimal.Parse(row["st_5"].ToString()), decimal.Parse(row["st_6"].ToString()),
                    decimal.Parse(row["st_10"].ToString()), decimal.Parse(row["st_11"].ToString()), decimal.Parse(row["st_12"].ToString()), 0, 0, 0, 0, 0, 0, 0, 0, 0, nam_qt, thang_qt);
                //end insert ana
            }

            this._dtketqua.AcceptChanges();
            //this.lblrefesh.Text = "Tổng số: " + this._dtketqua.Tables[0].Rows.Count.ToString();
            //this.lblrefesh.Refresh();

           
        }
        public DataRow f_get_tungay(string maql, string v_loaibn, string v_sothe)
        {
            string str = "";
            string str2 = "";
            string str3 = "";
            string str4 = "";
            string sql = "";
            string str6 = sql;
            string str7 = str6 + "#select distinct to_char(tungay,'yyyymmdd') tungay,to_char(denngay,'yyyymmdd') denngay,substr(sothe,1,15) sothe,mabv from " + m.user + maql.Substring(2, 2) + maql.Substring(0, 2) + ".bhyt WHERE maql in(" + maql + ") " + ((v_sothe == "") ? "" : (" and sothe in('" + v_sothe + "')"));
            sql = (str7 + "#select distinct to_char(tungay,'yyyymmdd') tungay,to_char(denngay,'yyyymmdd') denngay,substr(sothe,1,15) sothe,mabv from " + m.user + ".bhyt WHERE maql =" + maql + " " + ((v_sothe == "") ? "" : (" and sothe in('" + v_sothe + "')"))).Trim(new char[] { '#' }).Replace("#", " union all ");
            System.Data.DataTable table = m.get_data(sql).Tables[0];
            try
            {
                foreach (DataRow row in table.Select("", "tungay asc"))
                {
                    if (row["tungay"].ToString().Length > 0)
                    {
                        str = str + row["tungay"].ToString() + ";";
                    }
                    if (row["denngay"].ToString().Length > 0)
                    {
                        str2 = str2 + row["denngay"].ToString() + ";";
                    }
                    if (row["sothe"].ToString().Length > 0)
                    {
                        str3 = str3 + row["sothe"].ToString() + ";";
                    }
                    if (row["mabv"].ToString().Length > 0)
                    {
                        str4 = str4 + row["mabv"].ToString() + ";";
                    }
                    if (v_loaibn != "1")
                    {
                        goto Label_0273;
                    }
                }
            }
            catch
            {
            }
        Label_0273:
            table.Clear();
            table.Rows.Add(new object[] { str.Trim(new char[] { ';' }), str2.Trim(new char[] { ';' }), str3.Trim(new char[] { ';' }), str4.Trim(new char[] { ';' }) });
            return table.Rows[0];
        }
        private DataSet f_taodataset()
        {
            DataSet set = new DataSet();
            set.Tables.Add();
            set.Tables[0].Columns.Add("mabn");
            set.Tables[0].Columns.Add("HOTEN");
            set.Tables[0].Columns.Add("ngaysinh");
            set.Tables[0].Columns.Add("phai", typeof(int)).DefaultValue = 0;
            set.Tables[0].Columns.Add("diachi");
            set.Tables[0].Columns.Add("sothe");
            set.Tables[0].Columns.Add("manoidk");
            set.Tables[0].Columns.Add("tungay");
            set.Tables[0].Columns.Add("denngay");
            set.Tables[0].Columns.Add("maicd");
            set.Tables[0].Columns.Add("maicdkt");
            set.Tables[0].Columns.Add("lydo");
            set.Tables[0].Columns.Add("traituyen", typeof(int));
            set.Tables[0].Columns.Add("ngayvao");
            set.Tables[0].Columns.Add("ngayra");
            set.Tables[0].Columns.Add("songay", typeof(int));
            set.Tables[0].Columns.Add("tongcong", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_1", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_2", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_3", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_4", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_5", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_6", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_7", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_8", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_9", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_10", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_11", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_12", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_13", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("st_14", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("bntra", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("bhyttra", typeof(decimal)).DefaultValue = 0;
            set.Tables[0].Columns.Add("makhuvuc");
            set.Tables[0].Columns.Add("tenkp");
            return set;
        }
        private void f_load_maubaocao()
        {
            DataRow row;
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("id");
            table.Columns.Add("loai");
            table.Columns.Add("ten");
            if (this.rdb_noitru.Checked)
            {
                row = table.NewRow();
                row[0] = 1;
                row[1] = 1;
                row[2] = "Mẫu BHYT 80a-HD";
                table.Rows.Add(row);
                row = table.NewRow();
                row[0] = 3;
                row[1] = 1;
                row[2] = "Mẫu BHYT 80a-HD excel";
                table.Rows.Add(row);
            }
            else
            {
                row = table.NewRow();
                row[0] = 2;
                row[1] = 2;
                row[2] = "Mẫu BHYT 79a-HD";
                table.Rows.Add(row);
                row = table.NewRow();
                row[0] = 4;
                row[1] = 2;
                row[2] = "Mẫu BHYT 79a-HD excel";
                table.Rows.Add(row);
            }
            this.cmbmaubaocao.DataSource = table;
            this.cmbmaubaocao.DisplayMember = "ten";
            this.cmbmaubaocao.ValueMember = "id";
        }
        private void load_grid()
        {
            //sql = "select sohoso,case when loaivp=1 then 'Tạm ứng' when loaivp=2 then 'Thu Trực Tiếp'  when loaivp=3 then 'TTRV'  when loaivp=4 then  'BHYT Ngoại trú' when loaivp=5 then  'Nhà Thuốc' end as loaivp";
            //sql += " ,makhachhang,tenkhachhang,case when phai=0 then 'Nam' else 'Nữ' end as phai,namsinh,diachi,quyenso,sobienlai,to_char(ngaylap,'dd/mm/yyyy') as ngaylap from tb_master where to_date(to_char(ngaylap,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy') ";
            //if (cbLoaiVP.SelectedIndex != 0 && cbLoaiVP.SelectedIndex != -1) sql += " and loaivp=" + cbLoaiVP.SelectedIndex;
            //sql += "group by sohoso,loaivp,makhachhang,tenkhachhang,case when phai=0 then 'Nam' else 'Nữ' end,namsinh,diachi,quyenso,sobienlai,ngaylap";
            //sql += " order by loaivp,ngaylap,tenkhachhang";
            //ds = m.get_data(sql);
            //dtg.DataSource = ds.Tables[0];
            sql = "SELECT COUNT(*) AS SUM FROM tb_master;";
            ds = m.get_data_mySQL(sql);
            record.Text = "MASTER:" + ds.Tables[0].Rows[0]["SUM"].ToString();
            sql = "SELECT COUNT(*) AS SUM FROM tb_detail;";
            ds = m.get_data_mySQL(sql);
            record.Text += " - DETAIL : " + ds.Tables[0].Rows[0]["SUM"].ToString();

        }
        private void insert_tamung_para()
        {
            sql = "select a.id,a.mavaovien as sohoso," + TamUng + " as loaivp,a.quyenso,a.sobienlai,a.userid as idthungan,c.hoten as hotenthungan, a.mabn as makhachhang,b.hoten as tenkhachhang,b.namsinh,b.phai,";
            sql += " b.diachi,0 as mien,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngaylap,'" + cbMMYY.Text + "' as mmyy from xxx.v_tamung a inner join (" + s_hanhchinh + ") b on a.mabn=b.mabn left join " + user + ".v_dlogin c on a.userid=c.id";
            sql += " where a.quyenso|| '/' ||a.sobienlai not in (select quyenso|| '/' ||sobienlai from xxx.v_hoantra)";
            sql += " and to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            dtth = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];
            if (dtth.Rows.Count > 0)
            {
                foreach (DataRow rth in dtth.Rows)
                {
                    try
                    {
                        m.upd_hddt_tonghop(decimal.Parse(rth["id"].ToString()), rth["sohoso"].ToString(),0, int.Parse(rth["loaivp"].ToString()), int.Parse(rth["quyenso"].ToString()), int.Parse(rth["sobienlai"].ToString()), int.Parse(rth["idthungan"].ToString()), rth["hotenthungan"].ToString(), rth["makhachhang"].ToString(), rth["tenkhachhang"].ToString(), rth["namsinh"].ToString(), int.Parse(rth["phai"].ToString()), rth["diachi"].ToString(), decimal.Parse(rth["mien"].ToString()), rth["ngaylap"].ToString(), rth["mmyy"].ToString());
                    }
                    catch { }
                    
                }
            } 


        }
        private void insert_thuTT_para()
        {

            // sql = " select to_char(a.id) as id,case when a.mavaovien=0 then to_char(a.maql) else to_char(a.mavaovien) end as sohoso," + ThuTrucTiep + " as loaivp,0 as tongtien,a.userid as idthungan,g.hoten as hotenthungan, a.mabn as makhachhang,b.hoten as tenkhachhang,b.namsinh,";
            sql = " select to_char(a.id) as id, quyenso || '_' ||sobienlai as sohoso,0 as tongtien,g.userid as idthungan,g.hoten as hotenthungan, a.mabn as makhachhang,b.hoten as tenkhachhang,b.namsinh,";
            sql += " b.diachi,to_char(a.ngay,'dd-mm-yyyy hh24:mi') as ngaylap,A.LOAIBN,";
            sql += " d.tenkp as TENKHOA,c.sohieubl as QUYENSO,a.sobienlai,c.sohieu || ' - ' || a.sobienlai as sochungtu";
            sql += " from xxx.v_vienphill a inner join (" + s_hanhchinh + ") b on a.mabn=b.mabn";
            //sql += " left join xxx.v_mienngtru f on f.id=a.id";
            sql += " left join " + user + ".v_dlogin g on a.userid=g.id";
            sql += " inner join btdkp_bv d on a.makp=d.makp inner join v_quyenso c on c.id=a.quyenso";
            sql += " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            sql += " and a.sobienlai<>-1 and a.quyenso<>-1 and a.userid<>-1";
            //sql += " and a.sobienlai=4711 and a.quyenso=3114";
            //hddt.dataSetFromSql("orc", "1", ThuTrucTiep, usermmyy,user, s_hanhchinh, tungay.Text, denngay.Text);
            dtth = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            //get tongtien
            sql = "select to_char(a.id) as id,sum(a.soluong*a.dongia) as tongtien from xxx.v_vienphict a inner join xxx.v_vienphill b on a.id=b.id";
            sql += " where to_date(to_char(b.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            sql += " group by a.id order by a.id";
            dttongtien = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            //end
            ds.AcceptChanges();
            sql = " select to_char(a.id) as id,a.stt,b.ma AS mahang,b.ten as tenhang,b.dvt as donvitinh,a.soluong,a.dongia,a.soluong*a.dongia as thanhtien,0 as tylebh,0 as mucbhtra,0 as bhxhtra,b.manhomdv,b.nhomdv,b.loaidv";
            sql += " from xxx.v_vienphict a";
            sql += " inner join xxx.v_vienphill a1 on a.id=a1.id";
            sql += " inner join (" + s_giavp_dmbd + ") b on a.mavp=b.mavp";
            sql += " where a.tra=0";
            sql += " and to_date(to_char(a1.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            sql += " order by a.id";
            dtct = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            if (dtth.Rows.Count > 0)
            {
                foreach (DataRow rth in dtth.Rows)
                {
                    try
                    {
                        hMasterTT_temp = hddt.dataSetFromSql("mysql", "select sohoso from tb_master where sohoso='" + rth["sohoso"].ToString() + "'");
                    }
                    catch { }
                    if (hMasterTT_temp.Tables[0].Rows.Count == 0 || hMasterTT_temp == null)
                    //if (hMasterTT_temp == null || hMasterTT_temp.Container == null)
                    {
                        string s_tenkhachhang = rth["tenkhachhang"].ToString();
                        s_tenkhachhang = m.Replace_ErrorFont(s_tenkhachhang);
                        string s_diachi = rth["diachi"].ToString();
                        s_diachi = m.Replace_ErrorFont(s_diachi);
                        s_diachi = s_diachi.Trim();
                        string s_hotenthungan = rth["hotenthungan"].ToString();
                        s_hotenthungan = m.Replace_ErrorFont(s_hotenthungan);
                        string s_khoadieutri = rth["TENKHOA"].ToString();
                        s_khoadieutri = m.Replace_ErrorFont(s_khoadieutri);
                        try
                        {
                            foreach (DataRow rtt in dttongtien.Select("id = " + rth["id"].ToString()))
                            {
                                l_tongtien = Decimal.Parse(rtt["tongtien"].ToString()); // Tong tien chua co miengiam
                            }
                        }
                        catch { l_tongtien = 0; }
                        try
                        {
                            if (m.upd_hddt_tonghop_mysql_API(rth["sohoso"].ToString(), l_tongtien, s_tenkhachhang, s_diachi, rth["idthungan"].ToString(), s_hotenthungan, rth["ngaylap"].ToString(), rth["sochungtu"].ToString(), int.Parse(rth["loaibn"].ToString()), s_khoadieutri, "", rth["makhachhang"].ToString()))
                            {
                                paratt = "";
                                foreach (DataRow rct in dtct.Select("id=" + rth["id"].ToString()))
                                {
                                    string s_tenhang = rct["tenhang"].ToString();
                                    s_tenhang = m.Replace_ErrorFont(s_tenhang);
                                    string s_donvitinh = rct["donvitinh"].ToString();
                                    s_donvitinh = m.Replace_ErrorFont(s_donvitinh);
                                    string s_nhomdv = rct["nhomdv"].ToString();
                                    s_nhomdv = m.Replace_ErrorFont(s_nhomdv);
                                    paratt += $"('{rth["sohoso"].ToString()}','{rth["sohoso"].ToString()}','{ rct["mahang"].ToString()}','{ rct["manhomdv"].ToString()}','{s_nhomdv}','{ rct["loaidv"].ToString()}','{ s_tenhang}','{ s_donvitinh}',{ double.Parse(rct["soluong"].ToString())},{ double.Parse(rct["dongia"].ToString())},{ double.Parse(rct["thanhtien"].ToString())},{ double.Parse(rct["tylebh"].ToString())},{ double.Parse(rct["mucbhtra"].ToString())},{ double.Parse(rct["bhxhtra"].ToString())} ),";
                                    //m.upd_hddt_chitiet_mysql_API(rth["sohoso"].ToString(), rct["mahang"].ToString(), rct["loaidv"].ToString(), s_tenhang, double.Parse(rct["dongia"].ToString()), s_donvitinh, double.Parse(rct["soluong"].ToString()), double.Parse(rct["thanhtien"].ToString()), double.Parse(rct["tylebh"].ToString()), double.Parse(rct["mucbhtra"].ToString()), double.Parse(rct["bhxhtra"].ToString()));
                                }
                            }
                            paratt = paratt.TrimEnd(',');
                            m.upd_hddt_chitiet_mysql_API(paratt);
                        }
                        catch { }
                    }
                }

            }
            else
                return;
            //End thu truc tiep
        }
        private void insert_TTRV_para()
        {
            sql = " select to_char(a.id) as id,ll.quyenso || '_' || ll.sobienlai as sohoso,ll.sotien-ll.mien-ll.bhyt as tongtien, a.mabn as makhachhang,b.hoten as tenkhachhang,b.diachi,c.userid as idthungan,c.hoten as hotenthungan,to_char(ll.ngay,'dd-mm-yyyy hh24:mi') as ngaylap";
            sql += " ,d.sohieu || ' - ' || ll.sobienlai as sochungtu,ll.loaibn,kp.tenkp as TENKHOA,'Đợt điều trị:   ' || to_char(a.ngayvao,'dd/mm/yyyy') ||' -> ' || to_char(a.ngayra,'dd/mm/yyyy') dotdieutri";
            sql += " from xxx.v_ttrvds a inner join xxx.v_ttrvll ll on a.id=ll.id inner join (" + s_hanhchinh + ") b on a.mabn=b.mabn";
            sql += " left join " + user + ".v_dlogin c on c.id=ll.userid";
            sql += " inner join btdkp_bv kp on ll.makp = kp.makp inner join v_quyenso d on d.id=ll.quyenso";
            sql += " where to_date(to_char(ll.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            dtth_TTRV = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            sql = " select  to_char(a.id) as id,b.ma as mahang,b.ten as tenhang,b.manhomdv,b.nhomdv,b.loaidv,b.dvt as donvitinh,a.soluong,a.dongia,a.sotien as thanhtien,case when a.madoituong=1 then round(a.bhyttra/a.sotien*100,0) else 0 end as TYLEBH,a.sotien as MUCBHTRA,a.bhyttra as BHXHTRA";
            sql += " from xxx.v_ttrvct a inner join (" + s_giavp_dmbd + ") b on a.mavp=b.mavp";
            sql += " inner join xxx.v_ttrvll a1 on a.id=a1.id";
            sql += " where a.soluong>0 and a.dongia>0";
            sql += " and to_date(to_char(a1.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            dtct_TTRV = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];
            if (dtth_TTRV.Rows.Count > 0)
            {
                foreach (DataRow rthTTRV in dtth_TTRV.Rows)
                {
                    try
                    {
                        try
                        {
                            hMasterTTRV_temp = hddt.dataSetFromSql("mysql", "select sohoso from tb_master where sohoso='" + rthTTRV["sohoso"].ToString() + "'");
                        }
                        catch { }
                        if (hMasterTTRV_temp.Tables[0].Rows.Count == 0 || hMasterTTRV_temp == null)

                        // if (hMasterTTRV_temp == null || hMasterTTRV_temp.Container == null)
                        {
                            string s_tenkhachhang = rthTTRV["tenkhachhang"].ToString();
                            s_tenkhachhang = m.Replace_ErrorFont(s_tenkhachhang);

                            string s_diachi = rthTTRV["diachi"].ToString();
                            s_diachi = m.Replace_ErrorFont(s_diachi);
                            s_diachi = s_diachi.Trim();

                            string s_hotenthungan = rthTTRV["hotenthungan"].ToString();
                            s_hotenthungan = m.Replace_ErrorFont(s_hotenthungan);

                            string s_khoadieutri = rthTTRV["TENKHOA"].ToString();
                            s_khoadieutri = m.Replace_ErrorFont(s_khoadieutri);
                            try
                            {
                                if (m.upd_hddt_tonghop_mysql_API(rthTTRV["sohoso"].ToString(), l_tongtien, s_tenkhachhang, s_diachi, rthTTRV["idthungan"].ToString(), s_hotenthungan, rthTTRV["ngaylap"].ToString(), rthTTRV["sochungtu"].ToString(), int.Parse(rthTTRV["loaibn"].ToString()), s_khoadieutri, rthTTRV["dotdieutri"].ToString(), rthTTRV["makhachhang"].ToString()))
                                {
                                    paratt = "";
                                    foreach (DataRow rctTTRV in dtct_TTRV.Select("id=" + rthTTRV["id"].ToString()))
                                    {
                                        string s_tenhang = rctTTRV["tenhang"].ToString();
                                        s_tenhang = m.Replace_ErrorFont(s_tenhang);
                                        string s_donvitinh = rctTTRV["donvitinh"].ToString();
                                        s_donvitinh = m.Replace_ErrorFont(s_donvitinh);
                                        string s_nhomdv = rctTTRV["nhomdv"].ToString();
                                        s_nhomdv = m.Replace_ErrorFont(s_nhomdv);
                                        paratt += $"('{rthTTRV["sohoso"].ToString()}','{rthTTRV["sohoso"].ToString()}','{ rctTTRV["mahang"].ToString()}','{ rctTTRV["manhomdv"].ToString()}','{s_nhomdv}','{ rctTTRV["loaidv"].ToString()}','{ s_tenhang}','{ s_donvitinh}',{ double.Parse(rctTTRV["soluong"].ToString())},{ double.Parse(rctTTRV["dongia"].ToString())},{ double.Parse(rctTTRV["thanhtien"].ToString())},{ double.Parse(rctTTRV["tylebh"].ToString())},{ double.Parse(rctTTRV["mucbhtra"].ToString())},{ double.Parse(rctTTRV["bhxhtra"].ToString())} ),";
                                    }
                                    paratt = paratt.TrimEnd(',');
                                    m.upd_hddt_chitiet_mysql_API(paratt);
                                }
                            }
                            catch { }
                        }

                    }
                    catch { }


                }
            }
            else
                return;
        }
        private void insert_NhaThuoc_para()
        {

            sql = " select a.id,a.id as sohoso," + NhaThuoc + " as loaivp,a.quyenso,a.sobienlai,0 as tongtien,0 as idthungan,'' as hotenthungan, a.mabn as makhachhang,case when b.hoten is null or b.hoten ='' then a.hoten else b.hoten end as tenkhachhang,case when b.namsinh is null or b.namsinh ='' then a.namsinh else b.namsinh end as namsinh,case when to_char(b.phai) is null or to_char(b.phai) ='' then 2 else b.phai end as phai, case when b.diachi is null or b.diachi ='' then a.diachi else b.diachi end as diachi,0 as mien,to_char(a.ngay,'dd/mm/yyyy hh24:mi') as ngaylap,'" + cbMMYY.Text + "' mmyy";
            sql += " from xxxd.d_ngtrull a left join (" + s_hanhchinh + ") b on a.mabn=b.mabn";
            sql += " where to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            //sql += " and b.mabn in (select c.mabn from xxxd.d_ngtrull c)";
            //sql += " and a.id=170613000176702972";
            
            dtth = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            dttongtien = m.get_data_mmyy("select id,sum(soluong*giaban) as tongtien from xxx.d_ngtruct group by id order by id", tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];

            sql = " select a.id,a.stt,a.madoituong,b.ma,b.ten as tenhang,b.dvt as donvitinh,a.soluong,a.giaban as dongia,a.soluong*a.giaban as thanhtien,0 as TYLEBH,0 as MUCBHTRA,0 as BHXHTRA,b.nhom," + NhaThuoc + " as loai,'" + cbMMYY.Text + "' mmyy";
            sql += " from xxxd.d_ngtruct a inner join (" + s_giavp_dmbd + ") b on a.mabd=b.mavp";
            sql += " inner join xxxd.d_ngtrull c on c.id=a.id";
            sql += " where to_date(to_char(c.ngay,'dd/mm/yyyy'),'dd/mm/yyyy') between to_date('" + tungay.Text.Substring(0, 10) + "','dd/mm/yyyy') and to_date('" + denngay.Text.Substring(0, 10) + "','dd/mm/yyyy')";
            dtct = m.get_data_mmyy(sql, tungay.Text.Substring(0, 10), denngay.Text.Substring(0, 10)).Tables[0];
            if (dtth.Rows.Count > 0)
            {
                foreach (DataRow rth in dtth.Rows)
                {
                    try
                    {
                        foreach (DataRow rtt in dttongtien.Select("id=" + rth["id"].ToString()))
                        {
                            l_tongtien = Decimal.Parse(rtt["tongtien"].ToString()); // Tong tien chua co miengiam
                        }
                    }
                    catch { l_tongtien = 0; }
                    try
                    {
                        if (m.upd_hddt_tonghop(decimal.Parse(rth["id"].ToString()), rth["sohoso"].ToString(),l_tongtien, int.Parse(rth["loaivp"].ToString()), int.Parse(rth["quyenso"].ToString()), int.Parse(rth["sobienlai"].ToString()), int.Parse(rth["idthungan"].ToString()), rth["hotenthungan"].ToString(), rth["makhachhang"].ToString(), rth["tenkhachhang"].ToString(), rth["namsinh"].ToString(), int.Parse(rth["phai"].ToString()), rth["diachi"].ToString(), decimal.Parse(rth["mien"].ToString()), rth["ngaylap"].ToString(), rth["mmyy"].ToString()))
                        {
                            foreach (DataRow rct in dtct.Select("id=" + rth["id"].ToString()))
                            {
                                m.upd_hddt_chitiet(decimal.Parse(rct["id"].ToString()), int.Parse(rct["stt"].ToString()), rct["ma"].ToString(), rct["tenhang"].ToString(), rct["donvitinh"].ToString(), decimal.Parse(rct["soluong"].ToString()), decimal.Parse(rct["dongia"].ToString()),
                                    decimal.Parse(rct["thanhtien"].ToString()), decimal.Parse(rct["tylebh"].ToString()), decimal.Parse(rct["mucbhtra"].ToString()), decimal.Parse(rct["bhxhtra"].ToString()), rct["nhom"].ToString(), int.Parse(rct["loai"].ToString()), rct["mmyy"].ToString());
                            }
                        }
                    }
                    catch { }

                }
            }
        }
        private void timer_insert_Tick(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;
            record.Text = "Đang chuyển dữ liệu";
            // insert_NhaThuoc_para();
            //insert_tamung_para();
            if(bHDDT)
            {
                insert_thuTT_para();
                insert_TTRV_para();
            }
            
            if (chkAna.Checked && bAna)
            {
                f_KetxuatAna_KyQuy();
                f_KetxuatAna();
                f_KetxuatAna_Free();
                f_KetxuatAna_BHYT();
            }
            record.Text = "Hoàn thành chuyển dữ liệu";

            Cursor = Cursors.Default;
        }
        private void timer1_Tick(object sender, System.EventArgs e)
        {
            try
            {
                progressBar1.Value = (progressBar1.Value >= progressBar1.Maximum) ? 0 : progressBar1.Value + 1;
                progressBar1.Update();
            }
            catch
            {
            }
        }
        private void rd3_CheckedChanged(object sender, EventArgs e)
        {
            f_Load_Tree_Field();
        }
        private void rd4_CheckedChanged(object sender, EventArgs e)
        {
            f_Load_Tree_Field();
        }
        private void rd3_Click(object sender, EventArgs e)
        {
            if (rd3.Checked == true)
            {
                chkTheokhoa.Visible = true;
            }
            else
            {
                chkTheokhoa.Visible = false;
            }
        }
        private void rd4_Click(object sender, EventArgs e)
        {
            if (rd4.Checked == true)
            {
                chkTheokhoa.Visible = false;
            }
        }
        private void rd5_Click(object sender, EventArgs e)
        {
            if (rd5.Checked == true)
            {
                chkTheokhoa.Visible = true;
            }
            else
            {
                chkTheokhoa.Visible = false;
            }
        }
        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            f_Set_CheckID(tree_Loai, chkAll.Checked);
        }
        private void chkAll1_CheckedChanged(object sender, EventArgs e)
        {
            f_Set_CheckID(tree_Field, chkAll1.Checked);
        }
        private void rd1_CheckedChanged(object sender, EventArgs e)
        {
            f_Load_Tree_Loai();
        }
        private void rd2_CheckedChanged(object sender, EventArgs e)
        {
            f_Load_Tree_Loai();
        }
        private void chkLoaibn_CheckedChanged(object sender, EventArgs e)
        {
            f_Set_CheckID(tree_Loaibn, chkLoaibn.Checked);
        }
        private void bKetxuat_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            m = new AccessDataApi(ipApi);

            nbPhut.Enabled = cbMMYY.Enabled = bKetxuat.Enabled = bStart.Enabled = bClose.Enabled = false;
            bPause.Enabled = true;
            user = m.user;
            usermmyy = user + cbMMYY.Text;
            #region Medi --> HDDT
            if (tMediSoft_HDDT_TT == null || !tMediSoft_HDDT_TT.IsAlive)
            {
                tMediSoft_HDDT_TT = new System.Threading.Thread(() =>
                {
                    while (true)
                    {
                        insert_thuTT_para();
                        nbPhut.Invoke(new MethodInvoker(() =>
                        {
                            System.Threading.Thread.Sleep(int.Parse((nbPhut.Value).ToString()) * interval);
                        }));
                    }
                });
               tMediSoft_HDDT_TT.Start();
            }

            if (tMediSoft_HDDT_TTRV == null || !tMediSoft_HDDT_TTRV.IsAlive)
            {
                tMediSoft_HDDT_TTRV = new System.Threading.Thread(() =>
                {
                    while (true)
                    {
                        insert_TTRV_para();
                        nbPhut.Invoke(new MethodInvoker(() =>
                        {
                            System.Threading.Thread.Sleep(int.Parse((nbPhut.Value).ToString()) * interval);
                        }));
                    }
                });
              tMediSoft_HDDT_TTRV.Start();
            }
            #endregion
            #region Medi --> Ana
            if (chkAna.Checked && (tHDDT_Ana == null || !tHDDT_Ana.IsAlive))
            {
                //Đẩy sang table KYQUY
                tHDDT_Ana_KyQuy = new System.Threading.Thread(() =>
                {
                    while (true)
                    {
                        f_KetxuatAna_KyQuy();

                        nbPhut.Invoke(new MethodInvoker(() =>
                        {
                            System.Threading.Thread.Sleep(int.Parse((nbPhut.Value).ToString()) * interval);
                        }));
                    }
                });
                //tHDDT_Ana.IsBackground = false;
                tHDDT_Ana_KyQuy.Start();

                //Đẩy sang table THANHTOANVIENPHI Những HĐ đã được xuất trên hddt
                tHDDT_Ana = new System.Threading.Thread(() =>
                {
                    while (true)
                    {
                        f_KetxuatAna();
                        f_KetxuatAna_Free();

                        nbPhut.Invoke(new MethodInvoker(() =>
                        {
                            System.Threading.Thread.Sleep(int.Parse((nbPhut.Value).ToString()) * interval);
                        }));
                    }
                });
                //tHDDT_Ana.IsBackground = false;
                tHDDT_Ana.Start();               

                #endregion

            }
            Cursor = Cursors.Default;
        }

        private void rdb_ngoaitru_CheckedChanged(object sender, EventArgs e)
        {
            this.f_load_maubaocao();
        }
        

        private void f_set_tongtamung_tonggiuong(DataTable dtttrv, string fieldTongHoanUng, string fieldTongTamUng, string fieldGiuong, string fieldGiuongDV, string fieldDoiTuong)
        {
            //lấy tổng tiền đã tạm ứng của bệnh nhân(bao gồm đã hoàn trả biên lai).
            //lấy tổng tiền đã hoàn biên lai tạm ứng của bn.
            if (fieldTongHoanUng == "" && fieldTongTamUng == ""
                && fieldGiuong == "" && fieldGiuongDV == "" && fieldDoiTuong == "") return;
            DateTime date1;
            DateTime date2;
            decimal tongtamung = 0;
            sttAna = 0;
            string sq = "";
            string sq2 = "";
            string sq3 = "";
            string sq4 = "";
            string sq5 = "";
            string my = "";
            string smy = "";
            string madt = m.iTunguyen.ToString();//ma doi tuong dich vu
            for (int i = 0; i < dtttrv.Rows.Count; i++)
            {
                {
                    date1 = new DateTime(Convert.ToInt16(dtttrv.Rows[i]["ngayvao"].ToString().Substring(6, 4))
                        , Convert.ToInt16(dtttrv.Rows[i]["ngayvao"].ToString().Substring(3, 2))
                        , 1);
                    date2 = new DateTime(Convert.ToInt16(dtttrv.Rows[i]["ngay"].ToString().Substring(6, 4))
                        , Convert.ToInt16(dtttrv.Rows[i]["ngay"].ToString().Substring(3, 2))
                        , 1);

                    sq = "";//tong tam ung
                    sq2 = "";//tong hoan tam ung
                    sq3 = "";//tong giuong dich vu
                    sq4 = "";//tong giuong chenh lech
                    sq5 = "";//doi tuong benh nhan
                    smy = "";
                    while (date1 <= date2)
                    {
                        my = date1.ToString("MMyy");
                        if (m.bmMmyy(my))
                        {
                            sq += "select a.mabn,to_char(a.maql)  maql,to_char(a.mavaovien) mavaovien,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_tamung a"
                                + " left join " + m.user + my + ".v_hoantra b  on a.quyenso=b.quyenso and a.sobienlai=b.sobienlai"
                                + " where to_char(a.maql)=" + dtttrv.Rows[i]["maql"].ToString()
                                + " and b.id is null "
                                + " group by a.mabn,a.maql,a.mavaovien" + "#";

                            sq2 += "select a.mabn,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_hoantra a"
                                + " inner join " + m.user + my + ".v_tamung b on a.quyenso=b.quyenso and a.sobienlai=b.sobienlai"
                                + " where a.mabn='" + dtttrv.Rows[i]["mabn"].ToString() + "'"
                                + " and to_date('" + dtttrv.Rows[i]["ngayvao"].ToString()
                                + "','dd/mm/yyyy')<=to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"

                                //+" and to_date('"+dtttrv.Rows[i]["ngayra"].ToString()
                                //+"','dd/mm/yyyy')>=to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"

                                + " group by a.mabn" + "#";
                            sq3 += "select a.id,b2.stt,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_ttrvct a"
                                + " inner join v_giavp b on a.mavp=b.id"
                                + " inner join v_loaivp b1 on b.id_loai=b1.id"
                                + " inner join v_nhomvp b2 on b1.id_nhom=b2.ma"
                                + " inner join v_nhombhyt b3 on b2.idnhombhyt=b3.id"
                                + " inner join v_nhombhyt_medisoft b4 on b3.idnhombhytmedisoft=b4.id"
                                + " where a.id=" + dtttrv.Rows[i]["id"].ToString()
                                + " and a.madoituong in(" + madt + ")"
                                + " and b4.id=11"//tien giuong
                                + " group by a.id,b2.stt" + "#";
                            sq4 += "select a.id,b2.stt,sum(a.sotien-a.bhyttra) sotien "
                                + " from " + m.user + my + ".v_ttrvct a"
                                + " inner join v_giavp b on a.mavp=b.id"
                                + " inner join v_loaivp b1 on b.id_loai=b1.id"
                                + " inner join v_nhomvp b2 on b1.id_nhom=b2.ma"
                                + " inner join v_nhombhyt b3 on b2.idnhombhyt=b3.id"
                                + " inner join v_nhombhyt_medisoft b4 on b3.idnhombhytmedisoft=b4.id"
                                + " where a.id=" + dtttrv.Rows[i]["id"].ToString()
                                + " and a.madoituong not in(" + madt + ")"//thuoc doi tuong bhyt
                                + " and b4.id=11"//tien giuong
                                + " group by a.id,b2.stt" + "#";
                            if (dtttrv.Rows[i]["loaibn"].ToString() == "1"
                                || dtttrv.Rows[i]["loaibn"].ToString() == "2")
                            {
                                smy = "";

                            }
                            else
                                smy = my;
                            sq5 += "select mabn,doituong "
                                + " from " + m.user + smy + ".benhandt a "
                                + " inner join doituong b on a.madoituong=b.madoituong"
                                + " where mabn='" + dtttrv.Rows[i]["mabn"].ToString() + "'"
                                + " and to_char(maql)=" + dtttrv.Rows[i]["maql"].ToString()
                                + "#";
                        }
                        date1 = date1.AddMonths(1);
                    }
                    sq = sq.Trim('#').Replace("#", " union all ");
                    sq = "select to_char(maql) as maql,sum(sotien) tongtamung from (" + sq + ") group by maql";

                    sq2 = sq2.Trim('#').Replace("#", " union all ");
                    sq2 = "select mabn,sum(sotien) tonghoan from (" + sq2 + ") group by mabn";

                    sq3 = sq3.Trim('#').Replace("#", " union all ");
                    sq3 = "select to_char(id) id,stt,sum(sotien) tonggiuongdv from (" + sq3 + ") group by id,stt";

                    sq4 = sq4.Trim('#').Replace("#", " union all ");
                    sq4 = "select to_char(id) id,stt,sum(sotien) tonggiuongchenhlech from (" + sq4 + ") group by id,stt";

                    sq5 = sq5.Trim('#').Replace("#", " union all ");
                    sq5 = "select doituong from (" + sq5 + ")";

                    if (fieldTongTamUng != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq).Tables[0].Rows[0][1].ToString());
                            dtttrv.Rows[i][fieldTongTamUng] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldTongHoanUng != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq2).Tables[0].Rows[0][1].ToString());
                            dtttrv.Rows[i][fieldTongHoanUng] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldGiuong != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq4).Tables[0].Rows[0][2].ToString());
                            dtttrv.Rows[i][fieldGiuong] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldGiuongDV != "")
                    {
                        DataTable dttam = m.get_data(sq3).Tables[0];

                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(dttam.Rows[0][2].ToString());
                            dtttrv.Rows[i][fieldGiuongDV] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldDoiTuong != "")
                    {

                        try
                        {

                            dtttrv.Rows[i][fieldDoiTuong] = m.get_data(sq5).Tables[0].Rows[0][0].ToString();
                        }
                        catch { }
                    }
                    //Insert vào ana
                    try
                    {
                        string st_loaiphieu, st_maso, st_hoten, st_nguothu, st_tenkp, st_sobienlai, st_doituong;
                        string sSoHoaDon = "";
                        try
                        {
                            DataRow[] arrdr = dsHDDT.Tables[0].Select("sochungtu='" + dtttrv.Rows[i]["sochungtu"].ToString() + "'");
                            sSoHoaDon = arrdr[0]["sohoadon"].ToString();
                            st_loaiphieu = arrdr[0]["tencn"].ToString();
                            st_maso = arrdr[0]["mach_cn"].ToString();
                        }
                        catch
                        {
                            if (dtttrv.Rows[i]["loaibn"].ToString() == "1") //BN nội trú
                            {
                                st_loaiphieu = "III. DOANH THU NỘI TRÚ"; 
                                st_maso = "01.03"; 
                            }
                            else
                            {
                                st_loaiphieu = "II. DOANH THU NGOẠI TRÚ"; 
                                st_maso = "01.02"; 
                            }
                        }


                        st_hoten = dtttrv.Rows[i]["hoten"].ToString();
                        st_hoten = m.Replace_ErrorFont(st_hoten);

                        st_nguothu = dtttrv.Rows[i]["nguoithu"].ToString();
                        st_nguothu = m.Replace_ErrorFont(st_nguothu).ToUpper();
                        if (st_nguothu.ToUpper() == "NGUYỄN THU GIANG")
                            st_nguothu = "NGUYỄN THỊ THU GIANG";

                        st_tenkp = dtttrv.Rows[i]["tenkp"].ToString();
                        st_tenkp = m.Replace_ErrorFont(st_tenkp);

                        st_doituong = dtttrv.Rows[i]["doituong"].ToString();
                        st_doituong = m.Replace_ErrorFont(st_doituong);
                        
                        st_sobienlai = dtttrv.Rows[i]["quyenso"].ToString() + " - " + dtttrv.Rows[i]["sobienlai"].ToString();
                        sttAna++;
                        if (sttAna == 1)
                            sttAna = sttVienphi(sttAna, "BANGDULIEUVIENPHI_TAM", dtttrv.Rows[i]["ngay"].ToString());

                        m.upd_ana_BANGDULIEUVIENPHI_API(st_loaiphieu, st_maso, dtttrv.Rows[i]["ngay"].ToString(), "", st_nguothu, st_tenkp, dtttrv.Rows[i]["mabn"].ToString(),
                            st_hoten, dtttrv.Rows[i]["namsinh"].ToString(), st_doituong, dtttrv.Rows[i]["quyenso"].ToString(), dtttrv.Rows[i]["sobienlai"].ToString(), st_sobienlai,
                            decimal.Parse(dtttrv.Rows[i]["sotien"].ToString()), decimal.Parse(dtttrv.Rows[i]["bhyt"].ToString()), tongtamung, decimal.Parse(dtttrv.Rows[i]["tonghoantamung"].ToString()), decimal.Parse(dtttrv.Rows[i]["mien"].ToString()), decimal.Parse(dtttrv.Rows[i]["thucthu"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL7"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL11"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL2"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL3"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL4"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL6"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL1"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL10"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL8"].ToString()), 0, decimal.Parse(dtttrv.Rows[i]["LLLL12"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL13"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL5"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL14"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL0"].ToString()), dtttrv.Rows[i]["ngay"].ToString(), sttAna, "", sSoHoaDon, "");
                    }
                    catch { }
                    //end ana
                }
            }
        }
        private void f_set_tongtamung_tonggiuong(DataTable dtttrv, DataTable dtAna, string fieldTongHoanUng, string fieldTongTamUng, string fieldGiuong, string fieldGiuongDV, string fieldDoiTuong)
        {
            //lấy tổng tiền đã tạm ứng của bệnh nhân(bao gồm đã hoàn trả biên lai).
            //lấy tổng tiền đã hoàn biên lai tạm ứng của bn.
            if (fieldTongHoanUng == "" && fieldTongTamUng == ""
                && fieldGiuong == "" && fieldGiuongDV == "" && fieldDoiTuong == "") return;
            DateTime date1;
            DateTime date2;
            decimal tongtamung = 0;
            sttAna = 0;
            string sq = "";
            string sq2 = "";
            string sq3 = "";
            string sq4 = "";
            string sq5 = "";
            string my = "";
            string smy = "";
            string madt = m.iTunguyen.ToString();//ma doi tuong dich vu
            for (int i = 0; i < dtttrv.Rows.Count; i++)
            {
                //BAT DAU
                DataRow[] drAna = new DataRow[0];
                try
                {
                    drAna = dtAna.Select("sochungtu='" + dtttrv.Rows[i]["sochungtu"].ToString() + "'");
                }
                catch { }
                if (drAna.Length == 0)
                {

                    date1 = new DateTime(Convert.ToInt16(dtttrv.Rows[i]["ngayvao"].ToString().Substring(6, 4))
                        , Convert.ToInt16(dtttrv.Rows[i]["ngayvao"].ToString().Substring(3, 2))
                        , 1);
                    date2 = new DateTime(Convert.ToInt16(dtttrv.Rows[i]["ngay"].ToString().Substring(6, 4))
                        , Convert.ToInt16(dtttrv.Rows[i]["ngay"].ToString().Substring(3, 2))
                        , 1);

                    sq = "";//tong tam ung
                    sq2 = "";//tong hoan tam ung
                    sq3 = "";//tong giuong dich vu
                    sq4 = "";//tong giuong chenh lech
                    sq5 = "";//doi tuong benh nhan
                    smy = "";
                    while (date1 <= date2)
                    {
                        my = date1.ToString("MMyy");
                        if (m.bmMmyy(my))
                        {
                            sq += "select a.mabn,to_char(a.maql)  maql,to_char(a.mavaovien) mavaovien,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_tamung a"
                                + " left join " + m.user + my + ".v_hoantra b  on a.quyenso=b.quyenso and a.sobienlai=b.sobienlai"
                                + " where to_char(a.maql)=" + dtttrv.Rows[i]["maql"].ToString()
                                + " and b.id is null "
                                + " group by a.mabn,a.maql,a.mavaovien" + "#";

                            sq2 += "select a.mabn,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_hoantra a"
                                + " inner join " + m.user + my + ".v_tamung b on a.quyenso=b.quyenso and a.sobienlai=b.sobienlai"
                                + " where a.mabn='" + dtttrv.Rows[i]["mabn"].ToString() + "'"
                                + " and to_date('" + dtttrv.Rows[i]["ngayvao"].ToString()
                                + "','dd/mm/yyyy')<=to_date(to_char(a.ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"

                                //+" and to_date('"+dtttrv.Rows[i]["ngayra"].ToString()
                                //+"','dd/mm/yyyy')>=to_date(to_char(ngay,'dd/mm/yyyy'),'dd/mm/yyyy')"

                                + " group by a.mabn" + "#";
                            sq3 += "select a.id,b2.stt,sum(a.sotien) sotien "
                                + " from " + m.user + my + ".v_ttrvct a"
                                + " inner join v_giavp b on a.mavp=b.id"
                                + " inner join v_loaivp b1 on b.id_loai=b1.id"
                                + " inner join v_nhomvp b2 on b1.id_nhom=b2.ma"
                                + " inner join v_nhombhyt b3 on b2.idnhombhyt=b3.id"
                                + " inner join v_nhombhyt_medisoft b4 on b3.idnhombhytmedisoft=b4.id"
                                + " where a.id=" + dtttrv.Rows[i]["id"].ToString()
                                + " and a.madoituong in(" + madt + ")"
                                + " and b4.id=11"//tien giuong
                                + " group by a.id,b2.stt" + "#";
                            sq4 += "select a.id,b2.stt,sum(a.sotien-a.bhyttra) sotien "
                                + " from " + m.user + my + ".v_ttrvct a"
                                + " inner join v_giavp b on a.mavp=b.id"
                                + " inner join v_loaivp b1 on b.id_loai=b1.id"
                                + " inner join v_nhomvp b2 on b1.id_nhom=b2.ma"
                                + " inner join v_nhombhyt b3 on b2.idnhombhyt=b3.id"
                                + " inner join v_nhombhyt_medisoft b4 on b3.idnhombhytmedisoft=b4.id"
                                + " where a.id=" + dtttrv.Rows[i]["id"].ToString()
                                + " and a.madoituong not in(" + madt + ")"//thuoc doi tuong bhyt
                                + " and b4.id=11"//tien giuong
                                + " group by a.id,b2.stt" + "#";
                            if (dtttrv.Rows[i]["loaibn"].ToString() == "1"
                                || dtttrv.Rows[i]["loaibn"].ToString() == "2")
                            {
                                smy = "";

                            }
                            else
                                smy = my;
                            sq5 += "select mabn,doituong "
                                + " from " + m.user + smy + ".benhandt a "
                                + " inner join doituong b on a.madoituong=b.madoituong"
                                + " where mabn='" + dtttrv.Rows[i]["mabn"].ToString() + "'"
                                + " and to_char(maql)=" + dtttrv.Rows[i]["maql"].ToString()
                                + "#";
                        }
                        date1 = date1.AddMonths(1);
                    }
                    sq = sq.Trim('#').Replace("#", " union all ");
                    sq = "select to_char(maql) as maql,sum(sotien) tongtamung from (" + sq + ") group by maql";

                    sq2 = sq2.Trim('#').Replace("#", " union all ");
                    sq2 = "select mabn,sum(sotien) tonghoan from (" + sq2 + ") group by mabn";

                    sq3 = sq3.Trim('#').Replace("#", " union all ");
                    sq3 = "select to_char(id) id,stt,sum(sotien) tonggiuongdv from (" + sq3 + ") group by id,stt";

                    sq4 = sq4.Trim('#').Replace("#", " union all ");
                    sq4 = "select to_char(id) id,stt,sum(sotien) tonggiuongchenhlech from (" + sq4 + ") group by id,stt";

                    sq5 = sq5.Trim('#').Replace("#", " union all ");
                    sq5 = "select doituong from (" + sq5 + ")";

                    if (fieldTongTamUng != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq).Tables[0].Rows[0][1].ToString());
                            dtttrv.Rows[i][fieldTongTamUng] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldTongHoanUng != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq2).Tables[0].Rows[0][1].ToString());
                            dtttrv.Rows[i][fieldTongHoanUng] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldGiuong != "")
                    {
                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(m.get_data(sq4).Tables[0].Rows[0][2].ToString());
                            dtttrv.Rows[i][fieldGiuong] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldGiuongDV != "")
                    {
                        DataTable dttam = m.get_data(sq3).Tables[0];

                        tongtamung = 0;
                        try
                        {
                            tongtamung = Convert.ToDecimal(dttam.Rows[0][2].ToString());
                            dtttrv.Rows[i][fieldGiuongDV] = tongtamung.ToString();
                        }
                        catch { }
                    }
                    if (fieldDoiTuong != "")
                    {

                        try
                        {

                            dtttrv.Rows[i][fieldDoiTuong] = m.get_data(sq5).Tables[0].Rows[0][0].ToString();
                        }
                        catch { }
                    }
                    //Insert vào ana
                    try
                    {
                        string st_loaiphieu, st_maso, st_hoten, st_nguothu, st_tenkp, st_sobienlai, st_doituong;
                        string sSoHoaDon = "";

                        if (dtttrv.Rows[i]["loaibn"].ToString() == "1") //BN nội trú
                        {
                            st_loaiphieu = "III. DOANH THU NỘI TRÚ"; //II. DOANH THU NGOẠI TRÚ hoặc III. DOANH THU NỘI TRÚ
                            st_maso = "01.03"; //01.02 hoặc 01.03
                        }
                        else
                        {
                            st_loaiphieu = "II. DOANH THU NGOẠI TRÚ"; //II. DOANH THU NGOẠI TRÚ hoặc III. DOANH THU NỘI TRÚ
                            st_maso = "01.02"; //01.02 hoặc 01.03
                        }
                        st_hoten = dtttrv.Rows[i]["hoten"].ToString();
                        st_hoten = m.Replace_ErrorFont(st_hoten);

                        st_nguothu = dtttrv.Rows[i]["nguoithu"].ToString();
                        st_nguothu = m.Replace_ErrorFont(st_nguothu).ToUpper();

                        st_tenkp = dtttrv.Rows[i]["tenkp"].ToString();
                        st_tenkp = m.Replace_ErrorFont(st_tenkp);

                        st_doituong = dtttrv.Rows[i]["doituong"].ToString();
                        st_doituong = m.Replace_ErrorFont(st_doituong);


                        st_sobienlai = dtttrv.Rows[i]["quyenso"].ToString() + " - " + dtttrv.Rows[i]["sobienlai"].ToString();
                        sttAna++;
                        if (sttAna == 1)
                            sttAna = sttVienphi(sttAna, "BANGDULIEUVIENPHI_TAM", dtttrv.Rows[i]["ngay"].ToString());

                        m.upd_ana_BANGDULIEUVIENPHI_API(st_loaiphieu, st_maso, dtttrv.Rows[i]["ngay"].ToString(), "", st_nguothu, st_tenkp, dtttrv.Rows[i]["mabn"].ToString(),
                            st_hoten, dtttrv.Rows[i]["namsinh"].ToString(), st_doituong, dtttrv.Rows[i]["quyenso"].ToString(), dtttrv.Rows[i]["sobienlai"].ToString(), st_sobienlai,
                            decimal.Parse(dtttrv.Rows[i]["sotien"].ToString()), decimal.Parse(dtttrv.Rows[i]["bhyt"].ToString()), tongtamung, decimal.Parse(dtttrv.Rows[i]["tonghoantamung"].ToString()), decimal.Parse(dtttrv.Rows[i]["mien"].ToString()), decimal.Parse(dtttrv.Rows[i]["thucthu"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL7"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL11"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL2"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL3"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL4"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL6"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL1"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL10"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL8"].ToString()), 0, decimal.Parse(dtttrv.Rows[i]["LLLL12"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL13"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL5"].ToString()),
                            decimal.Parse(dtttrv.Rows[i]["LLLL14"].ToString()), decimal.Parse(dtttrv.Rows[i]["LLLL0"].ToString()), dtttrv.Rows[i]["ngay"].ToString(), sttAna, "", sSoHoaDon, "");
                    }
                    catch { }
                    //end ana
                }
                //KET THUC

            }
        }
        private decimal sttVienphi (decimal stt,string table,string ngay)
        {
            try
            {
                if (stt == 1)
                {
                    dtmaxSttAna = hddt.dataSetFromSql("mssql", "select max(stt) stt from  " + table + " where convert(varchar, ngayct, 103) = '" + ngay + "'").Tables[0];

                    stt = decimal.Parse(dtmaxSttAna.Rows[0]["stt"].ToString()) + 1;
                    
                }
            }
            catch { }
            return stt;
        }
        private void tungay_ValueChanged(object sender, EventArgs e)
        {
            if (tungay.Value > denngay.Value)
            {
                tungay.Value = denngay.Value;
            }
            //load_grid();
        }
        private void denngay_ValueChanged(object sender, EventArgs e)
        {
            if (denngay.Value < tungay.Value)
            {
                denngay.Value = tungay.Value;
            }
            //load_grid();
        }
        private void cbLoaiVP_SelectedIndexChanged(object sender, EventArgs e)
        {
            //load_grid();
        }
        private void tim_TextChanged(object sender, EventArgs e)
        {
            if (this.ActiveControl == tim)
            {
                CurrencyManager cm = (CurrencyManager)BindingContext[dtg.DataSource];
                DataView dv = (DataView)cm.List;
                dv.RowFilter = "tenkhachhang like '%" + tim.Text.Trim() + "%' or makhachhang like '%" + tim.Text.Trim() + "%'";
                dtg.DataSource = dv;
            }	
        }
        private void tim_MouseClick(object sender, MouseEventArgs e)
        {
            if(tim.Text=="Tìm kiếm") tim.Text = "";
        }
        private void tim_Validated(object sender, EventArgs e)
        {
            if (tim.Text == "") tim.Text = "Tìm kiếm";
        }
        private void bStart_Click(object sender, EventArgs e)
        {

            m = new AccessDataApi(ipApi);
            bHDDT = false;
            bAna = true;
            timer_insert_Tick(null, null);
            nbPhut.Enabled = cbMMYY.Enabled = bStart.Enabled = bClose.Enabled = false;
            bPause.Enabled = true;
            timer_insert.Interval = int.Parse((nbPhut.Value).ToString()) * interval;
            timer_insert.Start();
        }
        private void bHDDTStart_Click(object sender, EventArgs e)
        {
            m = new AccessDataApi(ipApi);
            bAna = false;
            bHDDT = true;
            timer_insert_Tick(null, null);
            nbPhut.Enabled = cbMMYY.Enabled = bStart.Enabled = bClose.Enabled = false;
            bPause.Enabled = true;
            timer_insert.Interval = int.Parse((nbPhut.Value).ToString()) * interval;
            timer_insert.Start();
        }
        private void bPause_Click(object sender, EventArgs e)
        {
            //nbPhut.Enabled = cbMMYY.Enabled = bStart.Enabled = bClose.Enabled = true;
            //bPause.Enabled = false;
            //timer_insert.Stop();
            if (tMediSoft_HDDT_TT != null && tMediSoft_HDDT_TT.IsAlive == true)
                tMediSoft_HDDT_TT.Suspend();
            if (tMediSoft_HDDT_TTRV !=null && tMediSoft_HDDT_TTRV.IsAlive == true)
                tMediSoft_HDDT_TTRV.Suspend();
            if (tHDDT_Ana != null && tHDDT_Ana.IsAlive == true)
                tHDDT_Ana.Suspend();
            if (tHDDT_Ana != null && tHDDT_Ana_KyQuy.IsAlive == true)
                tHDDT_Ana_KyQuy.Suspend();
            nbPhut.Enabled = cbMMYY.Enabled = bKetxuat.Enabled = bStart.Enabled = bClose.Enabled = true;
            bPause.Enabled = false;
            bHDDT = bAna = false;
            timer_insert.Stop();
            record.Text = "Dừng";
        }
        private void bClose_Click(object sender, EventArgs e)
        {
            timer_insert.Stop();
            this.Close();
        }
    }
}