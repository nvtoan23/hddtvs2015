using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using libSin;

namespace HDDT
{
    public partial class frmMain : Form
    {
        private AccessData m = new AccessData();
        private DataSet ds = new DataSet();
        private string sql = "";
        const int TamUng = 1, ThuTrucTiep = 2, TTRV = 3, bhNgoaiTru = 4;
        private DataGridView dtg;
        private Button bStart;
        private Timer timer_insert;
        private IContainer components;
        private Button bClose;
        public int interval = 1000;
    
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            load_grid();
            dtg.DataSource = ds.Tables[0];
        }
        private void load_grid()
        {
            sql = "select sohoso,makhachhang,tenkhachhang,phai,namsinh,diachi,ngayhd from hddt.tonghop";
            ds = m.get_data(sql);
        }

        private void bStart_Click(object sender, EventArgs e)
        {
            timer_insert.Start();
            //Insert thu truc tiep
            //sql = "INSERT INTO hddt.tonghop(sohoso,loai,quyenso,sobienlai,makhachhang,tenkhachhang,namsinh,phai,diachi,ngayhd,id,mmyy)";
            //sql += " select a.mavaovien as sohoso,'tt' as loai,a.quyenso,a.sobienlai, a.mabn as makhachhang,b.hoten as tenkhachhang,b.namsinh,b.phai,";
            //sql += " b.sonha || ' ' || b.thon || ' '|| c.tenpxa || ' ' || d.tenquan || ' ' || e.tentt as diachi,a.ngay as ngayhd,a.id,a.mmyy";
            //sql += " from medibv0516.v_vienphill a inner join medibv.btdbn b on a.mabn=b.mabn inner join medibv.btdpxa c on c.maphuongxa=b.maphuongxa";
            //sql += " inner join medibv.btdquan d on d.maqu=b.maqu inner join medibv.btdtt e on e.matt=b.matt";
            
            //End thu truc tiep
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.dtg = new System.Windows.Forms.DataGridView();
            this.bStart = new System.Windows.Forms.Button();
            this.bClose = new System.Windows.Forms.Button();
            this.timer_insert = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dtg)).BeginInit();
            this.SuspendLayout();
            // 
            // dtg
            // 
            this.dtg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dtg.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dtg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dtg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtg.Location = new System.Drawing.Point(12, 38);
            this.dtg.Name = "dtg";
            this.dtg.RowHeadersVisible = false;
            this.dtg.Size = new System.Drawing.Size(654, 331);
            this.dtg.TabIndex = 0;
            // 
            // bStart
            // 
            this.bStart.Location = new System.Drawing.Point(510, 386);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(75, 23);
            this.bStart.TabIndex = 1;
            this.bStart.Text = "Chạy";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // bClose
            // 
            this.bClose.Location = new System.Drawing.Point(591, 386);
            this.bClose.Name = "bClose";
            this.bClose.Size = new System.Drawing.Size(75, 23);
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
            // frmMain
            // 
            this.ClientSize = new System.Drawing.Size(678, 421);
            this.Controls.Add(this.bClose);
            this.Controls.Add(this.bStart);
            this.Controls.Add(this.dtg);
            this.Name = "frmMain";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtg)).EndInit();
            this.ResumeLayout(false);

        }

        private void bClose_Click(object sender, EventArgs e)
        {
            timer_insert.Stop();
            //this.Close();
        }

        private void timer_insert_Tick(object sender, EventArgs e)
        {
            
        }
    }
}