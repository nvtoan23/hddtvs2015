using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using LibMedi;

namespace HDDT
{
    public partial class Form1 : Form
    {
        private AccessDataApi m = new AccessDataApi();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sql = "select * from btdkp_bv";
            DataSet ds = m.get_data(sql);
            dataGridView1.DataSource = ds.Tables[0];
        }
    }
}