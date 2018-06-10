using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace HDDT
{
    public partial class FormAsynTask : Form
    {
        public delegate void DoWorkProcess(BackgroundWorker sender, DoWorkEventArgs e);
        public event DoWorkProcess DoWork;
        public delegate void WorkerComplete(object sender, RunWorkerCompletedEventArgs e);
        public event WorkerComplete DoWorkCompleted;
        public FormAsynTask()
        {
            InitializeComponent();
        }
        public void StartProcess(object agr)
        {
            backgroundWorker1.RunWorkerAsync(agr);
        }
        public void StartProcess()
        {
            backgroundWorker1.RunWorkerAsync();
        }
        public void ShowAndStart()
        {
            this.Show();
            StartProcess();
        }
        public string ProcessName { set { lb_name.Text = value; } get { return lb_name.Text; } }
        private void FormAsynTask_Load(object sender, EventArgs e)
        {
            lb_status.Text = "";
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (DoWork != null)
                DoWork(backgroundWorker1, e);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            if (e.UserState != null)
                lb_status.Text = e.UserState.ToString();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (DoWorkCompleted != null)
                DoWorkCompleted(this, e);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void FormAsynTask_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(backgroundWorker1.IsBusy)
            {
                e.Cancel = true;
            }
        }
    }
}