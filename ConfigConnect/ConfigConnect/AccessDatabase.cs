namespace ConfigConnect
{
    using System;
    using System.Data;
    using System.Data.OracleClient;

    public class AccessDatabase
    {
        private Config _config;
        private OracleCommand cmd;
        private OracleConnection con;
        private OracleDataAdapter dest;
        private int iRownum;
        private string sComputer;

        public AccessDatabase()
        {
            this.sComputer = null;
            this.iRownum = 1;
            this._config = new Config();
            this.khoitao();
        }

        public AccessDatabase(Config conf)
        {
            this.sComputer = null;
            this.iRownum = 1;
            this._config = conf;
            this.khoitao();
        }

        public bool f_execute_data(string sql)
        {
            return this.f_execute_data(this._config.pStringConnect, sql, true);
        }

        public bool f_execute_data(string sql, bool bLuuEror)
        {
            return this.f_execute_data(this._config.pStringConnect, sql, bLuuEror);
        }

        public bool f_execute_data(string connect, string sql)
        {
            return this.f_execute_data(connect, sql, true);
        }

        public bool f_execute_data(string connect, string sql, bool bLuuError)
        {
            try
            {
                this.con = new OracleConnection(connect);
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                this.con.Close();
                this.con.Dispose();
                return true;
            }
            catch (OracleException exception)
            {
                if (bLuuError)
                {
                    this.f_upd_error(exception.Message.ToString().Trim() + "-SQL:" + sql, this.sComputer, "?");
                }
                return false;
            }
        }

        public DataSet f_get_data(string sql)
        {
            return this.f_get_data("", sql);
        }

        public DataSet f_get_data(string mmyy, string sql)
        {
            DataSet dataSet = new DataSet();
            try
            {
                if (this.con != null)
                {
                    this.con.Close();
                    this.con.Dispose();
                }
                this.con = new OracleConnection(this.f_getConnect(mmyy, true));
                this.con.Open();
                this.cmd = new OracleCommand(sql, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.dest = new OracleDataAdapter(this.cmd);
                this.dest.Fill(dataSet);
                this.cmd.Dispose();
                this.con.Close();
                this.con.Dispose();
            }
            catch (OracleException exception)
            {
                this.f_upd_error(sql + "; " + exception.Message.ToString().Trim(), this.sComputer, "?");
            }
            return dataSet;
        }

        public string f_getConnect(string mmyy, bool bduoc)
        {
            if (mmyy != "")
            {
                mmyy = (bduoc ? "d" : "") + mmyy;
                return ("Data Source=" + this._config.pServiceName + ";user id=" + this._config.pUser + "d" + mmyy + ";password=" + this._config.pUser + "d" + mmyy);
            }
            return ("Data Source=" + this._config.pServiceName + ";user id=" + this._config.pUser + ";password=" + this._config.pUser);
        }

        public string f_getMMYY(string ngay)
        {
            if (ngay.Length < 10)
            {
                return ngay;
            }
            return (ngay.Substring(3, 2) + ngay.Substring(8, 2));
        }

        public void f_upd_dmcomputer(string m_computer)
        {
            string commandText = "update dmcomputer set computer=:m_computer where computer=:m_computer";
            this.con = new OracleConnection(this._config.pStringConnect);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(commandText, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                int num = this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
                if (num == 0)
                {
                    commandText = "insert into dmcomputer(id,computer,ngayud) values (1,:m_computer,sysdate)";
                    this.cmd = new OracleCommand(commandText, this.con);
                    this.cmd.CommandType = CommandType.Text;
                    this.cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                    this.cmd.ExecuteNonQuery();
                    this.cmd.Dispose();
                }
            }
            catch (OracleException exception)
            {
                this.f_upd_error(exception.Message, this.sComputer, "dmcomputer");
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
        }

        public void f_upd_error(string m_message, string m_computer, string m_table)
        {
            this.con.Close();
            this.con.Dispose();
            string commandText = "insert into error(message,computer,tables,ngayud) values (:m_message,:m_computer,:m_table,sysdate)";
            this.con = new OracleConnection(this._config.pStringConnect);
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(commandText, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("m_message", OracleType.VarChar).Value = m_message;
                this.cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                this.cmd.Parameters.Add("m_table", OracleType.VarChar, 20).Value = m_table;
                this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
            }
            catch
            {
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
        }

        public void f_upd_error(string m_ngay, string m_message, string m_computer, string m_table)
        {
            this.con.Close();
            this.con.Dispose();
            string commandText = "insert into error(message,computer,tables,ngayud) values (:m_message,:m_computer,:m_table,sysdate)";
            this.con = new OracleConnection(this.f_getConnect(this.f_getMMYY(m_ngay), true));
            try
            {
                this.con.Open();
                this.cmd = new OracleCommand(commandText, this.con);
                this.cmd.CommandType = CommandType.Text;
                this.cmd.Parameters.Add("m_message", OracleType.VarChar, 0xfe).Value = m_message;
                this.cmd.Parameters.Add("m_computer", OracleType.VarChar, 20).Value = m_computer;
                this.cmd.Parameters.Add("m_table", OracleType.VarChar, 20).Value = m_table;
                this.cmd.ExecuteNonQuery();
                this.cmd.Dispose();
            }
            catch
            {
            }
            finally
            {
                this.con.Close();
                this.con.Dispose();
            }
        }

        private void khoitao()
        {
            this.sComputer = Environment.MachineName.Trim().ToUpper();
            this.f_upd_dmcomputer(this.sComputer);
            DataRow[] rowArray = this.f_get_data("select rownum,computer from dmcomputer").Tables[0].Select("computer='" + this.sComputer + "'");
            if (rowArray.Length > 0)
            {
                this.iRownum = int.Parse(rowArray[0]["rownum"].ToString());
            }
        }
    }
}

