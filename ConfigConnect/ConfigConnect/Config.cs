namespace ConfigConnect
{
    using System;
    using System.Xml;

    public class Config
    {
        private string _sConn = "Data Source=MEDISOFT;user id=MEDIBV;password=MEDIBV";
        private string _userid = "medibv";
        private string service_name = "medisoft";

        public Config()
        {
            if (this.Maincode("Con") != "")
            {
                this._sConn = this.Maincode("Con");
            }
            this._userid = this._sConn.Substring(this._sConn.LastIndexOf("=") + 1).Trim();
            this.service_name = this._sConn.Substring(this._sConn.IndexOf("=") + 1, (this._sConn.IndexOf(";") - 1) - this._sConn.IndexOf("=")).Trim();
        }

        private string Maincode(string sql)
        {
            XmlDocument document = new XmlDocument();
            document.Load(@"..\..\..\xml\maincode.xml");
            return document.GetElementsByTagName(sql).Item(0).InnerText;
        }

        public string pDiaChibv
        {
            get
            {
                return this.Maincode("Diachi");
            }
        }

        public string pMaBV
        {
            get
            {
                string str = this.Maincode("Mabv");
                if (str == "")
                {
                    str = "701.1.01";
                }
                return str;
            }
        }

        public string pServiceName
        {
            get
            {
                return this.service_name;
            }
        }

        public string pSoYTe
        {
            get
            {
                return this.Maincode("Syte");
            }
        }

        public string pStringConnect
        {
            get
            {
                return this._sConn;
            }
        }

        public string pTenbv
        {
            get
            {
                return this.Maincode("Tenbv");
            }
        }

        public string pUser
        {
            get
            {
                return this._userid;
            }
        }
    }
}

