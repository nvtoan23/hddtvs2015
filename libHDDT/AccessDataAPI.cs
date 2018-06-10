using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace libHDDT
{
  public  class AccessDataAPI
    {
        public AccessDataAPI(string filename)
        {
            var urlfilename = Path.Combine("data", filename);
            if (!File.Exists(urlfilename))
            {
                return;
            }
            var contenttxt = File.ReadAllText(urlfilename);
            try
            {
            dataApi =    JsonConvert.DeserializeObject<DataAPI>(contenttxt);
                loadClient();
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        public AccessDataAPI()
        {
            var urlfilename = Path.Combine("data", "data.json");
            if (!File.Exists(urlfilename))
            {
                return;
            }
            var contenttxt = File.ReadAllText(urlfilename);
            try
            {
                dataApi = JsonConvert.DeserializeObject<DataAPI>(contenttxt);
                loadClient();
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        public   DataAPI dataApi { set; get; }
        private   HttpClient client { set; get; }
        public JObject truyvanSQL(string typeSql, string sql)
        {
            var httprs = client.PostAsync(string.Format(dataApi.urlbase,typeSql), new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                return JObject.Parse(tjson);
            }
            return null;

        }
        public JObject truyvanSQL(string typeSql, string keysql,  params object[] obparams)
        {
            if (!dataApi.querydata.ContainsKey(keysql))
            {
                return null;
            }
            string sql = dataApi.querydata[keysql];
            sql = string.Format(sql, obparams);
            
            var httprs = client.PostAsync(string.Format(dataApi.urlbase, typeSql), new StringContent(JsonConvert.SerializeObject(new { sql = sql }), Encoding.UTF8, "application/json")).Result;
            if (httprs.IsSuccessStatusCode)
            {
                var tjson = httprs.Content.ReadAsStringAsync().Result;
                return JObject.Parse(tjson);
            }
            return null;

        }
        public bool thucThiSql(string typeSql, string sql)
        {
            var obtruyvan = truyvanSQL(typeSql, sql);
            if(obtruyvan == null)
            {
                return false;
            }
            else
            {
                
                if ((bool)obtruyvan["ok"] && (int)obtruyvan["data"] > 0)
                    return true;
                else
                    return false;
            }         

        }
        public bool thucThiSql(string typeSql, string keysql, params object[] obparams)
        {
            var obtruyvan = truyvanSQL(typeSql, keysql,obparams);
            if (obtruyvan == null)
            {
                return false;
            }
            else
            {

                if ((bool)obtruyvan["ok"] && (int)obtruyvan["data"] > 0)
                    return true;
                else
                    return false;
            }

        }
        public DataSet dataSetFromSql(string typeSql, string sql)
        {
            var obdata = truyvanSQL(typeSql, sql);
            if (obdata == null)
                return null;
            if ((bool)obdata["ok"])
            {
                var arjson = obdata["data"];
                var objson = JsonConvert.SerializeObject(new { data = arjson });
                return JsonConvert.DeserializeObject<DataSet>(objson);
            }
            return null;
        }
        public DataSet dataSetFromSql(string typeSql, string keysql, params object[] obparams)
        {
            var obdata = truyvanSQL(typeSql, keysql, obparams);
            if (obdata == null)
                return null;
            if ((bool)obdata["ok"])
            {
                var arjson = obdata["data"];
                var objson = JsonConvert.SerializeObject(new { data = arjson });
                return JsonConvert.DeserializeObject<DataSet>(objson);
            }
            return null;
        }
        private void loadClient()
        {
            client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

        }
    }
}
