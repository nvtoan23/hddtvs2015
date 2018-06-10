using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace libHDDT
{
   public class DataAPI
    {        
        public DataAPI()
        {
            medibv = "medibv";
        }
        public string urlbase { set; get; }
        public Dictionary<string,string> querydata { set; get; }
        public string medibv { set; get; } 
    }
}
