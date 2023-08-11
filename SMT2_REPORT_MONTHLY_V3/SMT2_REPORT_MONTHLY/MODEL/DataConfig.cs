using QA_TVN2_REPORT_MONTHLY.FUNCTION;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.MODEL
{
    public class DataConfigDD
    {
        public string pathFile { get; set; }
        public string sheetName { get; set; }
        public Address Model { get; set; }
        public Address Qty { get; set; }
        public Address KH { get; set; }

        public Address Mat { get; set; }

        public string ColLast { get; set; }

        public DataConfigDD()
        {
            this.Mat = new Address();
            this.Model = new Address();
            this.Qty = new Address();
            this.KH = new Address();
        }
    }
   

    public class DataConfigLoi
    {
        public Address Model { get; set; }
        public Address KH { get; set; }
        public Address Mat { get; set; }
        public Address QtyError { get; set; }
        public Address TypeItem { get; set; }
        public Address Content { get; set; }

        public Address Content_Error_KH { get; set; }
        public string ColLast { get; set; }
        public string pathFile { get; set; }
        public string sheetName { get; set; }


        public DataConfigLoi()
        {
            this.Model = new Address();
            this.KH = new Address();
            this.Mat = new Address();
            this.QtyError = new Address();
            this.TypeItem = new Address();
            this.Content = new Address();
            this.Content_Error_KH = new Address();
        }
        
    }

}
