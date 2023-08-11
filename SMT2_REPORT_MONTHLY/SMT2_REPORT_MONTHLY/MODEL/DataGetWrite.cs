using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.MODEL
{
    public class DataDD
    {
        public string Model { get; set; }
        public long Qty { get; set; }
        public string KH { get; set; }

        public string Mat { get; set; }

        public long QtyError { get; set; }

        public DataDD()
        {

        }

        public DataDD(DataDD s)
        {
            this.Model = s.Model;
            this.Qty = s.Qty;
            this.KH = s.KH;
            this.Mat = s.Mat;
            
        }
        
    }
    public class DataError
    {
        public string Model { get; set; }
        public string KH { get; set; }
        public string Mat { get; set; }
        public long QtyError { get; set; }
        public string TypeItem { get; set; }
        public string Content { get; set; }

        public override string ToString()
        {
            return $"{this.Model}-{this.KH}-{this.Mat}-{this.QtyError}-{this.TypeItem}-{this.Content}";
        }

        public DataError()
        {

        }

        public DataError(DataError s)
        {
            this.Model = s.Model;
            this.KH = s.KH;
            this.Mat = s.Mat;
            this.QtyError = s.QtyError;
            this.TypeItem = s.TypeItem;
            this.Content = s.Content;
        }
    }
    public class DataError2
    {
        public string Model { get; set; }
        public long QtyError { get; set; }
        public string Content { get; set; }

        
    }
    public class DataError3
    {
        public string Model { get; set; }
        public string TypeItem { get; set; }
        public long QtyError { get; set; }
        public string Content { get; set; }


    }
    public class DataError4
    {
        public string Content { get; set; }
        public long QtyError { get; set; }
    }


    public class DataWrite
    {
        public string model { get; set; }
        public long qty { get; set; }
        public long qtyErr { get; set; }
        public string modelFirst { get; set; }

        public DataWrite()
        {

        }
        public DataWrite(DataWrite s)
        {
            this.model = s.model;
            this.qty = s.qty;
            this.qtyErr = s.qtyErr;
            this.modelFirst = s.modelFirst;
         
        }
    }
}
