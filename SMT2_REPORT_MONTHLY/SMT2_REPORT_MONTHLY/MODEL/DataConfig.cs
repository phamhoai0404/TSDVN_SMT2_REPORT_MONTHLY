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

        public static void GetConfig(ref DataConfigDD config)
        {
            config.Mat.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_MAT"];
            config.Model.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_MODEL"];
            config.KH.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_KH"];
            config.Qty.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_QTY"];

            config.Mat.GetIndexColumn();
            config.Model.GetIndexColumn();
            config.KH.GetIndexColumn();
            config.Qty.GetIndexColumn();

            int maxVariable = Math.Max(config.Model.Index, Math.Max(config.KH.Index, Math.Max(config.Mat.Index, config.Qty.Index)));
            config.ColLast = MyFunction1.ConvertNumberToName(maxVariable + 1);//vi no dang bi tru di 1 nen la
        }

    }
    public class Address
    {
        public int Index { get; set; }
        public string ColName { get; set; }
        public Address()
        {

        }
        public Address(Address s)
        {
            this.Index = s.Index;
            this.ColName = s.ColName;
        }
        public void GetIndexColumn()
        {
            this.Index = MyFunction1.ConvertNameToNumer(this.ColName) - 1;//Tru 1 vi de tinh toan se bat dau = 0
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

        }
        public static void GetConfig(ref DataConfigLoi config)
        {
            config.Model.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_MODEL"];
            config.KH.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_KH"];
            config.Mat.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_MAT"];
            config.QtyError.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_QTYERROR"];
            config.TypeItem.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_TYPEITEM"];
            config.Content.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_CONTENT"];

            config.Model.GetIndexColumn();
            config.KH.GetIndexColumn();
            config.Mat.GetIndexColumn();
            config.QtyError.GetIndexColumn();
            config.TypeItem.GetIndexColumn();
            config.Content.GetIndexColumn();
            
            int maxVariable = Math.Max(config.Model.Index, Math.Max(config.KH.Index, Math.Max(config.Mat.Index, Math.Max(config.QtyError.Index, Math.Max(config.TypeItem.Index, config.Content.Index)))));
            config.ColLast = MyFunction1.ConvertNumberToName(maxVariable + 1);//vi no dang bi tru di 1 nen la
        }
    }

}
