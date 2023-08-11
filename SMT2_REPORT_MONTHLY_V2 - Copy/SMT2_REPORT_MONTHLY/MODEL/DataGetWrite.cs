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
        public QtyErrType QtyType { get; set; }
          

        public DataDD()
        {
            this.QtyType = new QtyErrType();
        }

        public DataDD(DataDD s)
        {
            this.Model = s.Model;
            this.Qty = s.Qty;
            this.KH = s.KH;
            this.Mat = s.Mat;
            this.QtyType = new QtyErrType(s.QtyType);
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
        public QtyErrType QtyType { get; set; }
        public string ContentKH { get; set; }
        public override string ToString()
        {
            return $"{this.Model}-{this.KH}-{this.Mat}-{this.QtyError}-{this.TypeItem}-{this.Content}-{this.ContentKH}";
        }

        public DataError()
        {
            this.QtyType = new QtyErrType();
        }

        public DataError(DataError s)
        {
            this.Model = s.Model;
            this.KH = s.KH;
            this.Mat = s.Mat;
            this.QtyError = s.QtyError;
            this.TypeItem = s.TypeItem;
            this.Content = s.Content;

            this.QtyType = new QtyErrType(s.QtyType);
            this.ContentKH = s.ContentKH;
        }
    }

    public class QtyErrType
    {
        public long Qty_HanGia { get; set; }
        public long Qty_SaiVitri { get; set; }
        public long Qty_Kenh { get; set; }
        public long Qty_Baccau { get; set; }
        public long Qty_It_Thiec { get; set; }
        public long Qty_Thieu_LK { get; set; }
        public long Qty_Lat_Nguoc { get; set; }
        public long Qty_Nguoc_Huong { get; set; }
        public long Qty_Nham_LK { get; set; }
        public long Qty_Di_Vat { get; set; }
        public long Qty_Thua_LK { get; set; }
        public long Qty_Bong { get; set; }
        public long Qty_Khac { get; set; }
        public long Qty_Lech { get; set; }
        public long Qty_Vo { get; set; }
        public long Qty_Dung_dung { get; set; }

        public QtyErrType()
        {

        }
        public QtyErrType(QtyErrType s)
        {
            this.Qty_HanGia = s.Qty_HanGia;
            this.Qty_SaiVitri = s.Qty_SaiVitri;
            this.Qty_Kenh = s.Qty_Kenh;
            this.Qty_Baccau = s.Qty_Baccau;
            this.Qty_It_Thiec = s.Qty_It_Thiec;
            this.Qty_Thieu_LK = s.Qty_Thieu_LK;
            this.Qty_Lat_Nguoc = s.Qty_Lat_Nguoc;
            this.Qty_Nguoc_Huong = s.Qty_Nguoc_Huong;
            this.Qty_Nham_LK = s.Qty_Nham_LK;
            this.Qty_Di_Vat = s.Qty_Di_Vat;
            this.Qty_Thua_LK = s.Qty_Thua_LK;
            this.Qty_Bong = s.Qty_Bong;
            this.Qty_Khac = s.Qty_Khac;
            this.Qty_Lech = s.Qty_Lech;
            this.Qty_Vo = s.Qty_Vo;
            this.Qty_Dung_dung = s.Qty_Dung_dung;
        }
        public void SetValue(QtyErrType s)
        {
            this.Qty_HanGia = s.Qty_HanGia;
            this.Qty_SaiVitri = s.Qty_SaiVitri;
            this.Qty_Kenh = s.Qty_Kenh;
            this.Qty_Baccau = s.Qty_Baccau;
            this.Qty_It_Thiec = s.Qty_It_Thiec;
            this.Qty_Thieu_LK = s.Qty_Thieu_LK;
            this.Qty_Lat_Nguoc = s.Qty_Lat_Nguoc;
            this.Qty_Nguoc_Huong = s.Qty_Nguoc_Huong;
            this.Qty_Nham_LK = s.Qty_Nham_LK;
            this.Qty_Di_Vat = s.Qty_Di_Vat;
            this.Qty_Thua_LK = s.Qty_Thua_LK;
            this.Qty_Bong = s.Qty_Bong;
            this.Qty_Khac = s.Qty_Khac;
            this.Qty_Lech = s.Qty_Lech;
            this.Qty_Vo = s.Qty_Vo;
            this.Qty_Dung_dung = s.Qty_Dung_dung;
        }
        public void SetValueToList(List<DataDD> list)
        {
            var qtyTypeProperties = typeof(QtyErrType).GetProperties();
            foreach (var itemCurrent in list)
            {
                foreach (var qtyTypeProperty in qtyTypeProperties)
                {
                    var errorQtyTypeValue = (long)qtyTypeProperty.GetValue(itemCurrent.QtyType);
                    var itemQtyTypeValue = (long)qtyTypeProperty.GetValue(this);
                    qtyTypeProperty.SetValue(this, itemQtyTypeValue + errorQtyTypeValue);
                }
            }
        }

        public void SetAll0()
        {
            this.Qty_HanGia = 0;
            this.Qty_SaiVitri = 0;
            this.Qty_Kenh = 0;
            this.Qty_Baccau = 0;
            this.Qty_It_Thiec = 0;
            this.Qty_Thieu_LK = 0;
            this.Qty_Lat_Nguoc = 0;
            this.Qty_Nguoc_Huong = 0;
            this.Qty_Nham_LK = 0;
            this.Qty_Di_Vat = 0;
            this.Qty_Thua_LK = 0;
            this.Qty_Bong = 0;
            this.Qty_Khac = 0;
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
        public QtyErrType QtyType { get; set; }

        public DataWrite()
        {
            this.QtyType = new QtyErrType();
        }
        public DataWrite(DataWrite s)
        {
            this.model = s.model;
            this.qty = s.qty;
            this.qtyErr = s.qtyErr;
            this.modelFirst = s.modelFirst;

            this.QtyType = new QtyErrType();

            this.QtyType.SetValue(s.QtyType);
        }
    }
}
