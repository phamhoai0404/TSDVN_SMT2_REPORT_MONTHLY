using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class WriteKM
    {
        public static void GetKM(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string typeKH)
        {
            List<DataDD> listChild = listDD.Where(p => typeKH.Contains(p.KH)).ToList();
            if (listChild.Count() == 0)
                return;

            int startRow = 3;
            int lastRow = ws.Dimension.End.Row; // Lấy dòng cuối cùng
            object[,] listModel = ws.Cells[$"A{startRow}:B{lastRow}"].Value as object[,];

            string tempModel = "";
            List<DataWrite> listWrite = new List<DataWrite>();
            DataWrite tempValue = new DataWrite();
            for (int i = 0; i < listModel.GetLength(0); i++)
            {
                if (listModel[i, 1] is null || string.IsNullOrWhiteSpace(listModel[i, 1].ToString()))
                    break;

                tempModel = listModel[i, 1]?.ToString().Trim().Substring(0, 9).ToUpper();
                var listTemp = listChild.Where(p => p.Model == tempModel).ToList();
                
                tempValue.QtyType.SetAll0();//Set mac dinh ve all 0
                if (listTemp.Count() > 0)
                {
                    tempValue.model = listModel[i, 1]?.ToString() + listModel[i, 0].ToString();//Thuc hien lay truc tiep luon vi  du lieu cua SMT la lay truc tiep
                    tempValue.qty = listTemp.Sum(p => p.Qty);
                    tempValue.qtyErr = listTemp.Sum(p => p.QtyError);
                    tempValue.modelFirst = tempModel;

                    tempValue.QtyType.SetValueToList(listTemp);//Thuc hien lay du lieu
                }
                else
                {
                    tempValue.model = listModel[i, 1]?.ToString() + listModel[i, 0]?.ToString();
                    tempValue.modelFirst = tempModel;
                    tempValue.qty = 0;
                    tempValue.qtyErr = 0;
                    
                }
                listWrite.Add(new DataWrite(tempValue));
            }

            object[,] data1 = new object[listWrite.Count(), 1];
            object[,] data2 = new object[listWrite.Count(), 1];
            object[,] data3 = new object[listWrite.Count(), 1];
            object[,] data333 = new object[listWrite.Count(), 13];
            for (int i = 0; i < listWrite.Count(); i++)
            {
                data1[i, 0] = listWrite[i].model;
                data2[i, 0] = listWrite[i].qty;
                data3[i, 0] = listWrite[i].qtyErr;

                data333[i, 0] = listWrite[i].QtyType.Qty_HanGia == 0 ? (long?)null : listWrite[i].QtyType.Qty_HanGia;
                data333[i, 1] = listWrite[i].QtyType.Qty_SaiVitri == 0 ? (long?)null : listWrite[i].QtyType.Qty_SaiVitri;
                data333[i, 2] = listWrite[i].QtyType.Qty_Kenh == 0 ? (long?)null : listWrite[i].QtyType.Qty_Kenh;
                data333[i, 3] = listWrite[i].QtyType.Qty_Baccau == 0 ? (long?)null : listWrite[i].QtyType.Qty_Baccau;
                data333[i, 4] = listWrite[i].QtyType.Qty_It_Thiec == 0 ? (long?)null : listWrite[i].QtyType.Qty_It_Thiec;
                data333[i, 5] = listWrite[i].QtyType.Qty_Thieu_LK == 0 ? (long?)null : listWrite[i].QtyType.Qty_Thieu_LK;
                data333[i, 6] = listWrite[i].QtyType.Qty_Lat_Nguoc == 0 ? (long?)null : listWrite[i].QtyType.Qty_Lat_Nguoc;
                data333[i, 7] = listWrite[i].QtyType.Qty_Nguoc_Huong == 0 ? (long?)null : listWrite[i].QtyType.Qty_Nguoc_Huong;
                data333[i, 8] = listWrite[i].QtyType.Qty_Nham_LK == 0 ? (long?)null : listWrite[i].QtyType.Qty_Nham_LK;
                data333[i, 9] = listWrite[i].QtyType.Qty_Di_Vat == 0 ? (long?)null : listWrite[i].QtyType.Qty_Di_Vat;
                data333[i, 10] = listWrite[i].QtyType.Qty_Thua_LK == 0 ? (long?)null : listWrite[i].QtyType.Qty_Thua_LK;
                data333[i, 11] = listWrite[i].QtyType.Qty_Bong == 0 ? (long?)null : listWrite[i].QtyType.Qty_Bong;
                data333[i, 12] = (listWrite[i].QtyType.Qty_Khac +
                                    listWrite[i].QtyType.Qty_Vo +
                                    listWrite[i].QtyType.Qty_Lech +
                                    listWrite[i].QtyType.Qty_Dung_dung
                                    ) == 0 ? (long?)null : (listWrite[i].QtyType.Qty_Khac +
                                        listWrite[i].QtyType.Qty_Vo +
                                        listWrite[i].QtyType.Qty_Lech +
                                        listWrite[i].QtyType.Qty_Dung_dung
                                    );
            }
            ws.Cells[$"C{startRow}:C{listWrite.Count() + startRow}"].Value = data1;
            ws.Cells[$"E{startRow}:E{listWrite.Count() + startRow}"].Value = data2;
            ws.Cells[$"G{startRow}:G{listWrite.Count() + startRow}"].Value = data3;
            ws.Cells[$"J{startRow}:J{listChild.Count() + startRow}"].Value = data333;

            var listErrFirst = GetError(listError, typeKH);
            var listErr = listErrFirst.Where(p => listWrite.Any(z => z.modelFirst == p.Model)).ToList();
            if (listErr.Count() > 0)
            {
                object[,] data11 = new object[listErr.Count(), 3];
                

                for (int i = 0; i < listErr.Count(); i++)
                {
                    data11[i, 0] = listErr[i].Model;
                    data11[i, 1] = listErr[i].Content;
                    data11[i, 2] = listErr[i].QtyError;

                }
                ws.Cells[$"Z{startRow}:Z{listErr.Count() + startRow}"].Value = data11;
            }
        }
        private static List<DataError2> GetError(List<DataError> listTemp, string khachhang)
        {
            return listTemp.Where(p => khachhang.Contains(p.KH) &&
            (p.QtyType.Qty_Khac != 0 ||
                p.QtyType.Qty_Vo != 0 ||
                p.QtyType.Qty_Dung_dung != 0 ||
                p.QtyType.Qty_Lech != 0
            )
            ).GroupBy(p => new { p.Content, p.Model })
                            .Select(group =>
                            {
                                return new DataError2
                                {
                                    Model = group.Key.Model,
                                    Content = group.Key.Content,
                                    QtyError = group.Sum(p => p.QtyError),
                                };
                            })
                            .OrderBy(p => p.Model)
                            .ToList();
        }
    }
}
