using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public static class WriteRISO
    {
        public static void GetWriteRISO(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string date, string typeKH)
        {
            long numberQty1 = listDD.Where(p => typeKH.Contains(p.KH)).Sum(p => p.Qty);
            long numberErr1 = listError.Where(p => typeKH.Contains(p.KH)).Sum(p => p.QtyError);
            int rowFirst = 3;
            ws.Cells[$"A{rowFirst}"].Value = date + "月";
            ws.Cells[$"B{rowFirst}"].Value = numberQty1;
            ws.Cells[$"C{rowFirst}"].Value = numberErr1;

            if (numberErr1 > 0)
            {
                //DAY LA VIET CHUNG
                QtyErrType value = ActionWriteDD.CalculateQtyErrType(listError, typeKH);

                object[,] data333 = new object[1, 12];
                data333[0, 0] = value.Qty_Nham_LK == 0 ? (long?)null : value.Qty_Nham_LK;
                data333[0, 1] = value.Qty_Di_Vat == 0 ? (long?)null : value.Qty_Di_Vat;
                data333[0, 2] = value.Qty_Thua_LK == 0 ? (long?)null : value.Qty_Thua_LK;
                data333[0, 3] = value.Qty_Bong == 0 ? (long?)null : value.Qty_Bong;
                data333[0, 4] = value.Qty_Kenh == 0 ? (long?)null : value.Qty_Kenh;
                data333[0, 5] = value.Qty_Thieu_LK == 0 ? (long?)null : value.Qty_Thieu_LK;
                data333[0, 6] = value.Qty_Lat_Nguoc == 0 ? (long?)null : value.Qty_Lat_Nguoc;
                data333[0, 7] = value.Qty_Nguoc_Huong == 0 ? (long?)null : value.Qty_Nguoc_Huong;
                data333[0, 8] = value.Qty_It_Thiec == 0 ? (long?)null : value.Qty_It_Thiec;
                data333[0, 9] = value.Qty_Baccau == 0 ? (long?)null : value.Qty_Baccau;
                data333[0, 10] = value.Qty_SaiVitri == 0 ? (long?)null : value.Qty_SaiVitri;

                long qty = 
                data333[0, 11] = (value.Qty_Khac + value.Qty_Bong + value.Qty_Lech + value.Qty_Dung_dung + value.Qty_Vo) == 0 ? (long?)null : (value.Qty_Khac + value.Qty_Bong + value.Qty_Lech + value.Qty_Dung_dung + value.Qty_Vo);

                ws.Cells[$"F{rowFirst}:F{rowFirst + 1}"].Value = data333;

                //DAY LA VIET CHI TIET
                int currentRow = 7;
                var listErr = GetErrorRISO(listError, typeKH);

                object[,] data1 = new object[listErr.Count(), 5];
                for (int i = 0; i < listErr.Count(); i++)
                {
                    data1[i, 0] = i + 1;
                    data1[i, 1] = listErr[i].Model;
                    data1[i, 2] = listErr[i].TypeItem;
                    data1[i, 3] = listErr[i].Content;
                    data1[i, 4] = listErr[i].QtyError;
                }
                ws.Cells[$"A{currentRow}:A{listErr.Count() + currentRow}"].Value = data1;

                var listErr222 = GetErrorRISO2(listError, typeKH);
                object[,] data22 = new object[listErr.Count(), 3];
                for (int i = 0; i < listErr222.Count(); i++)
                {
                    data22[i, 0] = i + 1;
                    data22[i, 1] = listErr222[i].Content;
                    data22[i, 2] = listErr222[i].QtyError;
                }
                ws.Cells[$"G{currentRow}:G{listErr222.Count() + currentRow}"].Value = data22;

            }
        }

        private static List<DataError3> GetErrorRISO(List<DataError> listTemp, string khachhang)
        {
            return listTemp.Where(p => khachhang.Contains(p.KH) &&
            (   p.QtyType.Qty_Khac != 0 ||
                p.QtyType.Qty_Bong != 0 || 
                p.QtyType.Qty_Vo != 0 || 
                p.QtyType.Qty_Dung_dung != 0 || 
                p.QtyType.Qty_Lech != 0
            )
            ).GroupBy(p => new { p.Content, p.Model, p.TypeItem })
                            .Select(group =>
                            {
                                return new DataError3
                                {
                                    Model = group.Key.Model,
                                    Content = group.Key.Content,
                                    TypeItem = group.Key.TypeItem,
                                    QtyError = group.Sum(p => p.QtyError),
                                };
                            })
                            .OrderBy(p => p.Model)
                            .ToList();
        }
        private static List<DataError4> GetErrorRISO2(List<DataError> listTemp, string khachhang)
        {
            //Chi lay du lieu la cai khac thoi
            return listTemp.Where(p => khachhang.Contains(p.KH) &&
            (   p.QtyType.Qty_Khac != 0 ||
                p.QtyType.Qty_Bong != 0 || 
                p.QtyType.Qty_Vo != 0 || 
                p.QtyType.Qty_Dung_dung != 0 ||
                p.QtyType.Qty_Lech != 0
            )
            ).GroupBy(p => p.Content)
                    .Select(group => new DataError4
                    {
                        Content = group.Key,
                        QtyError = group.Sum(p => p.QtyError)
                    })
                    .ToList();
        }
    }
}
