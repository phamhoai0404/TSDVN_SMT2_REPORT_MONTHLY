using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class WriteKYO
    {
        public static void GetKYOCERA(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string typeKH)
        {
            List<DataDD> listChild = listDD.Where(p => typeKH.Contains(p.KH)).OrderBy(p => p.Model).ToList();
            if (listChild.Count() == 0)
                return;

            int startWrite = 3;
            object[,] data = new object[listChild.Count(), 3];
            object[,] data333 = new object[listChild.Count(), 13];
            for (int i = 0; i < listChild.Count(); i++)
            {
                data[i, 0] = listChild[i].Model;
                data[i, 1] = listChild[i].Qty;
                data[i, 2] = listChild[i].QtyError;

                data333[i, 0] = listChild[i].QtyType.Qty_HanGia == 0 ? (long?)null : listChild[i].QtyType.Qty_HanGia;
                data333[i, 1] = listChild[i].QtyType.Qty_SaiVitri == 0 ? (long?)null : listChild[i].QtyType.Qty_SaiVitri;
                data333[i, 2] = listChild[i].QtyType.Qty_Kenh == 0 ? (long?)null : listChild[i].QtyType.Qty_Kenh;
                data333[i, 3] = listChild[i].QtyType.Qty_Baccau == 0 ? (long?)null : listChild[i].QtyType.Qty_Baccau;
                data333[i, 4] = listChild[i].QtyType.Qty_It_Thiec == 0 ? (long?)null : listChild[i].QtyType.Qty_It_Thiec;
                data333[i, 5] = listChild[i].QtyType.Qty_Thieu_LK == 0 ? (long?)null : listChild[i].QtyType.Qty_Thieu_LK;
                data333[i, 6] = listChild[i].QtyType.Qty_Lat_Nguoc == 0 ? (long?)null : listChild[i].QtyType.Qty_Lat_Nguoc;
                data333[i, 7] = listChild[i].QtyType.Qty_Nguoc_Huong == 0 ? (long?)null : listChild[i].QtyType.Qty_Nguoc_Huong;
                data333[i, 8] = listChild[i].QtyType.Qty_Nham_LK == 0 ? (long?)null : listChild[i].QtyType.Qty_Nham_LK;
                data333[i, 9] = listChild[i].QtyType.Qty_Di_Vat == 0 ? (long?)null : listChild[i].QtyType.Qty_Di_Vat;
                data333[i, 10] = listChild[i].QtyType.Qty_Thua_LK == 0 ? (long?)null : listChild[i].QtyType.Qty_Thua_LK;
                data333[i, 11] = listChild[i].QtyType.Qty_Bong == 0 ? (long?)null : listChild[i].QtyType.Qty_Bong;
                data333[i, 12] = (listChild[i].QtyType.Qty_Khac +
                                    listChild[i].QtyType.Qty_Vo +
                                    listChild[i].QtyType.Qty_Lech +
                                    listChild[i].QtyType.Qty_Dung_dung
                                    ) == 0 ? (long?)null : (listChild[i].QtyType.Qty_Khac +
                                        listChild[i].QtyType.Qty_Vo +
                                        listChild[i].QtyType.Qty_Lech +
                                        listChild[i].QtyType.Qty_Dung_dung
                                    );

            }
            ws.Cells[$"A{startWrite}:C{listChild.Count() + startWrite}"].Value = data;
            ws.Cells[$"F{startWrite}:F{listChild.Count() + startWrite}"].Value = data333;

            var listErrorWrite = GetError3(listError, typeKH);
            if (listErrorWrite.Count() > 0)
            {
                object[,] data1 = new object[listErrorWrite.Count(), 4];
                for (int i = 0; i < listErrorWrite.Count(); i++)
                {
                    data1[i, 0] = listErrorWrite[i].Model;
                    data1[i, 1] = listErrorWrite[i].TypeItem;
                    data1[i, 2] = listErrorWrite[i].Content;
                    data1[i, 3] = listErrorWrite[i].QtyError;
                }
                ws.Cells[$"V{startWrite}:V{listErrorWrite.Count() + startWrite}"].Value = data1;
            }

            var listErrorWrite2 = GetError(listError, typeKH);
            if (listErrorWrite2.Count() > 0)
            {
                object[,] data1 = new object[listErrorWrite2.Count(), 3];
                for (int i = 0; i < listErrorWrite2.Count(); i++)
                {
                    data1[i, 0] = listErrorWrite2[i].Model;
                    data1[i, 1] = listErrorWrite2[i].Content;
                    data1[i, 2] = listErrorWrite2[i].QtyError;
                }
                ws.Cells[$"AA{startWrite}:AA{listErrorWrite2.Count() + startWrite}"].Value = data1;
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
        private static List<DataError3> GetError3(List<DataError> listTemp, string khachhang)
        {
            return listTemp.Where(p => khachhang.Contains(p.KH) &&
            (p.QtyType.Qty_Khac != 0 ||
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
    }
}
