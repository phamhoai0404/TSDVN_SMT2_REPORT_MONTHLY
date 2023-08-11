using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class ActionDD
    {
        public static void GetValueDD(ref List<DataDD> listDDAfter, DataConfigDD configDD, DateTime monthGet)
        {
            try
            {
                List<DataDD> listDD = new List<DataDD>();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(configDD.pathFile), false))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[configDD.sheetName];
                    if (worksheet == null)
                        throw new Exception($"File: {configDD.pathFile} ==== SheetName:{configDD.sheetName} => Không tồn tại!");

                    int lastRow = worksheet.Dimension.End.Row; // Lấy dòng cuối cùng
                    object[,] listAll = worksheet.Cells[$"A1:{configDD.ColLast}{lastRow}"].Value as object[,];
                    int rowAll = listAll.GetLength(0);

                    DataDD tempValue = new DataDD();
                    long numberQty = 0;
                    for (int i = 0; i < rowAll; i++)
                    {
                        if (!(listAll[i, configDD.Model.Index] is string) ||
                            string.IsNullOrWhiteSpace(listAll[i, configDD.Model.Index].ToString()) ||
                            !listAll[i, configDD.Model.Index].ToString().Contains("-"))
                            continue;

                        tempValue.Model = listAll[i, configDD.Model.Index].ToString().Trim().ToUpper();//Lay  gia tri cua model
                        if (tempValue.Model.Length < 9)
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "Model", configDD.Model.ColName + (i + 1), tempValue.Model));
                        tempValue.Model = tempValue.Model.Substring(0, 9);


                        //Kiem tra khach hang xem co du lieu hay khong neu khong co du lieu thi bao loi luon
                        if (!(listAll[i, configDD.KH.Index] is string) || string.IsNullOrWhiteSpace(listAll[i, configDD.KH.Index].ToString()))
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "Khách hàng", configDD.KH.ColName + (i + 1), listAll[i, configDD.KH.Index]));

                        //Thuc hien lay so luong
                        if (!long.TryParse(listAll[i, configDD.Qty.Index]?.ToString(), out numberQty))
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "QTY", configDD.Qty.ColName + (i + 1), listAll[i, configDD.Qty.Index]));

                        if (numberQty <= 0)//Neu so luong <= 0 thi duyet cai khac
                            continue;

                        if (!(listAll[i, configDD.Mat.Index] is string) || string.IsNullOrWhiteSpace(listAll[i, configDD.Mat.Index].ToString()))
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "Mặt", configDD.Mat.ColName + (i + 1), listAll[i, configDD.Mat.Index]));

                        tempValue.KH = listAll[i, configDD.KH.Index].ToString().Trim().ToUpper();
                        tempValue.Qty = numberQty;
                        tempValue.Mat = listAll[i, configDD.Mat.Index].ToString().Trim().ToUpper();

                        listDD.Add(new DataDD(tempValue));
                    }
                    if (listDD.Count == 0)
                    {
                        throw new Exception($"File: {configDD.pathFile}  = trong sheetName: {configDD.sheetName} => Không có dữ liệu tháng: {monthGet.ToString("MM.yyyy")}");
                    }
                }

                PairingDD(listDD, ref listDDAfter);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(RESULT.ERROR_015_CATCH, $"GetValueDD", ex.Message));
            }
        }

        private static void PairingDD(List<DataDD> listTemp, ref List<DataDD> listAfter)
        {
            var processedModels = new HashSet<string>();

            // Thực hiện kiểm tra dữ liệu
            foreach (var item in listTemp)
            {
                if (processedModels.Contains(item.Model))
                    continue;

                var listChild = listTemp.Where(p => p.Model == item.Model).ToList();
                if (listChild.Count() > 2)
                    throw new Exception($"Cần xem lại Model có: {listChild.Count} dòng dữ liệu!");

                if (listChild.Count() == 2 && listChild[0].Qty != listChild[1].Qty)
                {
                    throw new Exception($"Số lượng của model {listChild[0].Model} đang không đồng nhất!");
                }

                processedModels.Add(item.Model);
            }

            // Thuc hien gop du lieu
            listAfter = listTemp.GroupBy(p => new { p.KH, p.Model })
                            .Select(group =>
                            {
                                if (group.Count() == 1)
                                {
                                    return new DataDD
                                    {
                                        KH = group.Key.KH,
                                        Mat = group.First().Mat,
                                        Model = group.Key.Model,
                                        Qty = group.First().Qty
                                    };
                                }
                                else
                                {
                                    return new DataDD
                                    {
                                        KH = group.Key.KH,
                                        Mat = "2",
                                        Model = group.Key.Model,
                                        Qty = group.First().Qty
                                    };
                                }
                            })
                            .ToList();
        }
    }
}
