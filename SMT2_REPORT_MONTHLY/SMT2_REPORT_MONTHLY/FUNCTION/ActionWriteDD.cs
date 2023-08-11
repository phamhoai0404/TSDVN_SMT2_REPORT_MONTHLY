using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class ActionWriteDD
    {
        public static void WriteData(List<DataDD> listDD, List<DataError> listErr, TypeWrite type, ref string newFile, string date)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileTemp = new FileInfo(MdlCommon.PATH_TEMPLATE);
                using (var package = new ExcelPackage(fileTemp))
                {
                    ExcelWorksheet worksheet;
                    string typeName;

                    if (type.riso == true)
                    {
                        typeName = ConfigurationManager.AppSettings["RISO"];
                        worksheet = package.Workbook.Worksheets["RISO"];
                        GetWrite(worksheet, listDD, listErr, date, typeName);
                    }
                    if (type.oki == true)
                    {
                        typeName = ConfigurationManager.AppSettings["OKIDENKI"];
                        worksheet = package.Workbook.Worksheets["OKIDENKI"];
                        GetWrite(worksheet, listDD, listErr, date, typeName);
                    }
                    if (type.kyo == true)
                    {
                        typeName = ConfigurationManager.AppSettings["KYOCERA"];
                        worksheet = package.Workbook.Worksheets["KYOCERA"];
                        GetKYOCERA(worksheet, listDD, listErr, typeName);
                    }
                    if (type.km == true)
                    {
                        typeName = ConfigurationManager.AppSettings["KM"];
                        worksheet = package.Workbook.Worksheets["KM"];
                        GetKM(worksheet, listDD, listErr, typeName);
                    }

                    worksheet = package.Workbook.Worksheets[ConfigurationManager.AppSettings["ALL"]];
                    GetAll(worksheet, listDD, listErr);

                    worksheet = package.Workbook.Worksheets[ConfigurationManager.AppSettings["DATA_ERROR"]];
                    GetDataError(worksheet, listErr);

                    newFile = Path.Combine(Directory.GetCurrentDirectory(), "RESULT", DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileTemp.Extension);
                    FileInfo fileTempNew = new FileInfo(newFile);
                    package.SaveAs(fileTempNew);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(RESULT.ERROR_015_CATCH, "WriteData", ex.Message));
            }
        }

        private static void GetWrite(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string date, string typeKH)
        {
            long numberQty1 = listDD.Where(p => typeKH.Contains(p.KH)).Sum(p => p.Qty);
            long numberErr1 = listError.Where(p => typeKH.Contains(p.KH)).Sum(p => p.QtyError);
            ws.Cells["A3"].Value = date + "月";
            ws.Cells["C3"].Value = numberQty1;
            ws.Cells["D3"].Value = numberErr1;

            if (numberErr1 > 0)
            {
                var listErr = GetError(listError, typeKH);
                int currentRow = 7;
                object[,] data1 = new object[listErr.Count(), 4];
                for (int i = 0; i < listErr.Count(); i++)
                {
                    data1[i, 0] = i + 1;
                    data1[i, 1] = listErr[i].Model;
                    data1[i, 2] = listErr[i].Content;
                    data1[i, 3] = listErr[i].QtyError;
                }
                ws.Cells[$"A{currentRow}:D{listErr.Count() + currentRow}"].Value = data1;

                var listErr222 = GetError4(listError, typeKH);
                object[,] data22 = new object[listErr.Count(), 3];
                for (int i = 0; i < listErr222.Count(); i++)
                {
                    data22[i, 0] = i + 1;
                    data22[i, 1] = listErr222[i].Content;
                    data22[i, 2] = listErr222[i].QtyError;
                }
                ws.Cells[$"G{currentRow}:I{listErr222.Count() + currentRow}"].Value = data22;
            }
        }
        private static List<DataError2> GetError(List<DataError> listTemp, string khachhang)
        {
            return listTemp.Where(p => khachhang.Contains(p.KH)).GroupBy(p => new { p.Content, p.Model })
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
            return listTemp.Where(p => khachhang.Contains(p.KH)).GroupBy(p => new { p.Content, p.Model, p.TypeItem })
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
        private static List<DataError4> GetError4(List<DataError> listTemp)
        {
            return listTemp.GroupBy(p => p.Content)
                    .Select(group => new DataError4
                    {
                        Content = group.Key,
                        QtyError = group.Sum(p => p.QtyError)
                    })
                    .ToList();
        }
        private static List<DataError4> GetError4(List<DataError> listTemp, string khachhang)
        {
            return listTemp.Where(p => khachhang.Contains(p.KH)).GroupBy(p => p.Content)
                    .Select(group => new DataError4
                    {
                        Content = group.Key,
                        QtyError = group.Sum(p => p.QtyError)
                    })
                    .ToList();
        }

        private static void GetKYOCERA(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string typeKH)
        {
            List<DataDD> listChild = listDD.Where(p => typeKH.Contains(p.KH)).OrderBy(p => p.Model).ToList();
            if (listChild.Count() == 0)
                return;

            int startWrite = 3;
            object[,] data = new object[listChild.Count(), 3];
            for (int i = 0; i < listChild.Count(); i++)
            {
                data[i, 0] = listChild[i].Model;
                data[i, 1] = listChild[i].Qty;
                data[i, 2] = listChild[i].QtyError;

            }
            ws.Cells[$"A{startWrite}:C{listChild.Count() + startWrite}"].Value = data;

            var listErrorWrite = GetError3(listError, typeKH);
            if (listErrorWrite.Count() > 0)
            {
                object[,] data1 = new object[listErrorWrite.Count(), 1];
                object[,] data2 = new object[listErrorWrite.Count(), 2];
                object[,] data3 = new object[listErrorWrite.Count(), 1];
                for (int i = 0; i < listErrorWrite.Count(); i++)
                {
                    data1[i, 0] = listErrorWrite[i].Model;
                    data2[i, 0] = listErrorWrite[i].TypeItem;
                    data2[i, 1] = listErrorWrite[i].Content;
                    data3[i, 0] = listErrorWrite[i].QtyError;
                }
                ws.Cells[$"V{startWrite}:V{listErrorWrite.Count() + startWrite}"].Value = data1;
                ws.Cells[$"Y{startWrite}:AA{listErrorWrite.Count() + startWrite}"].Value = data2;
                ws.Cells[$"AB{startWrite}:AB{listErrorWrite.Count() + startWrite}"].Value = data3;
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
                ws.Cells[$"AD{startWrite}:AF{listErrorWrite2.Count() + startWrite}"].Value = data1;
            }
        }
        private static void GetKM(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string typeKH)
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
                if (listTemp.Count() > 0)
                {
                    tempValue.model = listModel[i, 1]?.ToString() + listModel[i, 0].ToString();//Thuc hien lay truc tiep luon vi  du lieu cua SMT la lay truc tiep
                    tempValue.qty = listTemp.Sum(p => p.Qty);
                    tempValue.qtyErr = listTemp.Sum(p => p.QtyError);
                    tempValue.modelFirst = tempModel;
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
            for (int i = 0; i < listWrite.Count(); i++)
            {
                data1[i, 0] = listWrite[i].model;
                data2[i, 0] = listWrite[i].qty;
                data3[i, 0] = listWrite[i].qtyErr;
            }
            ws.Cells[$"C{startRow}:C{listWrite.Count() + startRow}"].Value = data1;
            ws.Cells[$"E{startRow}:E{listWrite.Count() + startRow}"].Value = data2;
            ws.Cells[$"G{startRow}:G{listWrite.Count() + startRow}"].Value = data3;

            var listErrFirst = GetError(listError, typeKH);
            var listErr = listErrFirst.Where(p => listWrite.Any(z => z.modelFirst == p.Model)).ToList();
            if (listErr.Count() > 0)
            {
                object[,] data11 = new object[listErr.Count(), 1];
                object[,] data12 = new object[listErr.Count(), 1];
                object[,] data13 = new object[listErr.Count(), 1];

                for (int i = 0; i < listErr.Count(); i++)
                {
                    data11[i, 0] = listErr[i].Model;
                    data12[i, 0] = listErr[i].Content;
                    data13[i, 0] = listErr[i].QtyError;

                }
                ws.Cells[$"Z{startRow}:Z{listErr.Count() + startRow}"].Value = data11;
                ws.Cells[$"AB{startRow}:AB{listErr.Count() + startRow}"].Value = data12;
                ws.Cells[$"AE{startRow}:AE{listErr.Count() + startRow}"].Value = data13;
            }
        }
        private static void GetAll(ExcelWorksheet ws, List<DataDD> listG2, List<DataError> listErr)
        {
            listG2 = listG2.OrderBy(p => p.Model).ToList();
            int startRow = 3;
            object[,] data = new object[listG2.Count(), 6];
            for (int i = 0; i < listG2.Count(); i++)
            {
                data[i, 0] = i + 1;
                data[i, 1] = listG2[i].KH;
                data[i, 2] = listG2[i].Model;
                data[i, 3] = listG2[i].Mat;
                data[i, 4] = listG2[i].Qty;
                data[i, 5] = listG2[i].QtyError;
            }
            ws.Cells[$"A{startRow}:F{listG2.Count() + startRow}"].Value = data;

            object[,] data2 = new object[listErr.Count(), 7];
            for (int i = 0; i < listErr.Count(); i++)
            {
                data2[i, 0] = i + 1;
                data2[i, 1] = listErr[i].KH;
                data2[i, 2] = listErr[i].Model;
                data2[i, 3] = listErr[i].TypeItem;
                data2[i, 4] = listErr[i].Mat;
                data2[i, 5] = listErr[i].QtyError;
                data2[i, 6] = listErr[i].Content;
            }
            ws.Cells[$"I{startRow}:N{listG2.Count() + startRow}"].Value = data2;
        }

        private static void GetDataError(ExcelWorksheet ws, List<DataError> listErr)
        {
            List<DataError4> listErrGroup = GetError4(listErr);
            if (listErrGroup.Count() == 0)
                return;
            int startRow = 3;
            object[,] data = new object[listErrGroup.Count(), 3];
            for (int i = 0; i < listErrGroup.Count(); i++)
            {
                data[i, 0] = i + 1;
                data[i, 1] = listErrGroup[i].Content;
                data[i, 2] = listErrGroup[i].QtyError;
              
            }
            ws.Cells[$"A{startRow}:C{listErrGroup.Count() + startRow}"].Value = data;
        }
    }
}
