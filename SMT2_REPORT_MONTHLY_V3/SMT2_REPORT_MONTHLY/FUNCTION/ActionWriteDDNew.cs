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
    public class ActionWriteDDNew
    {
        public static void WriteData(List<DataDD> listDataDDFrist, List<DataError> listErr, TypeWrite type, ref string newFile, string date)
        {
            try
            {
                List<DataDD> listDD = new List<DataDD>();//Du lieu diem dan sau gop
                ActionDD.MerageValue(ref listDataDDFrist, ref listDD);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileTemp = new FileInfo(MdlCommon.PATH_TEMPLATE);
                using (var package = new ExcelPackage(fileTemp))
                {
                    ExcelWorksheet worksheet;
                    if (type.riso == true)
                    {
                        worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_RISO];
                        WriteDataChild(worksheet, listDD, listErr, date, MdlCommon.TYPE_NAME_RISO, MdlCommon.LIST_ERROR_RISO);
                    }
                    if (type.oki == true)
                    {
                        worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_OKIDENKI];
                        WriteDataChild(worksheet, listDD, listErr, date, MdlCommon.TYPE_NAME_OKIDENKI, MdlCommon.LIST_ERROR_OKIDENKI);
                    }
                    if (type.kyo == true)
                    {
                        worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_KYOCERA];
                        WriteKYO(worksheet, listDD, listErr, MdlCommon.TYPE_NAME_KYOCERA, MdlCommon.LIST_ERROR_KYOCERA);
                    }
                    if (type.km == true)
                    {
                        worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_KM];
                        WriteKM(worksheet, listDD, listErr, MdlCommon.TYPE_NAME_KM, MdlCommon.LIST_ERROR_KM);
                    }

                    worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_ALL];
                    GetAll(worksheet, listDataDDFrist, listErr, listDD);

                    worksheet = package.Workbook.Worksheets[MdlCommon.TYPE_NAME_DATA_ERROR];
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

        private static void WriteKM(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listErr, string typeKH, List<TitleError> listTitle)
        {
            List<DataDD> listChild = listDD.Where(p => typeKH.Contains(p.KH)).ToList();
            if (listChild.Count() == 0)
                return;

            int rowWrite = MdlCommon.ROW_TITTLE_GET_TEMPLATE + 1;
            int lastRow = ws.Dimension.End.Row; // Lấy dòng cuối cùng
            object[,] listModel = ws.Cells[$"A{rowWrite}:B{lastRow}"].Value as object[,];

            string colStartError = listTitle[0].Address.ColName;
            int numberCol = listTitle.Count();
            object[,] data = new object[listModel.GetLength(0), 5];
            object[,] data333 = new object[listModel.GetLength(0), numberCol];
            long[] tempSum = new long[numberCol];//Thuc hien luu tru mang rieng
            string tempModel = "";

            int numberAllItem = listModel.GetLength(0);
            List<DataError> listErrChildAll = new List<DataError>();//Luu toan bo child con
            for (int i = 0; i < numberAllItem; i++)
            {
                if (listModel[i, 1] is null || string.IsNullOrWhiteSpace(listModel[i, 1].ToString()))
                    break;

                tempModel = listModel[i, 1]?.ToString().Trim().Substring(0, 9).ToUpper();
                var listTemp = listChild.Where(p => p.Model == tempModel).ToList();

                if (listTemp.Count() > 0)
                {
                    data[i, 0] = listModel[i, 1]?.ToString() + listModel[i, 0].ToString();//Thuc hien lay truc tiep luon vi  du lieu cua SMT la lay truc tiep
                    data[i, 1] = listTemp.Sum(p => p.Qty);

                    var listErrorChild = listErr.Where(p => p.Model == tempModel && typeKH.Contains(p.KH)).ToList();
                    listErrChildAll.AddRange(listErrorChild);
                    for (int u = 0; u < numberCol; u++)
                    {
                        tempSum[u] = 0; // Reset tổng trước khi tính lại
                        foreach (var errorChild in listErrorChild)
                        {
                            tempSum[u] += errorChild.listErr[u];
                        }
                        data333[i, u] = tempSum[u] == 0 ? (long?)null : tempSum[u];
                    }
                    
                    data[i, 2] = listTemp.Sum(p => p.PointQty);
                    data[i, 3] = listTemp.Sum(p => p.PointQty * p.Qty);
                    data[i, 4] = listErrorChild.Sum(p => p.listErr.Sum(z => z));

                }
                else
                {
                    data[i, 0] = listModel[i, 1]?.ToString() + listModel[i, 0].ToString();//Thuc hien lay truc tiep luon vi  du lieu cua SMT la lay truc tiep
                    data[i, 1] = 0;
                    data[i, 2] = 0;
                    data[i, 3] = 0;
                    data[i, 4] = 0;
                }
            }

            ws.Cells[$"C{rowWrite}:C{numberAllItem + rowWrite}"].Value = data;
            ws.Cells[$"{colStartError}{rowWrite}:{colStartError}{numberAllItem + rowWrite}"].Value = data333;

            if (listErrChildAll.Count == 0)//
                return;

            //Viet chi tiet loi neu co 
            WriteDetail(ws, listErrChildAll, typeKH, rowWrite, numberCol, listTitle[numberCol - 1].Address.Index);

        }

        private static void WriteKYO(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listErr, string typeKH, List<TitleError> listTitle)
        {
            List<DataDD> listChild = listDD.Where(p => typeKH.Contains(p.KH)).OrderBy(p => p.Model).ToList();
            if (listChild.Count() == 0)
                return;

            string colStartError = listTitle[0].Address.ColName;
            int numberCol = listTitle.Count();

            object[,] data = new object[listChild.Count(), 5];
            object[,] data333 = new object[listChild.Count(), numberCol];
            long[] tempSum = new long[numberCol];
            for (int i = 0; i < listChild.Count(); i++)
            {
                data[i, 0] = listChild[i].Model;
                data[i, 1] = listChild[i].Qty;

                var listErrorChild = listErr.Where(p => p.Model == listChild[i].Model && typeKH.Contains(listChild[i].KH)).ToList();
                for (int u = 0; u < numberCol; u++)
                {
                    tempSum[u] = 0; // Reset tổng trước khi tính lại
                    foreach (var errorChild in listErrorChild)
                    {
                        tempSum[u] += errorChild.listErr[u];
                    }
                    data333[i, u] = tempSum[u] == 0 ? (long?)null : tempSum[u];
                }
                
                data[i, 2] = listChild[i].PointQty;
                data[i, 3] = listChild[i].PointQty * listChild[i].Qty;
                data[i, 4] = listErrorChild.Sum(p => p.listErr.Sum(z => z));
            }
            int rowWrite = MdlCommon.ROW_TITTLE_GET_TEMPLATE + 1;
            ws.Cells[$"A{rowWrite}:A{listChild.Count() + rowWrite}"].Value = data;
            ws.Cells[$"{colStartError}{rowWrite}:{colStartError}{listChild.Count() + rowWrite}"].Value = data333;

            //Viet chi tiet loi neu co 
            WriteDetail(ws, listErr, typeKH, rowWrite, numberCol, listTitle[numberCol - 1].Address.Index);
        }

        private static void WriteDataChild(ExcelWorksheet ws, List<DataDD> listDD, List<DataError> listError, string date, string typeKH, List<TitleError> listTitle)
        {
            long numberQty1 = listDD.Where(p => typeKH.Contains(p.KH)).Sum(p => p.Qty * p.PointQty);
            long numberErr1 = listError.Where(p => typeKH.Contains(p.KH)).Sum(p => p.QtyError);
            int rowWrite = MdlCommon.ROW_TITTLE_GET_TEMPLATE + 1;
            ws.Cells[$"A{rowWrite}"].Value = date + "月";
            ws.Cells[$"B{rowWrite}"].Value = numberQty1;
            ws.Cells[$"C{rowWrite}"].Value = numberErr1;

            if (numberErr1 == 0)//Neu khong co loi thi dung lai luon
                return;

            string colStartError = listTitle[0].Address.ColName;
            int numberCol = listTitle.Count();
            object[,] data333 = new object[1, numberCol];
            List<long> valueSum = GetSumQTy(listError, typeKH, numberCol);//Thuc hien lay gia tri tong
            for (int i = 0; i < numberCol; i++)
                data333[0, i] = valueSum[i] == 0 ? (long?)null : valueSum[i];

            ws.Cells[$"{colStartError}{rowWrite}:{colStartError}{rowWrite + 1}"].Value = data333;

            //VIET CHI TIET KHAC NEU CO
            WriteDetail(ws, listError, typeKH, rowWrite, numberCol, listTitle[numberCol - 1].Address.Index);

        }
        private static void WriteDetail(ExcelWorksheet ws, List<DataError> listError, string typeKH, int rowWrite, int numberCol, int colEndTitleErr)
        {
            var listErr = ActionGetWriteError.GetError3(listError, typeKH, numberCol);
            if (listErr.Count() != 0)
            {
                int colEndTitleError = colEndTitleErr + 1;//Vi tri can cong 1 vi no dang tru di 1 
                string colWriteError = MyFunction1.ConvertNumberToName(colEndTitleError + 4);//Cach xa 4 cot
                object[,] data1 = new object[listErr.Count(), 6];
                for (int i = 0; i < listErr.Count(); i++)
                {
                    data1[i, 0] = i + 1;
                    data1[i, 1] = listErr[i].Model;
                    data1[i, 2] = listErr[i].TypeItem;
                    data1[i, 3] = listErr[i].Content;
                    data1[i, 4] = listErr[i].QtyError;
                    data1[i, 5] = listErr[i].ContentKH;
                }
                ws.Cells[$"{colWriteError}{rowWrite}:{colWriteError}{listErr.Count() + rowWrite}"].Value = data1;

                var listErr222 = ActionGetWriteError.GetError2(listError, typeKH, numberCol);
                object[,] data22 = new object[listErr.Count(), 4];
                for (int i = 0; i < listErr222.Count(); i++)
                {
                    data22[i, 0] = i + 1;
                    data22[i, 1] = listErr222[i].Content;
                    data22[i, 2] = listErr222[i].QtyError;
                    data22[i, 3] = listErr222[i].ContentKH;
                }
                colWriteError = MyFunction1.ConvertNumberToName(colEndTitleError + 11);//Cach xa 11 cột
                ws.Cells[$"{colWriteError}{rowWrite}:{colWriteError}{listErr222.Count() + rowWrite}"].Value = data22;
            }
        }
        private static List<long> GetSumQTy(List<DataError> listError, string typeKH, int numberCol)
        {
            List<long> tempValue = new List<long>();

            var listAll = listError.Where(p => typeKH.Contains(p.KH)).ToList();
            long tempSum = 0;
            for (int i = 0; i < numberCol; i++)
            {
                tempSum = listAll.Sum(p => p.listErr[i]);
                tempValue.Add(tempSum);
            }
            return tempValue;
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
            ws.Cells[$"A{startRow}:A{listErrGroup.Count() + startRow}"].Value = data;
        }
        private static void GetAll(ExcelWorksheet ws, List<DataDD> listG2, List<DataError> listErr, List<DataDD> listAfter)
        {
            listG2 = listG2.OrderBy(p => p.Model).ToList();
            int startRow = 3;
            WriteChildDD(ws, listG2, startRow, "A");
            WriteChildDD(ws, listAfter, startRow, "S");
            object[,] data2 = new object[listErr.Count(), 8];
            for (int i = 0; i < listErr.Count(); i++)
            {
                data2[i, 0] = i + 1;
                data2[i, 1] = listErr[i].KH;
                data2[i, 2] = listErr[i].Model;
                data2[i, 3] = listErr[i].TypeItem;
                data2[i, 4] = listErr[i].Mat;
                data2[i, 5] = listErr[i].QtyError;
                data2[i, 6] = listErr[i].Content;
                data2[i, 7] = listErr[i].ContentKH;
            }
            ws.Cells[$"I{startRow}:I{listG2.Count() + startRow}"].Value = data2;
        }
        private static void WriteChildDD(ExcelWorksheet ws, List<DataDD> listG2, int startRow, string nameColFirst)
        {
            object[,] data = new object[listG2.Count(), 6];
            for (int i = 0; i < listG2.Count(); i++)
            {
                data[i, 0] = i + 1;
                data[i, 1] = listG2[i].KH;
                data[i, 2] = listG2[i].Model;
                data[i, 3] = listG2[i].Mat;
                data[i, 4] = listG2[i].Qty;
                data[i, 5] = listG2[i].PointQty;
            }
            ws.Cells[$"{nameColFirst}{startRow}:{nameColFirst}{listG2.Count() + startRow}"].Value = data;
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
    }


}
