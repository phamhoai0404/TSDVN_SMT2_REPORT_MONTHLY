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
                        WriteRISO.GetWriteRISO(worksheet, listDD, listErr, date, typeName);
                    }
                    if (type.oki == true)
                    {
                        typeName = ConfigurationManager.AppSettings["OKIDENKI"];
                        worksheet = package.Workbook.Worksheets["OKIDENKI"];
                        WriteOKI.GetWriteOKI(worksheet, listDD, listErr, date, typeName);
                    }
                    if (type.kyo == true)
                    {
                        typeName = ConfigurationManager.AppSettings["KYOCERA"];
                        worksheet = package.Workbook.Worksheets["KYOCERA"];
                        WriteKYO.GetKYOCERA(worksheet, listDD, listErr, typeName);
                    }
                    if (type.km == true)
                    {
                        typeName = ConfigurationManager.AppSettings["KM"];
                        worksheet = package.Workbook.Worksheets["KM"];
                        WriteKM.GetKM(worksheet, listDD, listErr, typeName);
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

        public static QtyErrType CalculateQtyErrType(List<DataError> listError, string typeKH)
        {
            QtyErrType value = new QtyErrType();

            int i = 0;
            value.SetAll0();
            var listAll = listError.Where(p => typeKH.Contains(p.KH)).ToList();
            foreach (var error in listAll)
            {
                var qtyTypeProperties = typeof(QtyErrType).GetProperties();
                foreach (var qtyTypeProperty in qtyTypeProperties)
                {
                    var errorQtyTypeValue = (long)qtyTypeProperty.GetValue(error.QtyType);
                    var itemQtyTypeValue = (long)qtyTypeProperty.GetValue(value);
                    qtyTypeProperty.SetValue(value, itemQtyTypeValue + errorQtyTypeValue);
                }
                i++;
            }

            return value;
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
