using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class ActionLoi
    {

        public static void GetValueError(ref List<DataError> listDataErr, DataConfigLoi configErr, DateTime monthGet)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(configErr.pathFile), false))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[configErr.sheetName];
                    if (worksheet == null)
                        throw new Exception($"File: {configErr.pathFile} ==== SheetName:{configErr.sheetName} => Không tồn tại!");

                    int lastRow = worksheet.Dimension.End.Row; // Lấy dòng cuối cùng
                    object[,] listAll = worksheet.Cells[$"A1:{configErr.ColLast}{lastRow}"].Value as object[,];
                    int rowAll = listAll.GetLength(0);

                    DataError tempValue = new DataError();
                    long numberQty = 0;
                    for (int i = 0; i < rowAll; i++)
                    {
                        if (!(listAll[i, configErr.Model.Index] is string) ||
                            string.IsNullOrWhiteSpace(listAll[i, configErr.Model.Index].ToString()) ||
                            !listAll[i, configErr.Model.Index].ToString().Contains("-"))
                            continue;

                        tempValue.Model = listAll[i, configErr.Model.Index].ToString().Trim().ToUpper();//Lay  gia tri cua model
                        if (tempValue.Model.Length < 9)
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "Model", configErr.Model.ColName + (i + 1), tempValue.Model));

                        tempValue.Model = tempValue.Model.Substring(0, 9);

                        //Thuc hien lay so luong
                        if (!long.TryParse(listAll[i, configErr.QtyError.Index]?.ToString(), out numberQty))
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "QTY Error", configErr.QtyError.ColName + (i + 1), listAll[i, configErr.QtyError.Index]));

                        if (numberQty <= 0)//Neu so luong <= 0 thi duyet cai khac
                            continue;

                        //Kiem tra khach hang xem co du lieu hay khong neu khong co du lieu thi bao loi luon
                        if (!(listAll[i, configErr.KH.Index] is string) || string.IsNullOrWhiteSpace(listAll[i, configErr.KH.Index].ToString()))
                            throw new Exception(string.Format(RESULT.ERROR_DATA, "Khách hàng", configErr.KH.ColName + (i + 1), listAll[i, configErr.KH.Index]));
                        tempValue.KH = listAll[i, configErr.KH.Index]?.ToString().Trim().ToUpper();
                        tempValue.QtyError = numberQty;
                        tempValue.Mat = listAll[i, configErr.Mat.Index]?.ToString().Trim().ToUpper();
                        tempValue.TypeItem = listAll[i, configErr.TypeItem.Index]?.ToString().Trim();
                        tempValue.Content = listAll[i, configErr.Content.Index]?.ToString().Trim();

                        object value = listAll[i, configErr.Content_Error_KH.Index];
                        if (!(value is string) || string.IsNullOrWhiteSpace(value.ToString()))
                            tempValue.ContentKH = "";
                        else
                            tempValue.ContentKH = value?.ToString().ToUpper();

                        listDataErr.Add(new DataError(tempValue));
                    }
                    if (listDataErr.Count == 0)
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(RESULT.ERROR_015_CATCH, $"GetValueError", ex.Message));
            }
        }

        public static void ParingError(List<DataDD> listDD, ref List<DataError> listError)
        {
            ////NEU MA THAT SU THI BO COMMENT
            ////Kiem tra model ma khong ton tai trong listDD thi bao lỗi
            var processedModels = new HashSet<string>();
            foreach (var item in listError)
            {
                if (processedModels.Contains(item.Model))
                    continue;
                //LUC THAT THI ENABLE CAI NAY
                if (!listDD.Any(p => p.Model == item.Model))
                {
                    string temp = $"Cần xem lại lỗi ở Model: {item.ToString()} => Không tồn tại Model trong danh sách điểm dán hoặc số lượng của Model trong file Điểm dán = 0! => Bạn có muốn tiếp tục?";
                    DialogResult result = MessageBox.Show(temp, "Cảnh báo!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No)
                    {
                        throw new Exception("Bạn đã dừng chương trình!");
                    }
                }
                processedModels.Add(item.Model);
            }

            int colRISO = MdlCommon.LIST_ERROR_RISO.Count();
            int colOKI = MdlCommon.LIST_ERROR_OKIDENKI.Count();
            int colKYO = MdlCommon.LIST_ERROR_KYOCERA.Count();
            int colKM = MdlCommon.LIST_ERROR_KM.Count();
            foreach (var error in listError)
            {
                var tempError = error;
                switch (error.KH)
                {
                    case string s when MdlCommon.TYPE_NAME_OKIDENKI.Contains(s):
                        GetTypeError(ref tempError, MdlCommon.LIST_ERROR_OKIDENKI, colOKI);
                        break;
                    case string s when MdlCommon.TYPE_NAME_KYOCERA.Contains(s):
                        GetTypeError(ref tempError, MdlCommon.LIST_ERROR_KYOCERA, colKYO);
                        break;
                    case string s when MdlCommon.TYPE_NAME_RISO.Contains(s):
                        GetTypeError(ref tempError, MdlCommon.LIST_ERROR_RISO, colRISO);
                        break;
                    case string s when MdlCommon.TYPE_NAME_KM.Contains(s):
                        GetTypeError(ref tempError, MdlCommon.LIST_ERROR_KM, colKM);
                        break;
                    default:
                        throw new Exception($"Thông tin Item lỗi: {error.ToString()} => Không tồn tại trong 4 khách hàng!");
                }
            }


        }
        private static void GetTypeError(ref DataError value, List<TitleError> listTitle, int col)
        {
            //Neu du lieu loi khong co thi mac dinh se la khac
            if (string.IsNullOrWhiteSpace(value.ContentKH))
            {
                for (int i = 0; i < col - 1; i++)
                {
                    value.listErr.Add(0);
                }
                value.listErr.Add(value.QtyError);
                return;
            }

            bool checkType = false;
            for (int i = 0; i < col; i++)
            {
                if (!checkType && listTitle[i].NameTitle.Contains(value.ContentKH))
                {
                    value.listErr.Add(value.QtyError);
                    checkType = true;
                }
                else
                {
                    value.listErr.Add(0);
                }
            }

            if(checkType == false)
            {
                value.listErr[col -1 ] = value.QtyError;//Gan phan tu cuoi cung bang du lieu
            }
        }
        
    }
}
