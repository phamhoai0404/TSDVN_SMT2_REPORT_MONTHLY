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
                        GetTypeError(listAll[i, configErr.Content_Error_KH.Index], ref tempValue , numberQty);
                        tempValue.KH = listAll[i, configErr.KH.Index]?.ToString().Trim().ToUpper();
                        tempValue.QtyError = numberQty;
                        tempValue.Mat = listAll[i, configErr.Mat.Index]?.ToString().Trim().ToUpper();
                        tempValue.TypeItem = listAll[i, configErr.TypeItem.Index]?.ToString().Trim();
                        tempValue.Content = listAll[i, configErr.Content.Index]?.ToString().Trim();

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

        public static void ParingError(ref List<DataDD> listDD, ref List<DataError> listError)
        {
            //Kiem tra model ma khong ton tai trong listDD thi bao lỗi
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

            var qtyTypeProperties = typeof(QtyErrType).GetProperties();
            foreach (var item in listDD)
            {
                item.QtyError = 0;

                foreach (var error in listError)
                {
                    if (error.Model == item.Model)
                    {
                        item.QtyError += error.QtyError;
                        foreach (var qtyTypeProperty in qtyTypeProperties)
                        {
                            var errorQtyTypeValue = (long)qtyTypeProperty.GetValue(error.QtyType);
                            var itemQtyTypeValue = (long)qtyTypeProperty.GetValue(item.QtyType);
                            qtyTypeProperty.SetValue(item.QtyType, itemQtyTypeValue + errorQtyTypeValue);
                        }
                    }
                }
            }
        }

        private static void GetTypeError( object value , ref DataError input, long data)
        {
            input.QtyType.SetAll0();//Thuc hien set toan bo ve 0
            if (!(value is string) || string.IsNullOrWhiteSpace(value.ToString()))
            {
                input.QtyType.Qty_Khac = data;
                input.ContentKH = "";
                return;
            }
            string tempValue = value?.ToString().ToUpper();
            input.ContentKH = tempValue;
            switch (tempValue)
            {
                case string s when s.Contains(MdlCommon.STRING_IT_THIEC):
                    input.QtyType.Qty_It_Thiec = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_HAN_GIA):
                    input.QtyType.Qty_HanGia = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_SAI_VITRI):
                    input.QtyType.Qty_SaiVitri = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_KENH):
                    input.QtyType.Qty_Kenh = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_BAC_CAU):
                    input.QtyType.Qty_Baccau = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_THIEU_LK):
                    input.QtyType.Qty_Thieu_LK = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_LAT_NGUOC):
                    input.QtyType.Qty_Lat_Nguoc = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_NGUOC_HUONG):
                    input.QtyType.Qty_Nguoc_Huong = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_NHAM_LK):
                    input.QtyType.Qty_Nham_LK = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_DI_VAT):
                    input.QtyType.Qty_Di_Vat = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_THUA_LK):
                    input.QtyType.Qty_Thua_LK = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_BONG):
                    input.QtyType.Qty_Bong = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_LECH):
                    input.QtyType.Qty_Lech= data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_VO):
                    input.QtyType.Qty_Vo = data;
                    break;
                case string s when s.Contains(MdlCommon.STRING_DUNG_DUNG):
                    input.QtyType.Qty_Dung_dung = data;
                    break;
                default:
                    input.QtyType.Qty_Khac = data;
                    return;
            }
        }
    }
}
