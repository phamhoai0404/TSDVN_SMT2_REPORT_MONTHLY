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
    public class ActionGetConfig
    {
        public static void GetConfigAll()
        {
            string riso = "RISO";
            string kyo = "KYOCERA";
            string km = "KM";
            string oki = "OKIDENKI";

            MdlCommon.LIST_ERROR_RISO = new List<TitleError>();
            GetConfigTitleError(ref MdlCommon.LIST_ERROR_RISO, riso);

            MdlCommon.LIST_ERROR_KM = new List<TitleError>();
            GetConfigTitleError(ref MdlCommon.LIST_ERROR_KM, km);

            MdlCommon.LIST_ERROR_OKIDENKI = new List<TitleError>();
            GetConfigTitleError(ref MdlCommon.LIST_ERROR_OKIDENKI, oki);

            MdlCommon.LIST_ERROR_KYOCERA = new List<TitleError>();
            GetConfigTitleError(ref MdlCommon.LIST_ERROR_KYOCERA, kyo);

            MdlCommon.TYPE_NAME_RISO = ConfigurationManager.AppSettings[riso];
            MdlCommon.TYPE_NAME_KYOCERA = ConfigurationManager.AppSettings[kyo];
            MdlCommon.TYPE_NAME_KM = ConfigurationManager.AppSettings[km];
            MdlCommon.TYPE_NAME_OKIDENKI = ConfigurationManager.AppSettings[oki];
            MdlCommon.TYPE_NAME_ALL = ConfigurationManager.AppSettings["ALL"];
            MdlCommon.TYPE_NAME_DATA_ERROR = ConfigurationManager.AppSettings["DATA_ERROR"];

            GetNameTitle(riso, kyo, km, oki);
        }
        private static void GetConfigTitleError(ref List<TitleError> listTitleErr, string typeKH)
        {
            listTitleErr = new List<TitleError>();
            TitleError tempValue = new TitleError();
            tempValue.Address.ColName = ConfigurationManager.AppSettings[$"{typeKH}_COL_START"];
            tempValue.Address.GetIndexColumn();

            TitleError tempValueEnd = new TitleError();
            tempValueEnd.Address.ColName = ConfigurationManager.AppSettings[$"{typeKH}_COL_END"];
            tempValueEnd.Address.GetIndexColumn();

            listTitleErr.Add(new TitleError(tempValue));//Thuc hien add cai dau 
            for (int i = tempValue.Address.Index + 1; i < tempValueEnd.Address.Index; i++)//Thuc hien add nhung cai giua
            {
                tempValue.Address.Index = i;
                tempValue.Address.GetNameColumn();
                listTitleErr.Add(new TitleError(tempValue));
            }
            listTitleErr.Add(new TitleError(tempValueEnd));//Add Cuoi 
        }
        private static void GetNameTitle(string riso, string kyo, string km, string oki)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(MdlCommon.PATH_TEMPLATE), false))
            {
               
                ExcelWorksheet worksheet = package.Workbook.Worksheets[riso];
                if (worksheet == null)
                    throw new Exception($"File Template: {MdlCommon.PATH_TEMPLATE} ==== SheetName:{riso} => Không tồn tại!");
                GetTitleChild(worksheet, ref MdlCommon.LIST_ERROR_RISO);

                worksheet = package.Workbook.Worksheets[oki];
                if (worksheet == null)
                    throw new Exception($"File Template: {MdlCommon.PATH_TEMPLATE} ==== SheetName:{oki} => Không tồn tại!");
                GetTitleChild(worksheet, ref MdlCommon.LIST_ERROR_OKIDENKI);

                worksheet = package.Workbook.Worksheets[kyo];
                if (worksheet == null)
                    throw new Exception($"File Template: {MdlCommon.PATH_TEMPLATE} ==== SheetName:{kyo} => Không tồn tại!");
                GetTitleChild(worksheet, ref MdlCommon.LIST_ERROR_KYOCERA);

                worksheet = package.Workbook.Worksheets[km];
                if (worksheet == null)
                    throw new Exception($"File Template: {MdlCommon.PATH_TEMPLATE} ==== SheetName:{km} => Không tồn tại!");
                GetTitleChild(worksheet, ref MdlCommon.LIST_ERROR_KM);

            }
        }

        private static void GetTitleChild(ExcelWorksheet ws, ref List<TitleError> listTilteKH)
        {
            string colStart = listTilteKH[0].Address.ColName;
            string colEnd = listTilteKH[listTilteKH.Count() - 1].Address.ColName;
            object[,] listData = ws.Cells[$"{colStart}{MdlCommon.ROW_TITTLE_GET_TEMPLATE}:{colEnd}{MdlCommon.ROW_TITTLE_GET_TEMPLATE}"].Value as object[,];

            for (int i = 0; i < listTilteKH.Count(); i++)
            {
                listTilteKH[i].NameTitle = listData[0, i]?.ToString().ToUpper();
            }
        }
        public static void GetConfig(ref DataConfigLoi config)
        {
            config.Model.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_MODEL"];
            config.KH.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_KH"];
            config.Mat.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_MAT"];
            config.QtyError.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_QTYERROR"];
            config.TypeItem.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_TYPEITEM"];
            config.Content.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_CONTENT"];
            config.Content_Error_KH.ColName = ConfigurationManager.AppSettings["FILE_ERROR_COL_NAME_ERR_KH"];

            config.Model.GetIndexColumn();
            config.KH.GetIndexColumn();
            config.Mat.GetIndexColumn();
            config.QtyError.GetIndexColumn();
            config.TypeItem.GetIndexColumn();
            config.Content.GetIndexColumn();
            config.Content_Error_KH.GetIndexColumn();

            int maxVariable = Math.Max(config.Model.Index, Math.Max(config.KH.Index, Math.Max(config.Mat.Index, Math.Max(config.QtyError.Index, Math.Max(config.TypeItem.Index, Math.Max(config.Content_Error_KH.Index, config.Content.Index))))));
            config.ColLast = MyFunction1.ConvertNumberToName(maxVariable + 1);//vi no dang bi tru di 1 nen la
        }
        public static void GetConfig(ref DataConfigDD config)
        {
            config.Mat.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_MAT"];
            config.Model.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_MODEL"];
            config.KH.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_KH"];
            config.Qty.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_QTY"];
            config.Qty.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_QTY"];
            config.PointQty.ColName = ConfigurationManager.AppSettings["FILE_DD_COL_POINTQTY"];

            config.Mat.GetIndexColumn();
            config.Model.GetIndexColumn();
            config.KH.GetIndexColumn();
            config.Qty.GetIndexColumn();
            config.PointQty.GetIndexColumn();

            int maxVariable = Math.Max(config.Model.Index, Math.Max(config.KH.Index, Math.Max(config.Mat.Index, config.Qty.Index)));
            config.ColLast = MyFunction1.ConvertNumberToName(maxVariable + 1);//vi no dang bi tru di 1 nen la
        }

    }
}
