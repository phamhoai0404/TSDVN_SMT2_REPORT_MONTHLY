using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.MODEL
{
    public class MdlCommon
    {
        public static string PATH_TEMPLATE  = ConfigurationManager.AppSettings["PathFileTemplate"];
        public const string TYPE_FILE_SELECT = "Excel file: |*.xlsx;*xlsm";

        public static string STRING_IT_THIEC;
        public static string STRING_HAN_GIA;
        public static string STRING_SAI_VITRI;
        public static string STRING_KENH;
        public static string STRING_BAC_CAU;
        public static string STRING_THIEU_LK;
        public static string STRING_LAT_NGUOC;
        public static string STRING_NGUOC_HUONG;
        public static string STRING_NHAM_LK;
        public static string STRING_DI_VAT;
        public static string STRING_THUA_LK;
        public static string STRING_BONG;
        public static string STRING_LECH;
        public static string STRING_VO;
        public static string STRING_DUNG_DUNG;
    }
    public class RESULT
    {
        public const string OK = "OK";
        public const string ERROR_015_CATCH = "Lỗi {0} Catch - {1}";
        public const string ERROR_DATA = "Cell {1}: {2} => Không phải dữ liệu {0}";
    }
}
