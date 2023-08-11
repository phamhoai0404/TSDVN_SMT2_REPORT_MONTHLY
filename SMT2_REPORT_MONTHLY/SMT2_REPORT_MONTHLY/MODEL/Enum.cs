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
    }
    public class RESULT
    {
        public const string OK = "OK";
        public const string ERROR_015_CATCH = "Lỗi {0} Catch - {1}";
        public const string ERROR_DATA = "Cell {1}: {2} => Không phải dữ liệu {0}";
    }
}
