using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class MyFunction1
    {
        public static int ConvertNameToNumer(string columnName)
        {
            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += columnName[i] - 'A' + 1;
            }
            return sum;
        }
        public static string ConvertNumberToName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static bool IsDate(object value, ref DateTime dateTime)
        {
            if (value is DateTime dateValue)
            {
                dateTime = dateValue;
                return true;
            }

            if (value is double numericValue)
            {
                dateTime = DateTime.FromOADate(numericValue);
                return true;
            }
            return false;
        }
    }
}
