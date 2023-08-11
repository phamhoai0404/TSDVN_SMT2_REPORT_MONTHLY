using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class MyFuntion2
    {
        public static string SelectFile()
        {
            try
            {
                using (var ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Filter = MdlCommon.TYPE_FILE_SELECT;
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string s = ofd.FileName;
                        int index = s.IndexOf(':') + 1;
                        string rootPath = GetUNCPath(s.Substring(0, index));
                        string directory = s.Substring(index);
                        return rootPath + directory;
                    }
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }


        }

       
        /// <summary>
        /// Chuyen doi tu (vd P: => \\192.168.3.6\public)
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string GetUNCPath(string path)
        {
            try
            {
                if (path.StartsWith(@"\\"))
                {
                    return path;
                }

                ManagementObject mo = new ManagementObject();
                mo.Path = new ManagementPath(String.Format("Win32_LogicalDisk='{0}'", path));

                // DriveType 4 = Network Drive
                if (Convert.ToUInt32(mo["DriveType"]) == 4)
                {
                    return Convert.ToString(mo["ProviderName"]);
                }
                else
                {
                    return path;
                }
            }
            catch (Exception)
            {
                return path;
            }
        }

    }
}

