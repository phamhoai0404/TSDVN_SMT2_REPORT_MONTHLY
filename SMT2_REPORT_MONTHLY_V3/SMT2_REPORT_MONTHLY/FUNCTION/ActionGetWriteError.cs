using QA_TVN2_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.FUNCTION
{
    public class ActionGetWriteError
    {
        public static List<DataError2> GetError2(List<DataError> listTemp, string khachhang, int colTitle)
        {
            //Chi hien thi du lieu khac
            return listTemp.Where(p => khachhang.Contains(p.KH) && p.listErr[colTitle - 1] != 0).GroupBy(p => new { p.Content })
                            .Select(group =>
                            {
                                string contentKH = group.All(p => p.ContentKH == group.First().ContentKH)
                                ? group.First().ContentKH
                                : string.Join(",", group.Select(p => p.ContentKH));

                                return new DataError2
                                {
                                    Content = group.Key.Content,
                                    QtyError = group.Sum(p => p.QtyError),
                                    ContentKH = contentKH,
                                };
                            })
                            .OrderBy(p => p.Model)
                            .ToList();
        }
        public static List<DataError3> GetError3(List<DataError> listTemp, string khachhang, int colTitle)
        {
            //chi hien thi du lieu khac
            return listTemp.Where(p => khachhang.Contains(p.KH) && p.listErr[colTitle - 1] != 0).GroupBy(p => new { p.Content, p.Model, p.TypeItem })
                            .Select(group =>
                            {
                                return new DataError3
                                {
                                    Model = group.Key.Model,
                                    Content = group.Key.Content,
                                    TypeItem = group.Key.TypeItem,
                                    ContentKH = group.First().ContentKH,
                                    QtyError = group.Sum(p => p.QtyError),
                                };
                            })
                            .OrderBy(p => p.Model)
                            .ToList();
        }
    }
}
