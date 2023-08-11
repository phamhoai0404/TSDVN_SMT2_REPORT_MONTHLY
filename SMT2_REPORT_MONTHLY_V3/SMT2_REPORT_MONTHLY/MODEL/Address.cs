using QA_TVN2_REPORT_MONTHLY.FUNCTION;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_TVN2_REPORT_MONTHLY.MODEL
{
    public class Address
    {
        public int Index { get; set; }
        public string ColName { get; set; }
        public Address()
        {

        }
        public Address(Address s)
        {
            this.Index = s.Index;
            this.ColName = s.ColName;
        }
        public void GetIndexColumn()
        {
            this.Index = MyFunction1.ConvertNameToNumer(this.ColName) - 1;//Tru 1 vi de tinh toan se bat dau = 0
        }
        public void GetNameColumn()
        {
            this.ColName = MyFunction1.ConvertNumberToName(this.Index + 1);
        }

    }
    public class TitleError
    {
        public Address Address { get; set; }
        public string  NameTitle { get; set; }

        public TitleError()
        {
            this.Address = new Address();
        }
        public TitleError(TitleError s)
        {
            this.NameTitle = s.NameTitle;
            this.Address = new Address(s.Address);
        }
    }

}
