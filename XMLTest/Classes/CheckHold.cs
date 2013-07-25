using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class CheckHold
    {
        public CheckHold()
        {
            EffectiveDate = new DateTime();
            ExpiredDate = new DateTime();
        }

        public DateTime EffectiveDate
        {
            get;
            set;

        }

        public DateTime ExpiredDate
        {
            get;
            set;
        }

        public decimal Amount
        {
            get;
            set;
        }
    }
}
