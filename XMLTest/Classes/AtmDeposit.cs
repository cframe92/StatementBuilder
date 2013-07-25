using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class AtmDeposit : SortTableItem
    {
        public AtmDeposit()
            : base()
        {

        }

        public AtmDeposit(string description, decimal amount, DateTime date)
            : base(description, amount, date)
        {

        }
    }
}
