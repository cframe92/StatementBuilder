using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class Withdrawal : SortTableItem
    {
        public Withdrawal()
            : base()
        {

        }

        public Withdrawal(string description, decimal amount, DateTime date)
            : base(description, amount, date)
        {

        }
    }
}
