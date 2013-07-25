using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class SortTableItem
    {
        public SortTableItem()
        {
            Date = new DateTime();
        }

        public SortTableItem(string description, decimal amount, DateTime date)
        {
            Description = description;
            Amount = amount;
            Date = date;
        }

        public string Description
        {
            get;
            set;
        }

        public decimal Amount
        {
            get;
            set;
        }

        public DateTime Date
        {
            get;
            set;
        }
    }
}
