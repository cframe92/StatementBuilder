using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class SortTableRow
    {
       public SortTableRow()
        {
            Column = new List<string>();
        }

       public List<string> Column
       {
           get;
           set;
       }
    }
}
