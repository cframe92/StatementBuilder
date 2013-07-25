using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class Advertisement
    {
        public Advertisement()
        {
            MessageLines = new string[MAX_MESSAGE_LINES];

            for (int i = 0; i < MessageLines.Count(); i++)
            {
                MessageLines[i] = string.Empty;
            }
        }

        public string[] MessageLines
        {
            set;
            get;
        }

        public int TotalLines
        {
            get
            {
                int totalLines = 0;

                for (int i = 0; i < MessageLines.Count(); i++)
                {
                    if (MessageLines[i] != string.Empty)
                    {
                        totalLines = i + 1;
                    }
                }

                return totalLines;
            }
        }

        public const int MAX_MESSAGE_LINES = 20;
    }
}
