using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class Border
    {
        public Border(float widthLeft, float widthTop, float widthRight, float widthBottom)
        {
            WidthLeft = widthLeft;
            WidthTop = widthTop;
            WidthRight = widthRight;
            WidthBottom = widthBottom;
        }

        public float WidthLeft
        {
            get;
            set;
        }

        public float WidthTop
        {
            get;
            set;
        }

        public float WidthRight
        {
            get;
            set;
        }

        public float WidthBottom
        {
            get;
            set;
        }
    }
}
