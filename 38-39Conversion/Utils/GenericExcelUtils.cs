using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _38_39Conversion.Utils
{
    public class GenericExcelUtils
    {
        public static int GetLineCount(string text, int columnWidth, string type = "")
        {
            if(type=="xls")
            {
                columnWidth = (columnWidth / 220);
            }
            var lineCount = 1;
            var textPosition = 0;

            while (textPosition <= text.Length)
            {
                textPosition = Math.Min(textPosition + columnWidth, text.Length);
                if (textPosition == text.Length)
                    break;

                if (text[textPosition - 1] == ' ' || text[textPosition] == ' ')
                {
                    lineCount++;
                    textPosition++;
                }
                else
                {
                    textPosition = text.LastIndexOf(' ', textPosition) + 1;

                    var nextSpaceIndex = text.IndexOf(' ', textPosition);
                    if (nextSpaceIndex - textPosition >= columnWidth)
                    {
                        lineCount += (nextSpaceIndex - textPosition) / columnWidth;
                        textPosition = textPosition + columnWidth;
                    }
                    else
                        lineCount++;
                }
            }

            return lineCount;
        }

        public static double getRangeWidth(List<double> colWidths, int start, int end)
        {
            double width = 0;
            for (int i = start; i <= end; ++i)
            {
                width += colWidths[i];
            }
            return width;
        }
    }
}
