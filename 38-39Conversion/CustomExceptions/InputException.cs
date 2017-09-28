using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _38_39Conversion.CustomExceptions
{
    public class InputException : System.Exception
    {
        public InputException()
        {
        }

        public InputException(string message)
        : base(message)
        {
        }

        public InputException(string message, Exception inner)
        : base(message, inner)
        {
        }
    }
}
