using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    internal class InvalidFormulaException : Exception
    {
        public int InvalidChar { get; set; } 
        
        public InvalidFormulaException(string msg, int invalidChar) : base(msg)
        { 
            InvalidChar = invalidChar;
        }

    }
}
