using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    internal class InvalidFormulaException : Exception
    {
        public char InvalidChar {  get; set; }  
        public InvalidFormulaException(string msg, char invalidChar) : base(msg)
        { 
            InvalidChar = invalidChar;
        }
    }
}
