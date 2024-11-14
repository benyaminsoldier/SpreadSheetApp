using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    public class ColorChosenEventArgs : EventArgs
    {
        public Color ChosenColor { get; set; }          
    }
}
