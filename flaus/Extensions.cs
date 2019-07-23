using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace peteli.flaus
{
    internal static class FlausExtensions
    {
        internal static FlausModel Flaus(this Excel.Workbook Wb)
        {
            if (MyAddIn.FlausModelByWorkbook.ContainsKey(Wb))
            {
                return MyAddIn.FlausModelByWorkbook[Wb];
            }
            else
            {
                return null;
            }
        }
    }
}
