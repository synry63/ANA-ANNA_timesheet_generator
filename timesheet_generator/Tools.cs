using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace timesheet_generator
{
    internal static class Tools
    {
        public static Dictionary<CellValues, dynamic> FormatValString(string objVal)
        {

            Dictionary<CellValues, dynamic> dict = new Dictionary<CellValues, dynamic>();
            DateTime date;
            int intValue;
            if (DateTime.TryParse(objVal, out date))
            {
                dict.Add(CellValues.String, date.ToString("dd/MM/yyyy"));
            }
            else if (int.TryParse(objVal, out intValue))
            {
                dict.Add(CellValues.Number, intValue);
            }
            else
            {
                dict.Add(CellValues.String, objVal);
               
            }
            return dict;
        }
    }
}
