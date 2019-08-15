using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CombinationOfExcelColumns
{
    public class FilledColumnAndRows
    {
        public FilledColumnAndRows()
        {
            RowsOfFilledColumn = new List<string>();
        }

        public int FilledColumnIndeks { get; set; }
        public List<string> RowsOfFilledColumn { get; set; }
    }
}
