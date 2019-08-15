using System.Collections.Generic;

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
