using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetControl_API
{
    #region #samplecomparer
    class SampleComparer : IComparer<DevExpress.Spreadsheet.CellValue>
    {
        public int Compare(DevExpress.Spreadsheet.CellValue a, DevExpress.Spreadsheet.CellValue b)
        {
            if (!a.IsText || !b.IsText) return 0;
            if (a.TextValue.Length == b.TextValue.Length) return 0;
            return (a.TextValue.Length > b.TextValue.Length) ? 1 : -1;
        }
    }
    #endregion #samplecomparer
}
