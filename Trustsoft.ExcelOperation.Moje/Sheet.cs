using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public class Sheet
    {
        /// <summary>
        /// Sheet index in the workbook.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Sheet name in the workbook.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Creates a sheet which represents the index and name of the sheet.
        /// </summary>
        /// <param name="index">Sheet index.</param>
        /// <param name="name">Sheet name.</param>
        public Sheet(int index, string name)
        {
            this.Index = index;
            this.Name = name;
        }

        /// <summary>
        /// Returns the full index and name of the sheet.
        /// </summary>
        /// <returns>Full index and name if the sheet.</returns>
        public override string ToString()
        {
            return $"{Index} {Name}";
        }

    }
}
