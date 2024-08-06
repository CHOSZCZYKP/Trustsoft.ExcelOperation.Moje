using System;
using System.Collections.Generic;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public enum BorderIndex
    {
        /// <summary>
        /// No broder.
        /// </summary>
        None = 0,

        /// <summary>
        /// The top border.
        /// </summary>
        Top = 1,

        /// <summary>
        /// The bottom border.
        /// </summary>
        Bottom = 2,

        /// <summary>
        /// The left border.
        /// </summary>
        Left = 3,

        /// <summary>
        /// The right border.
        /// </summary>
        Right = 4,

        /// <summary>
        /// All borders.
        /// </summary>
        All = 5
    }
}
