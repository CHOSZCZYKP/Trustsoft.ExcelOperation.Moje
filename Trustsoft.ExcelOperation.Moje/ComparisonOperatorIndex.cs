using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public enum ComparisonOperatorIndex
    {
        /// <summary>
        /// No comparison operator.
        /// </summary>
        None = 0,

        /// <summary>
        /// The comparison operator between.
        /// </summary>
        Between = 1,

        /// <summary>
        /// The comparison operator not between.
        /// </summary>
        NotBetween = 2,

        /// <summary>
        /// The comparison operator equal.
        /// </summary>
        Equal = 3,

        /// <summary>
        /// The comparison operator not equal.
        /// </summary>
        NotEqual = 4,

        /// <summary>
        /// The comparison operator greater than.
        /// </summary>
        GreaterThan = 5,

        /// <summary>
        /// The comparison operator less than.
        /// </summary>
        LessThan = 6,

        /// <summary>
        /// The comparison operator greater than or equal.
        /// </summary>
        GreaterThanOrEqual = 7,

        /// <summary>
        /// The comparison operator less than or equal.
        /// </summary>
        LessThanOrEqual = 8
    }
}
