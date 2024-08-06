using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje.SyncfusionException
{
    internal class SyncfusionNullApplicationException : Exception
    {
        public SyncfusionNullApplicationException() { }

        public new string Message = "The application did not start correctly";
    }
}
