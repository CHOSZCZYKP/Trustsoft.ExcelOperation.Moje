using Soneta.EwidencjaVat;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public class SyncfusionHelper
    {
        /// <summary>
        /// Returns the ExcelBordersIndex array and information whether somthing was returned.
        /// </summary>
        /// <param name="borderIndex">The index of the border to be set. This can be one of the values from the <see cref="BorderIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Returns the ExcelBordersIndex array and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static ICollection<ExcelBordersIndex> ConvertFromBordexIndexSyncfusion(BorderIndex borderIndex, out bool isEmpty)
        {
            switch (borderIndex)
            {
                case BorderIndex.Left:
                    isEmpty = false;
                    return new[] { ExcelBordersIndex.EdgeLeft };
                case BorderIndex.Right:
                    isEmpty = false;
                    return new[] { ExcelBordersIndex.EdgeRight };
                case BorderIndex.Top:
                    isEmpty = false;
                    return new[] { ExcelBordersIndex.EdgeTop};
                case BorderIndex.Bottom:
                    isEmpty = false;
                    return new[] { ExcelBordersIndex.EdgeBottom };
                case BorderIndex.All:
                    isEmpty = false;
                    return new[] { ExcelBordersIndex.EdgeLeft, ExcelBordersIndex.EdgeRight, ExcelBordersIndex.EdgeTop , ExcelBordersIndex.EdgeBottom};
                case BorderIndex.None:
                    break;
                default:
                    throw new NotImplementedException();
            }
            isEmpty = true;
            return new List<ExcelBordersIndex>(); 
            
        }

        /// <summary>
        /// Returns the ExcelLineStyle and information whether somthing was returned.
        /// </summary>
        /// <param name="linesIndex">The index of the line style to be set. This can be one of the values from the <see cref="LinesIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Returns ExcelLineStyle and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static ExcelLineStyle ConvertFromLineStyleSyncfusion(LinesIndex linesIndex, out bool isEmpty)
        {
            switch (linesIndex)
            {
                case LinesIndex.Thin:
                    isEmpty = false;
                    return ExcelLineStyle.Thin;

                case LinesIndex.Thick:
                    isEmpty = false;
                    return ExcelLineStyle.Thick;
                case LinesIndex.Medium:
                    isEmpty = false;
                    return ExcelLineStyle.Medium;
                case LinesIndex.Hair:
                    isEmpty = false;
                    return ExcelLineStyle.Hair;
                case LinesIndex.Dotted:
                    isEmpty = false;
                    return ExcelLineStyle.Dotted;
                case LinesIndex.Dashed:
                    isEmpty = false;
                    return ExcelLineStyle.Dashed;
                case LinesIndex.Double:
                    isEmpty = false;
                    return ExcelLineStyle.Double;
                case LinesIndex.None:
                    isEmpty = false;
                    return ExcelLineStyle.None;
                default:
                    isEmpty = true;
                    throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Returns the excelHAlign and information whether somthing was returned.
        /// </summary>
        /// <param name="horizontalAlignmentIndex">The index of the horizontal alignment to be set. This can be one of the values from the <see cref="HorizontalAlignmentIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Return horizontal alignment and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static ExcelHAlign ConvertFromHAlign(HorizontalAlignmentIndex horizontalAlignmentIndex, out bool isEmpty)
        {
            switch(horizontalAlignmentIndex)
            {
                case HorizontalAlignmentIndex.General:
                    isEmpty = false;
                    return ExcelHAlign.HAlignGeneral;
                case HorizontalAlignmentIndex.Left:
                    isEmpty = false;
                    return ExcelHAlign.HAlignLeft;
                case HorizontalAlignmentIndex.Right:
                    isEmpty = false;
                    return ExcelHAlign.HAlignRight;
                case HorizontalAlignmentIndex.Center:
                    isEmpty = false;
                    return ExcelHAlign.HAlignCenter;
                case HorizontalAlignmentIndex.Fill:
                    isEmpty = false;
                    return ExcelHAlign.HAlignFill;
                case HorizontalAlignmentIndex.Justify:
                    isEmpty = false;
                    return ExcelHAlign.HAlignJustify;
                case HorizontalAlignmentIndex.Distributed:
                    isEmpty = false;
                    return ExcelHAlign.HAlignDistributed;
                default: 
                    isEmpty = true;
                    throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Returns the excelVAlign and information whether somthing was returned.
        /// </summary>
        /// <param name="verticalAlignmentIndex">The index of the vertical alignment to be set. This can be one of the values from the <see cref="VerticalAlignmentIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Return vertical alignment and information whether somthing was returned.</returns>
        /// <exception cref="NotFiniteNumberException">Not Implemented Exceotion.</exception>
        public static ExcelVAlign ConvertFromVAlign(VerticalAlignmentIndex verticalAlignmentIndex, out bool isEmpty)
        {
            switch(verticalAlignmentIndex)
            {
                case VerticalAlignmentIndex.Top:
                    isEmpty = false;
                    return ExcelVAlign.VAlignTop;
                case VerticalAlignmentIndex.Bottom:
                    isEmpty = false;
                    return ExcelVAlign.VAlignBottom;
                case VerticalAlignmentIndex.Center:
                    isEmpty = false;
                    return ExcelVAlign.VAlignCenter;
                case VerticalAlignmentIndex.Justify:
                    isEmpty = false;
                    return ExcelVAlign.VAlignJustify;
                case VerticalAlignmentIndex.Distributed:
                    isEmpty = false;
                    return ExcelVAlign.VAlignDistributed;
                default:
                    isEmpty = true;
                    throw new NotFiniteNumberException();
            }
        }
    }
}
