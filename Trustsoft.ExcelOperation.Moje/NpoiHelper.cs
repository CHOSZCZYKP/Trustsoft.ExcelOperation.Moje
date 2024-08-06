using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IFont = NPOI.SS.UserModel.IFont;

namespace Trustsoft.ExcelOperation.Moje
{
    public class NpoiHelper
    {
        /// <summary>
        /// Returns the BorderStyle and information whether somthing was returned.
        /// </summary>
        /// <param name="linesIndex">The index of the line style to be set. This can be one of the values from the <see cref="LinesIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Returns BorderStyle and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static BorderStyle ConvertFromLineStyleNpoi(LinesIndex linesIndex, out bool isEmpty)
        {
            switch (linesIndex)
            {
                case LinesIndex.Thin:
                    isEmpty = false;
                    return BorderStyle.Thin;
                case LinesIndex.Thick:
                    isEmpty = false;
                    return BorderStyle.Thick;
                case LinesIndex.Medium:
                    isEmpty = false;
                    return BorderStyle.Medium;
                case LinesIndex.Hair:
                    isEmpty = false;
                    return BorderStyle.Hair;
                case LinesIndex.Dotted:
                    isEmpty = false;
                    return BorderStyle.Dotted;
                case LinesIndex.Dashed:
                    isEmpty = false;
                    return BorderStyle.Dashed;
                case LinesIndex.Double:
                    isEmpty = false;
                    return BorderStyle.Double;
                case LinesIndex.None:
                    isEmpty = false;
                    return BorderStyle.None;
                default:
                    isEmpty = true;
                    throw new NotImplementedException();
            }
            
            
        }

        /// <summary>
        /// Returns the HorizontalAlignment and information whether somthing was returned.
        /// </summary>
        /// <param name="horizontalAligmentIndex">The index of the horizontal alignment to be set. This can be one of the values from the <see cref="HorizontalAlignmentIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Return HorizontalAlignment and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static HorizontalAlignment ConverFromHorizontalAlignmentNpoi(HorizontalAlignmentIndex horizontalAligmentIndex, out bool isEmpty)
        {
            switch(horizontalAligmentIndex)
            {
                case HorizontalAlignmentIndex.General:
                    isEmpty = false;
                    return HorizontalAlignment.General;
                case HorizontalAlignmentIndex.Left:
                    isEmpty = false;
                    return HorizontalAlignment.Left;
                case HorizontalAlignmentIndex.Right:
                    isEmpty = false;
                    return HorizontalAlignment.Right;
                case HorizontalAlignmentIndex.Center:
                    isEmpty = false;
                    return HorizontalAlignment.Center;
                case HorizontalAlignmentIndex.Fill:
                    isEmpty = false;
                    return HorizontalAlignment.Fill;
                case HorizontalAlignmentIndex.Justify:
                    isEmpty = false;
                    return HorizontalAlignment.Justify;
                case HorizontalAlignmentIndex.Distributed:
                    isEmpty = false;
                    return HorizontalAlignment.Distributed;
                default:
                    isEmpty = true;
                    throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Returns the VerticalAlignment and information whether somthing was returned.
        /// </summary>
        /// <param name="verticalAlignmentIndex">The index of the vertical alignment to be set. This can be one of the values from the <see cref="VerticalAlignmentIndex"/>enum.</param>
        /// <param name="isEmpty">Returns true if the worksheet is empty otherwise false.</param>
        /// <returns>Return VerticalAlignment and information whether somthing was returned.</returns>
        /// <exception cref="NotImplementedException">Not Implemented Exceotion.</exception>
        public static VerticalAlignment ConverFromVerticalAligmentNpoi(VerticalAlignmentIndex verticalAlignmentIndex, out bool isEmpty)
        {
            switch(verticalAlignmentIndex)
            {
                case VerticalAlignmentIndex.Top:
                    isEmpty = false;
                    return VerticalAlignment.Top;
                case VerticalAlignmentIndex.Bottom:
                    isEmpty = false;
                    return VerticalAlignment.Bottom;
                case VerticalAlignmentIndex.Center:
                    isEmpty = false;
                    return VerticalAlignment.Center;
                case VerticalAlignmentIndex.Justify:
                    isEmpty = false;
                    return VerticalAlignment.Justify;
                case VerticalAlignmentIndex.Distributed:
                    isEmpty = false;
                    return VerticalAlignment.Distributed;
                default:
                    isEmpty = true; 
                    throw new NotImplementedException();
            }
        }

        
    }
}
