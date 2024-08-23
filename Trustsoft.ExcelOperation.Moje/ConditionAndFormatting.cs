using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public class ConditionAndFormatting
    {
        /// <summary>
        /// Get or set comparison operator.
        /// </summary>
        public ComparisonOperatorIndex ComparisonOperatorIndex { get; set; }

        /// <summary>
        /// Get or set a condition for the style.
        /// </summary>
        public string Condition { get; set; }

        /// <summary>
        /// Get or set the alpha component of the background color (0-255).
        /// </summary>
        public int? BackgroundColorA { get; set; }

        /// <summary>
        /// Get or set the red component of the background color (0-255).
        /// </summary>
        public int? BackgroundColorR { get; set; }

        /// <summary>
        /// Get or set the green component of the background color (0-255).
        /// </summary>
        public int? BackgroundColorG { get; set; }

        /// <summary>
        /// Get or set the blue component of the background color (0-255).
        /// </summary>
        public int? BackgroundColorB { get; set; }

        /// <summary>
        /// Gets or sets whether the font is bold.
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Get or set whether the font is italics.
        /// </summary>
        public bool? Italics { get; set; }

        /// <summary>
        /// Get or set whether the font is underline.
        /// </summary>
        public bool? Underline { get; set; }

        /// <summary>
        /// Get or set whether the font is double underline.
        /// </summary>
        public bool? DoubleUnderline { get; set; }

        /// <summary>
        /// Get or set the alpha component of the font color (0-255).
        /// </summary>
        public int? TextColorA { get; set; }

        /// <summary>
        /// Get or set the red component of the font color (0-255).
        /// </summary>
        public int? TextColorR { get; set; }

        /// <summary>
        /// Get or set the green component of the font color (0-255).
        /// </summary>
        public int? TextColorG { get; set; }

        /// <summary>
        /// Get or set the blue component of the font color (0-255).
        /// </summary>
        public int? TextColorB { get; set; }

        /// <summary>
        /// Creates a ConditionAndFormatting object by setting comparison operator and condition.
        /// </summary>
        /// <param name="comparisonOperatorIndex">Comparison operator.</param>
        /// <param name="condition">Condition that must be met to set the format.</param>
        public ConditionAndFormatting(ComparisonOperatorIndex comparisonOperatorIndex, string condition)
        {
            this.ComparisonOperatorIndex = comparisonOperatorIndex;
            this.Condition = condition;
        }

        /// <summary>
        /// Sets parameters ARGB for background.
        /// </summary>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public ConditionAndFormatting SetBackgroundColor(int a, int r, int g, int b)
        {
            this.BackgroundColorA = a; 
            this.BackgroundColorR = r;
            this.BackgroundColorG = g;
            this.BackgroundColorB = b;
            return this;
        }

        /// <summary>
        /// Sets parameters ARGB for font.
        /// </summary>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public ConditionAndFormatting SetTextColor(int a, int r, int g,int b)
        {
            this.TextColorA = a;
            this.TextColorR = r;
            this.TextColorG = g;
            this.TextColorB = b;
            return this;
        }

        /// <summary>
        /// Sets the font to bold.
        /// </summary>
        /// <param name="bold">Indicates whether the font should be bold.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public ConditionAndFormatting SetBold(bool bold)
        {
            this.Bold = bold;
            return this;
        }

        /// <summary>
        /// Sets the font to italics.
        /// </summary>
        /// <param name="italics">Indicates whether the font should be italics.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public ConditionAndFormatting SetItalics(bool italics)
        {
            this.Italics = italics;
            return this;
        }

        /// <summary>
        /// Sets the font to underline.
        /// </summary>
        /// <param name="underline">Indicates whether the font should be underline.</param>
        /// <param name="doubleUnderLine">Indicates whether the font should be double underline.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public ConditionAndFormatting SetUnderline(bool underline, bool doubleUnderLine = false)
        {
            this.Underline = underline;
            this.DoubleUnderline = doubleUnderLine;
            return this;
        }
    }
}
