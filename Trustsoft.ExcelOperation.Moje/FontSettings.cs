using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    public class FontSettings
    {
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
        /// Get or set whether the text is strikethrough.
        /// </summary>
        public bool? TextCrossed { get; set; }

        /// <summary>
        /// Get or set whether the text is wrapping.
        /// </summary>
        public bool? TextWrapping { get; set; }

        /// <summary>
        /// Get or set font name.
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// Get or set font size.
        /// </summary>
        public double? FontSize { get; set; }

        /// <summary>
        /// Get or set the alpha component of the color (0-255).
        /// </summary>
        public int? A { get; set; }

        /// <summary>
        /// Get or set the red component of the color (0-255).
        /// </summary>
        public int? R { get; set; }

        /// <summary>
        /// Get or set the green component of the color (0-255).
        /// </summary>
        public int? G { get; set; }

        /// <summary>
        /// Get or set the blue component of the color (0-255).
        /// </summary>
        public int? B { get; set; }

        /// <summary>
        /// Sets the font to bold.
        /// </summary>
        /// <param name="bold">Indicates whether the font should be bold.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetBold(bool bold)
        {
            this.Bold = bold;
            return this;
        }

        /// <summary>
        /// Sets the font to italics.
        /// </summary>
        /// <param name="italics">Indicates whether the font should be italics.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetItalics(bool italics)
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
        public FontSettings SetUnderline(bool underline, bool doubleUnderLine = false)
        {
            this.Underline = underline;
            this.DoubleUnderline = doubleUnderLine;
            return this;
        }

        /// <summary>
        /// Sets the text is strikethrough.
        /// </summary>
        /// <param name="textCrossed">Indicates whether the text should be strikethrough.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetTextCrossed(bool textCrossed)
        {
            this.TextCrossed = textCrossed;
            return this;
        }

        /// <summary>
        /// Sets the text is wrapping.
        /// </summary>
        /// <param name="textWrapping">Indicates whether the text should be wrapping.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetTextWrapping(bool textWrapping)
        { 
            this.TextWrapping = textWrapping;
            return this;
        }

        /// <summary>
        /// Sets the font name.
        /// </summary>
        /// <param name="fontName">Font name.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetFontName(string fontName)
        {
            this.FontName = fontName;
            return this;
        }

        /// <summary>
        /// Sets the font size.
        /// </summary>
        /// <param name="fontSize">Font size.</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetFontSize(double fontSize)
        {
            this.FontSize = fontSize;
            return this;
        }

        /// <summary>
        /// Sets parameters ARGB.
        /// </summary>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        /// <returns>The current instance of the <see cref="FontSettings"/> class to allow for method chaining.</returns>
        public FontSettings SetTextColorARGB(int a, int r, int g, int b)
        {
            this.A = a;
            this.R = r;
            this.G = g;
            this.B = b;
            return this;
        }
    }
}
