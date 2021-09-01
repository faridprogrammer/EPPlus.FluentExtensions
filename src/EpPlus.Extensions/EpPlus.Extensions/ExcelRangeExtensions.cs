using System;
using System.Drawing;
using OfficeOpenXml.Style;

namespace OfficeOpenXml
{
    public static class ExcelRangeExtensions
    {
        /// <summary>
        /// Sets border around to excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color">Default is black</param>
        /// <param name="style">Default is thin</param>
        /// <returns></returns>
        public static ExcelRange SetBorder(this ExcelRange range, Color color = default, ExcelBorderStyle style = ExcelBorderStyle.Thin)
        {
            AssertNull(range);
            if (color == default)
                color = Color.Black;
            range.Style.Border.BorderAround(style, color);
            return range;
        }

        /// <summary>
        /// Sets custom text format to excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="format">String format</param>
        /// <returns></returns>
        public static ExcelRange SetFormat(this ExcelRange range, string format)
        {
            AssertNull(range);
            if (string.IsNullOrEmpty(format))
                return range;
            range.Style.Numberformat.Format = format;
            return range;
        }

        /// <summary>
        /// Sets value of excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ExcelRange SetValue(this ExcelRange range, object value)
        {
            AssertNull(range);
            range.Value = value;
            return range;
        }

        /// <summary>
        /// Sets excel range vertical align to Center
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetVerticalAlignCenter(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            return range;
        }

        /// <summary>
        /// Sets excel range vertical align to bottom
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetVerticalAlignBottom(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            return range;
        }

        /// <summary>
        /// Sets excel range vertical align to top
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetVerticalAlignTop(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            return range;
        }

        /// <summary>
        /// Sets excel range horizontal align to Center
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetHorizontalAlignCenter(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            return range;
        }

        /// <summary>
        /// Sets excel range horizontal align to right
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetHorizontalAlignRight(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            return range;
        }

        /// <summary>
        /// Sets excel range horizontal align to left
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetHorizontalAlignLeft(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            return range;
        }

        /// <summary>
        /// Fills the excel range with a color. Default color is Black and default style is Solid
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelRange SetFill(this ExcelRange range, Color color = default, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            AssertNull(range);
            if (color == default)
                color = Color.White;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
            return range;
        }

        /// <summary>
        /// Sets font of excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="size">em size of font</param>
        /// <param name="name">font name</param>
        /// <param name="color">font color</param>
        /// <param name="style">font style, Default is regular</param>
        /// <param name="bold">Default is false</param>
        /// <param name="italic">Default is false</param>
        /// <param name="underline">Default is false</param>
        /// <returns></returns>
        public static ExcelRange SetFont(this ExcelRange range, float size = 10F, string name = "Calibri", Color color = default, FontStyle style = FontStyle.Regular, bool bold = false, bool italic = false, bool underline = false)
        {
            AssertNull(range);
            if (color == default)
                color = Color.Black;

            range.Style.Font.Color.SetColor(color);
            range.Style.Font.Size = size;
            range.Style.Font.Name = name;
            range.Style.Font.Bold = bold;
            range.Style.Font.Italic = italic;
            range.Style.Font.UnderLine = underline;

            return range;
        }

        /// <summary>
        /// Makes excel range auto fit columns
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetAutoFitColumn(this ExcelRange range)
        {
            AssertNull(range);
            range.AutoFitColumns();
            return range;
        }

        /// <summary>
        /// Merges selected excel range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetMerge(this ExcelRange range)
        {
            AssertNull(range);
            range.Merge = true;
            return range;
        }

        /// <summary>
        /// Sets reading order of excel range to right
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetReadingOrderRight(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
            return range;
        }

        /// <summary>
        /// Sets reading order of excel range to left
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetReadingOrderLeft(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.ReadingOrder = ExcelReadingOrder.LeftToRight;
            return range;
        }

        /// <summary>
        /// Sets wrap text of excel range to true
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static ExcelRange SetWrapText(this ExcelRange range)
        {
            AssertNull(range);
            range.Style.WrapText = true;
            return range;
        }

        /// <summary>
        /// Sets formula to excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="formula"></param>
        /// <returns></returns>
        public static ExcelRange SetFormula(this ExcelRange range, string formula)
        {
            AssertNull(range);
            if (string.IsNullOrEmpty(formula))
                return range;
            range.Formula = formula;
            return range;
        }

        /// <summary>
        /// Sets comment on excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="comment"></param>
        /// <param name="author"></param>
        /// <returns></returns>
        public static ExcelRange SetComment(this ExcelRange range, string comment, string author = null)
        {
            AssertNull(range);
            if (string.IsNullOrEmpty(comment))
                return range;
            range.Comment.Text = comment;
            if(author != null)
                range.Comment.Author = author;

            return range;
        }

        /// <summary>
        /// Sets hyperlink to excel range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="link"></param>
        /// <returns></returns>
        public static ExcelRange SetLink(this ExcelRange range, string link)
        {
            AssertNull(range);
            if (string.IsNullOrEmpty(link))
                return range;
            range.Hyperlink = new Uri(link);
            return range;
        }

        private static void AssertNull(ExcelRange range)
        {
            if(range == null)
                throw new ArgumentNullException(nameof(range));
        }
    }
}
