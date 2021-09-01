using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlus.FluentExtensions.Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Test]
        public void SimpleUsageWithoutAnyExtraUsing()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("SimpleUsage");
            sheet.Cells["A1:N1"].SetBorder().SetHorizontalAlignCenter();
        }
    }
}