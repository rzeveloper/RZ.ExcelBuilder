using ClosedXML.Excel;

namespace RZ.ExcelBuilder.Core.Extensions
{
    internal static class ClosedXMLExtension
    {
        public static void AddHeaders(this IXLWorksheet worksheet, string[] headers)
        {
            for (int index = 0; index < headers.Length; index++)
            {
                var cell = worksheet.Cell(1, index + 1);
                cell.Value = headers[index];
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Fill.BackgroundColor = XLColor.BlueGray;
            }

            worksheet.Range(1, 1, 1, headers.Length)
                .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold(true);
        }

        public static void ToDateFormat(this IXLStyle style, string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                format = "yyyy-MM-dd";
            }

            style.DateFormat.SetFormat(format);
        }

        public static void ToMoneyFormat(this IXLStyle style, bool showDecimals)
        {
            if (showDecimals)
            {
                style.NumberFormat.Format = "#,##0.00";
            }
            else
            {
                style.NumberFormat.Format = "#,##0";
            }
        }

        public static void ToNumberFormat(this IXLStyle style, bool isDecimal)
        {
            if (isDecimal)
            {
                style.NumberFormat.NumberFormatId = 2;
            }
            else
            {
                style.NumberFormat.NumberFormatId = 1;
            }
        }
    }
}
