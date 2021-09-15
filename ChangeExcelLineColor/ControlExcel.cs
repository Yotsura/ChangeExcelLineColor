using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;

namespace ChangeExcelLineColor
{
    public static class ControlExcel
    {
        public static void ChangeAllBooks(string filename)
        {
            var fileInfo = new FileInfo(filename);
            var newFileName = $"{fileInfo.Directory}\\FIXED_{fileInfo.Name}";
            File.Copy(filename, newFileName, true);
            using (var book = new XLWorkbook(newFileName))
            {
                var theme = book.Theme;
                foreach (var sheet in book.Worksheets)
                {
                    var printArea = sheet.PageSetup.PrintAreas.FirstOrDefault();
                    if (printArea == null) continue;
                    ChangeRanges(printArea, theme);
                    //foreach (var printArea in sheet.PageSetup.PrintAreas)
                    //{
                    //    if (printArea == null) continue;
                    //    ChangeRanges(printArea, theme);
                    //}
                }
                book.SaveAs(newFileName);
            }
            Console.WriteLine($"修正済みファイル：{new FileInfo(newFileName).Name}");
        }

        private static void ChangeRanges(IXLRange range, IXLTheme theme)
        {
            foreach(var targetCell in range.Cells())
            {
                ChangeCell(targetCell, theme);
            }
        }

        private static void ChangeCell(IXLCell targetCell, IXLTheme theme)
        {
            if (targetCell.Style.Border.BottomBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.BottomBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.BottomBorderColor.IsNonColored(theme))
                targetCell.Style.Border.BottomBorderColor = XLColor.Black;
            if (targetCell.Style.Border.TopBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.TopBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.TopBorderColor.IsNonColored(theme))
                targetCell.Style.Border.TopBorderColor = XLColor.Black;
            if (targetCell.Style.Border.LeftBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.LeftBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.LeftBorderColor.IsNonColored(theme))
                targetCell.Style.Border.LeftBorderColor = XLColor.Black;
            if (targetCell.Style.Border.RightBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.RightBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.RightBorderColor.IsNonColored(theme))
                targetCell.Style.Border.RightBorderColor = XLColor.Black;

            if (targetCell.Style.Border.DiagonalBorder != XLBorderStyleValues.None &&
                targetCell.Style.Border.DiagonalBorderColor.IsNonColored(theme))
                targetCell.Style.Border.DiagonalBorderColor = XLColor.Black;
        }

        private static bool IsNonColored(this XLColor xlColor, IXLTheme theme)
        {
            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    return xlColor.ToString() == "00000000";
                case XLColorType.Theme:
                    var xlThemeColor = theme.ResolveThemeColor(xlColor.ThemeColor).ToString();
                    var black = XLColor.Black.ToString();
                    return xlThemeColor == black;
                case XLColorType.Indexed:
                    var color = XLColor.IndexedColors[xlColor.Indexed];
                    if (xlColor.Indexed >= 64) return true;
                    break;
            }
            return false;
        }
    }
}
