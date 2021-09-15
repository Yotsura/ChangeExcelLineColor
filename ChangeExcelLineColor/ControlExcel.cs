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
            //File.Copy(filename, newFileName, true);
            using (var book = new XLWorkbook(filename))
            {
                var theme = book.Theme;
                foreach (var sheet in book.Worksheets)
                {
                    var printArea = sheet.PageSetup.PrintAreas.FirstOrDefault();
                    if (printArea == null) continue;
                    ChangeRanges(printArea, theme);
                }
                book.SaveAs(newFileName);
            }
            Console.WriteLine($"修正済みファイル：{new FileInfo(newFileName).Name}");
        }

        private static void ChangeRanges(IXLRange range, IXLTheme theme)
        {
            foreach (var targetCell in range.Cells())
            {
                ChangeCell(targetCell, theme);
            }
        }

        private static void ChangeCell(IXLCell targetCell, IXLTheme theme)
        {
            var setColor = XLColor.Black;
            if (targetCell.Style.Font.FontColor.IsNonColored(theme))
                targetCell.RichText.SetFontColor(setColor);
            if (targetCell.Style.Border.BottomBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.BottomBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.BottomBorderColor.IsNonColored(theme))
                targetCell.Style.Border.BottomBorderColor = setColor;
            if (targetCell.Style.Border.TopBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.TopBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.TopBorderColor.IsNonColored(theme))
                targetCell.Style.Border.TopBorderColor = setColor;
            if (targetCell.Style.Border.LeftBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.LeftBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.LeftBorderColor.IsNonColored(theme))
                targetCell.Style.Border.LeftBorderColor = setColor;
            if (targetCell.Style.Border.RightBorder == XLBorderStyleValues.None)
            {
                targetCell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                targetCell.Style.Border.RightBorderColor = XLColor.White;
            }
            else if (targetCell.Style.Border.RightBorderColor.IsNonColored(theme))
                targetCell.Style.Border.RightBorderColor = setColor;

            if (targetCell.Style.Border.DiagonalBorder != XLBorderStyleValues.None &&
                targetCell.Style.Border.DiagonalBorderColor.IsNonColored(theme))
                targetCell.Style.Border.DiagonalBorderColor = setColor;
        }

        private static bool IsNonColored(this XLColor xlColor, IXLTheme theme)
        {
            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    return xlColor.ToString().IsBlack();
                case XLColorType.Theme:
                    var xlThemeColor = theme.ResolveThemeColor(xlColor.ThemeColor).ToString();
                    return xlThemeColor.IsBlack();
                case XLColorType.Indexed:
                    var color = XLColor.IndexedColors[xlColor.Indexed];
                    return xlColor.Indexed >= 64;
            }
            return false;
        }

        private static bool IsBlack(this string color)
        {
            var black = new List<string>
            {
                XLColor.FromArgb(255, 255, 255).ToString(),
                XLColor.Black.ToString(),
                XLColor.FromArgb(00000000).ToString(),
                XLColor.FromIndex(1).ToString(),
                XLColor.FromName("Black").ToString()
            };

            return black.Contains(color);

        }
    }
}