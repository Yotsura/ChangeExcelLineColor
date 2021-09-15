using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChangeExcelLineColor
{
    class Program
    {
        static void Main(string[] args)
        {
            //ControlExcel.ChangeAllBooks(@"C:\Users\pc-yuminaka\Desktop\TestBook.xlsx");
            var files = System.Environment.GetCommandLineArgs();
            var filenames = new List<string>();
            if (files.Length > 1)
            {
                filenames = files.Skip(1).ToList();
            }
            else
            {
                Console.WriteLine(
                    "「ChangeExcelLineColor.exe」に直接EXCELファイルをドラックアンドドロップしてください。※複数可能\r\n" +
                    "＜仕様＞\r\n印刷範囲内のすべてのセルについて以下の変更を行います。\r\n" +
                    "線のない辺　→　白の線を設定　※斜め線は対象外\r\n" +
                    "色自動の線あり　→　線の色を黒に変更\r\n" +
                    "色が自動・黒以外の線　→　そのまま\r\n\r\n※この画面はなにかキーを押せば閉じます。");
                Console.ReadKey();
                return;
            }
            var xlsxfiles = filenames.Where(file => file.EndsWith("xlsx") || file.EndsWith("xlsm") || file.EndsWith("xls")).ToList();
            if (xlsxfiles.Count() < 1)
            {
                Console.WriteLine("EXCELファイルがありません。");
                Console.ReadKey();
                return;
            }

            xlsxfiles.ForEach(file => ControlExcel.ChangeAllBooks(file));
            Console.WriteLine("作業完了");
            Console.ReadKey();
        }
    }
}
