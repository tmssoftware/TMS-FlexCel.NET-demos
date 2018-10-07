using FlexCel.Core;
using FlexCel.Render;
using FlexCel.XlsAdapter;
using System;

namespace GettingStarted
{
    class Program
    {
        static void Main(string[] args)
        {
            var xls = new XlsFile(1, TExcelFileFormat.v2019, true);
            xls.SetCellValue(1, 1, "Hello");
            xls.SetCellValue(2, 1, "World");
            xls.SetCellValue(3, 1, new TFormula("=A1 & \" \" & A2"));

            xls.Save("helloworld.xlsx");
            using (FlexCelPdfExport pdf = new FlexCelPdfExport(xls, true))
            {
                pdf.Export("helloworld.pdf");
            }

            Console.WriteLine("All done!");

        }
    }
}
