using System;
using System.Collections.Generic;
using System.IO;
using dBASE.NET;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

//select COM object Microsoft Excel xx.x Object Library before compiling!
namespace XtoXtool
{
    class Program
    {
        static void Main(string[] args)
        {
            string remotePath = @"\\helpdata-1\HR_OR\";
            Console.WriteLine("XML to XLS converter for PFU v.1.0 © 2023 Ihor Holovko");

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: xtoxtool <filename.xml>");
                ErrorExit("Nothing to convert.");
                return;
            }

            string filePath = Path.GetDirectoryName(args[0]);
            string xml;
            try
            {
                xml = File.ReadAllText(args[0]);
            }
            catch (Exception ex)
            {
                ErrorExit(ex.Message);
                return;
            }
            Console.WriteLine("Parsing XML ...");
            var table1 = xml.ParseXML<Table>();
            List<Part> dbfList = new List<Part>();

            string dBFfile = Path.Combine(remotePath, $"{DateTime.Now.Year}{DateTime.Now.Month:D2}{DateTime.Now.Day:D2}", "people.dbf");

            Dbf dbf = new Dbf();
            if (File.Exists(dBFfile))
            {
                dbf.Read(dBFfile);
                Console.Write("Parsing DBF ");
                int i = dbf.Records.Count;
                i /= 20;
                int j = 0;
                foreach (DbfRecord record in dbf.Records)
                {
                    dbfList.Add(new Part()
                    {
                        Tabn = Int32.Parse(record.Data[0]?.ToString() ?? "0"),
                        Icnum = record.Data[26]?.ToString() ?? "",
                        Ceh = record.Data[11]?.ToString() ?? "",
                    }
                    );
                    j++;
                    if (j > i)
                    {
                        Console.Write(".");
                        j = 0;
                    }

                }
            }
            else
            {
                ErrorExit($"DBF file not found. Please check {dBFfile}");
                return;
            }

            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                ErrorExit("Excel not found!");
                return;
            }
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Excel.Range aRange;
            object misValue = Type.Missing;

            Console.WriteLine();
            Console.Write("Compiling XLS ");

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 2] = $"{DateTime.Now:dd.MM.yyyy}";
            xlWorkSheet.Cells[1, 11] = "ПАТ \"Укртатнафта\"";
            aRange = xlWorkSheet.get_Range("A2", "N2");
            aRange.Merge();
            aRange.Font.Bold = true;
            aRange.Font.Size = 12;
            aRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[2, 1] = "Лікарняні листки тимчасової непрацездатності";

            short firstRow = 4;

            xlWorkSheet.Cells[firstRow, 1] = "Цех";
            xlWorkSheet.Cells[firstRow, 2] = "Табельний номер";
            xlWorkSheet.Cells[firstRow, 3] = "Прізвище";
            xlWorkSheet.Cells[firstRow, 4] = "Ім'я";
            xlWorkSheet.Cells[firstRow, 5] = "По-батькові";
            xlWorkSheet.Cells[firstRow, 6] = "Номер ЛН";
            xlWorkSheet.Cells[firstRow, 7] = "Дата відкриття";
            xlWorkSheet.Cells[firstRow, 8] = "Дата закриття";
            xlWorkSheet.Cells[firstRow, 9] = "ІПН";
            xlWorkSheet.Cells[firstRow, 10] = "Статус ЛН";
            xlWorkSheet.Cells[firstRow, 11] = "Порушення режиму";
            xlWorkSheet.Cells[firstRow, 12] = "Порушення алко/нарко";
            xlWorkSheet.Cells[firstRow, 13] = "Причина лікарняного";
            xlWorkSheet.Cells[firstRow, 14] = "Назва медзакладу";

            aRange = xlWorkSheet.get_Range("A" + firstRow, "N" + firstRow);
            aRange.Font.Bold = true;
            aRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            aRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            aRange.WrapText = true;
            aRange.AutoFilter(3);

            firstRow = 5;

            for (int i = 0; i < table1.Rows.Count(); i++)
            {
                int index = dbfList.FindIndex(t => t.Icnum == (table1.Rows[i].NP_NUMIDENT.ToString()));
                if (index != -1)
                {
                    xlWorkSheet.Cells[i + firstRow, 1] = dbfList[index].Ceh;
                    xlWorkSheet.Cells[i + firstRow, 2] = dbfList[index].Tabn;
                }

                xlWorkSheet.Cells[i + firstRow, 3] = table1.Rows[i].NP_SURNAME;
                xlWorkSheet.Cells[i + firstRow, 4] = table1.Rows[i].NP_NAME;
                xlWorkSheet.Cells[i + firstRow, 5] = table1.Rows[i].NP_PATRONYMIC;

                xlWorkSheet.Cells[i + firstRow, 6] = table1.Rows[i].WIC_NUM;
                xlWorkSheet.Cells[i + firstRow, 7] = table1.Rows[i].WIC_DT_BEGIN;
                xlWorkSheet.Cells[i + firstRow, 8] = table1.Rows[i].WIC_DT_END;
                xlWorkSheet.Cells[i + firstRow, 9] = table1.Rows[i].NP_NUMIDENT;

                xlWorkSheet.Cells[i + firstRow, 10] = table1.Rows[i].WIC_STATUS == "P" ? "Готовий до сплати" : "Закритий";
                xlWorkSheet.Cells[i + firstRow, 11] = table1.Rows[i].VIOLATION_EXTENSION == true ? "Є порушення" : "Немає";
                xlWorkSheet.Cells[i + firstRow, 12] = table1.Rows[i].SIGN_ANLK_NARKOTIK_INTOXICATION == true ? "Є порушення" : "Немає";

                xlWorkSheet.Cells[i + firstRow, 13] = table1.Rows[i].WIC_CD_Name;
                xlWorkSheet.Cells[i + firstRow, 14] = table1.Rows[i].HOSPITAL_NAME;
                Console.Write(".");
            }

            firstRow = 4;

            aRange = xlWorkSheet.get_Range("A" + (firstRow), "N" + (table1.Rows.Count() + firstRow));
            aRange.Borders.Color = Color.Black.ToArgb();
            aRange.Columns.AutoFit();

            xlWorkSheet.Cells[table1.Rows.Count() + 6, 8] = "Начальник ВК СУП _______________________ Маргарита ІЦЕНКО";

            dBFfile = Path.Combine(filePath, $"Лікарняні листки {DateTime.Now:ddMMyyyy}");

            if (File.Exists(dBFfile + ".xls"))
            {
                dBFfile += $"_{DateTime.Now:HH_mm_ss}.xls";
            }
            else dBFfile += ".xls";

            Console.WriteLine();
            Console.WriteLine($"Save to file {dBFfile}");

            try
            {
                xlWorkBook.SaveAs(dBFfile, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            }
            catch (Exception ex)
            {
                ErrorExit(ex.Message);
                return;
            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("Exiting.");

        }

        private static void ErrorExit(string msg)
        {
            Console.WriteLine(msg);
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}
