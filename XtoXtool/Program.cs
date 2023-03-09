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

namespace XtoXtool
{
    class Program
    {
        static void Main(string[] args)
        {
            string _remotePath = @"\\helpdata-1\HR_OR\";
            Console.WriteLine("XML to XLS converter for PFU v.1.0 © 2023 Ihor Holovko");

            if (args.Length == 0)
            {
                Console.WriteLine("usage: xtoxtool <filename.xml>");
                ErrorExit("Nothing to convert.");
                return;
            }
            else 
            {
                string _filePath = Path.GetDirectoryName(args[0]);
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
                List<Part> _dbf = new List<Part>();

                string _file = String.Concat(_remotePath, DateTime.Now.Year, DateTime.Now.Month.ToString("D2"), DateTime.Now.Day.ToString("D2"), Path.DirectorySeparatorChar, "people.dbf");

                Dbf dbf = new Dbf();
                if (File.Exists(_file))
                {
                    dbf.Read(_file);
                    Console.Write("Parsing DBF ");
                    int i = dbf.Records.Count;
                    i /= 20;
                    int j = 0;
                    foreach (DbfRecord record in dbf.Records)
                    {
                        _dbf.Add(new Part()
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
                    ErrorExit($"DBF file not found. Please check {_file}");
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

                short _firstColumn = 4;

                xlWorkSheet.Cells[_firstColumn, 1] = "Цех";
                xlWorkSheet.Cells[_firstColumn, 2] = "Табельний номер";
                xlWorkSheet.Cells[_firstColumn, 3] = "Прізвище";
                xlWorkSheet.Cells[_firstColumn, 4] = "Ім'я";
                xlWorkSheet.Cells[_firstColumn, 5] = "По-батькові";
                xlWorkSheet.Cells[_firstColumn, 6] = "Номер ЛН";
                xlWorkSheet.Cells[_firstColumn, 7] = "Дата відкриття";
                xlWorkSheet.Cells[_firstColumn, 8] = "Дата закриття";
                xlWorkSheet.Cells[_firstColumn, 9] = "ІПН";
                xlWorkSheet.Cells[_firstColumn, 10] = "Статус ЛН";
                xlWorkSheet.Cells[_firstColumn, 11] = "Порушення режиму";
                xlWorkSheet.Cells[_firstColumn, 12] = "Порушення алко/нарко";
                xlWorkSheet.Cells[_firstColumn, 13] = "Причина лікарняного";
                xlWorkSheet.Cells[_firstColumn, 14] = "Назва медзакладу";

                aRange = xlWorkSheet.get_Range("A" + _firstColumn, "N" + _firstColumn);
                aRange.Font.Bold = true;
                aRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                aRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                aRange.WrapText = true;

                _firstColumn = 5;

                for (int i = 0; i < table1.Rows.Count(); i++)
                {
                    int index = _dbf.FindIndex(t => t.Icnum == (table1.Rows[i].NP_NUMIDENT.ToString()));
                    if (index != -1)
                    {
                        xlWorkSheet.Cells[i + _firstColumn, 1] = _dbf[index].Ceh;
                        xlWorkSheet.Cells[i + _firstColumn, 2] = _dbf[index].Tabn;
                    }

                    xlWorkSheet.Cells[i + _firstColumn, 3] = table1.Rows[i].NP_SURNAME;
                    xlWorkSheet.Cells[i + _firstColumn, 4] = table1.Rows[i].NP_NAME;
                    xlWorkSheet.Cells[i + _firstColumn, 5] = table1.Rows[i].NP_PATRONYMIC;

                    xlWorkSheet.Cells[i + _firstColumn, 6] = table1.Rows[i].WIC_NUM;
                    xlWorkSheet.Cells[i + _firstColumn, 7] = table1.Rows[i].WIC_DT_BEGIN;
                    xlWorkSheet.Cells[i + _firstColumn, 8] = table1.Rows[i].WIC_DT_END;
                    xlWorkSheet.Cells[i + _firstColumn, 9] = table1.Rows[i].NP_NUMIDENT;

                    xlWorkSheet.Cells[i + _firstColumn, 10] = table1.Rows[i].WIC_STATUS == "P" ? "Готовий до сплати" : "Закритий";
                    xlWorkSheet.Cells[i + _firstColumn, 11] = table1.Rows[i].VIOLATION_EXTENSION == true ? "Є порушення" : "Немає";
                    xlWorkSheet.Cells[i + _firstColumn, 12] = table1.Rows[i].SIGN_ANLK_NARKOTIK_INTOXICATION == true ? "Є порушення" : "Немає";

                    xlWorkSheet.Cells[i + _firstColumn, 13] = table1.Rows[i].WIC_CD_Name;
                    xlWorkSheet.Cells[i + _firstColumn, 14] = table1.Rows[i].HOSPITAL_NAME;
                    Console.Write(".");
                }

                _firstColumn = 4;

                xlWorkSheet.get_Range("A2", "N2").Font.Bold = true;
                //aRange.Font.Bold = true;
                aRange = xlWorkSheet.get_Range("A" + (_firstColumn), "N" + (table1.Rows.Count() + _firstColumn));
                aRange.Borders.Color = Color.Black.ToArgb();
                aRange.Columns.AutoFit();

                aRange = xlWorkSheet.get_Range("A4", "N4");
                aRange.AutoFilter(3);

                xlWorkSheet.Cells[table1.Rows.Count() + 6, 8] = "Начальник ВК СУП _______________________ Маргарита ІЦЕНКО";

                _file = $"{_filePath}{Path.DirectorySeparatorChar}Лікарняні листки {DateTime.Now:ddMMyyyy}";

                if (File.Exists(_file + ".xls"))
                {
                    Random _rnd = new Random(DateTime.Now.Second);
                    _file += $"_{_rnd.Next(1, 10)}.xls";
                }
                else _file += ".xls";
                
                Console.WriteLine();
                Console.WriteLine($"Save to file {_file}");

                try
                {
                    xlWorkBook.SaveAs(_file, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

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
        }

        private static void ErrorExit(string msg)
        {
            Console.WriteLine(msg);
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}
