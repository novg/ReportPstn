using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;

namespace ReportPstn
{
    class Reader
    {
        private static List<string> Parse(List<string> list)
        {
            List<string> parsedList = new List<string>();
            parsedList.Add(list.ElementAt(1).Split(' ')[0]);
            parsedList.Add(list.ElementAt(1).Split(' ')[1]);
            parsedList.Add(list.ElementAt(2));
            parsedList.Add(list.ElementAt(0).Replace("39042", ""));
            parsedList.Add(list.ElementAt(5));
            parsedList.Add(list.ElementAt(6).Replace(".", ","));
            parsedList.Add(list.ElementAt(4));
            parsedList.Add(list.ElementAt(4));
            parsedList.Add("");
            parsedList.Add("");

            return parsedList;
        }

        internal List<List<string>> ReadFile(string fileName, out int separator)
        {
            separator = 0;
            //try
            //{
                using (FileStream fs = File.OpenRead(fileName))
                {
                    IWorkbook wb = WorkbookFactory.Create(fs);
                    ISheet sheet = wb.GetSheetAt(0);
                    Regex pattern = new Regex(@"\W+");
                    bool rtk = false;
                    int correction = 0;
                    List<List<string>> table = new List<List<string>>();

                    for (int row = 0; row < sheet.LastRowNum; row++)
                    {
                        List<string> list = new List<string>();
                        IRow currentRow = sheet.GetRow(row);
                        for (int col = 0; col < currentRow.LastCellNum; col++)
                        {
                            ICell cell = currentRow.GetCell(col);
                            cell.SetCellType(CellType.String);
                            if (!rtk && cell.StringCellValue.Contains("Всего по поставщику"))
                            {
                                rtk = true;
                                separator = row;
                            }

                            MatchCollection matches = pattern.Matches(cell.StringCellValue);
                            if (col == 0 && matches.Count > 0)
                            {
                                if (!rtk) correction++;
                                break;
                            }

                            list.Add(cell.StringCellValue);
                        }

                        if (list.Count > 0)
                            table.Add(Parse(list));
                    }

                    separator -= correction;
                    return table;
                }
            //}
            //catch (ArgumentException ex)
            //{
            //    throw new ArgumentException(string.Format("Неподдерживаемый формат файла:\n{0}", ex.Message));
            //}
            //catch (InvalidOperationException ex)
            //{
            //    throw new InvalidOperationException(string.Format("Неподдерживаемая структура файла:\n{0}", ex.Message));
            //}

        }
    }
}
