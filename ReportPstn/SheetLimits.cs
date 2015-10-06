using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ReportPstn
{
    class SheetLimits
    {
        IWorkbook workbook;
        ISheet sheet;
        string date;
        List<int> summedCells;
        int totalRow;
        CellStyles cellStyles;
        ICellStyle styleInnerHeading;
        ICellStyle styleHeading;
        ICellStyle styleHeadingSum;
        ICellStyle styleHeadline;
        ICellStyle styleTable;
        ICellStyle styleDouble;
        ICellStyle styleDoubleBold;
        ICellStyle styleCorpColor;
        ICellStyle styleLimitsDoubleSum;
        ICellStyle styleResultHeadline;

        public SheetLimits(IWorkbook workbook, string date)
        {
            this.workbook = workbook;
            this.date = date;
            sheet = workbook.CreateSheet(date + "-Лимиты");
            summedCells = new List<int>();

            cellStyles = new CellStyles(workbook);
            styleInnerHeading = cellStyles.InnerHeadline();
            styleHeading = cellStyles.Heading();
            styleHeadingSum = cellStyles.HeadingSum();
            styleHeadline = cellStyles.Headline();
            styleTable = cellStyles.LimitsTable();
            styleDouble = cellStyles.LimitsDouble();
            styleDoubleBold = cellStyles.LimitsDoubleBold();
            styleCorpColor = cellStyles.CorpColor();
            styleLimitsDoubleSum = cellStyles.LimitsDoubleSum();
            styleResultHeadline = cellStyles.ResultHeadline();
        }

        internal void Create()
        {
            IRow rowHeadline = sheet.CreateRow(0);
            ICell cellHeadLine = rowHeadline.CreateCell(0);
            cellHeadLine.SetCellValue("Затраты на междугородние переговоры за " + date);
            cellHeadLine.CellStyle = styleHeadline;
            int row = 2;
            int rowMerge = row + 3;
            string[,] table = LimitsInitializer.ManagementSaz();
            CreateTableLimits("РУКОВОДСТВО  ОАО \"Саяногорский Алюминиевый завод\"", table, ref row, false);
            MergeCells(rowMerge, row - 1, 0);

            rowMerge = row + 3;
            table = LimitsInitializer.Otb();
            CreateTableLimits("ОТДЕЛ ОХРАНЫ ТРУДА И ПРОМЫШЛЕННОЙ БЕЗОПАСНОСТИ", table, ref row);
            MergeCells(rowMerge, row - 1, 200);

            rowMerge = row + 3;
            table = LimitsInitializer.Pdo();
            CreateTableLimits("ПРОИЗВОДСТВЕННО-ДИСПЕТЧЕРСКИЙ ОТДЕЛ", table, ref row);
            MergeCells(rowMerge, row - 1, 450);

            rowMerge = row + 3;
            table = LimitsInitializer.Lawyers();
            CreateTableLimits("ЮРИДИЧЕСКИЙ ОТДЕЛ", table, ref row);
            MergeCells(rowMerge, row - 1, 600);

            rowMerge = row + 3;
            table = LimitsInitializer.Electrolyse();
            CreateTableLimits("ДИРЕКЦИЯ ПО ЭЛЕКТРОЛИЗНОМУ ПРОИЗВОДСТВУ", table, ref row);
            MergeCells(rowMerge, rowMerge + 1, 250);
            rowMerge += 2;
            MergeCells(rowMerge, rowMerge, 250);

            rowMerge = row + 3;
            table = LimitsInitializer.Foundry();
            CreateTableLimits("ДИРЕКЦИЯ ПО ЛИТЕЙНОМУ ПРОИЗВОДСТВУ", table, ref row);
            MergeCells(rowMerge, row - 1, 500);

            rowMerge = row + 3;
            table = LimitsInitializer.Electrode();
            CreateTableLimits("ДИРЕКЦИЯ ПО ПРОИЗВОДСТВУ ЭЛЕКТРОДОВ", table, ref row);
            MergeCells(rowMerge, row - 1, 500);

            rowMerge = row + 3;
            table = LimitsInitializer.Energy();
            CreateTableLimits("СЛУЖБА ГЛАВНОГО ЭНЕРГЕТИКА", table, ref row);
            MergeCells(rowMerge, row - 1, 600);

            rowMerge = row + 3;
            table = LimitsInitializer.Ecology();
            CreateTableLimits("ДИРЕКЦИЯ ПО ЭКОЛОГИИ И КАЧЕСТВУ", table, ref row);
            MergeCells(rowMerge, row - 1, 200);

            rowMerge = row + 3;
            table = LimitsInitializer.Commerce();
            CreateTableLimits("КОММЕРЧЕСКАЯ ДИРЕКЦИЯ", table, ref row);
            MergeCells(rowMerge, rowMerge + 3, 1000);
            rowMerge += 4;
            MergeCells(rowMerge, rowMerge + 1, 500);
            rowMerge += 2;
            MergeCells(rowMerge, rowMerge + 11, 2850);
            rowMerge += 12;
            MergeCells(rowMerge, rowMerge + 3, 400);
            rowMerge += 4;
            MergeCells(rowMerge, rowMerge + 1, 50);

            rowMerge = row + 3;
            table = LimitsInitializer.Personnel();
            CreateTableLimits("ДИРЕКЦИЯ ПО ПЕРСОНАЛУ", table, ref row);
            MergeCells(rowMerge, rowMerge + 7, 850);
            rowMerge += 9;
            MergeCells(rowMerge, rowMerge, 400);

            rowMerge = row + 3;
            table = LimitsInitializer.Finance();
            CreateTableLimits("ФИНАНСОВАЯ ДИРЕКЦИЯ", table, ref row);
            MergeCells(rowMerge, row - 1, 1000);

            rowMerge = row + 3;
            table = LimitsInitializer.Security();
            CreateTableLimits("ДИРЕКЦИЯ ПО ЗАЩИТЕ РЕСУРСОВ", table, ref row);
            MergeCells(rowMerge, row - 1, 1000);

            rowMerge = row + 3;
            table = LimitsInitializer.TradeUnion();
            CreateTableLimits("ПРОФКОМ  ЗАВОДА", table, ref row);
            MergeCells(rowMerge, row - 1, 500);

            rowMerge = row + 3;
            table = LimitsInitializer.PressService();
            CreateTableLimits("ПРЕСС-СЛУЖБА", table, ref row);
            MergeCells(rowMerge, row - 1, 100);

            rowMerge = row + 3;
            table = LimitsInitializer.VeteranUnion();
            CreateTableLimits("Союз ветеранов", table, ref row);
            MergeCells(rowMerge, row - 1, 100);

            ResultTable(ref row);

            rowMerge = row + 3;
            table = LimitsInitializer.Dis();
            CreateTableLimits("ДИС", table, ref row, false);
            MergeCells(rowMerge, row - 1, 0);

            rowMerge = row + 3;
            table = LimitsInitializer.Other();
            CreateTableLimits("ПРОЧИЕ", table, ref row, false);
            MergeCells(rowMerge, row - 1, 0);

            TotalTable(ref row);

            sheet.SetColumnWidth(0, 256 * 80);
            sheet.SetColumnWidth(1, 256 * 12);
            sheet.SetColumnWidth(2, 256 * 10);
            sheet.SetColumnWidth(3, 256 * 10);
            sheet.SetColumnWidth(4, 256 * 10);
            sheet.SetColumnWidth(5, 256 * 10);
            sheet.SetColumnWidth(6, 256 * 10);
            sheet.SetColumnWidth(7, 256 * 10);
        }

        private void CreateTableLimits(string name, string[,] table, ref int row, bool limit = true)
        {
            row++;
            IRow currentRow = sheet.CreateRow(row);
            ICell currentCell = currentRow.CreateCell(0);
            currentCell.SetCellValue(name);
            currentCell.CellStyle = styleHeadline;
            FillTable(table, ref row, limit);
        }

        private void MergeCells(int rowBegin, int rowEnd, int limit)
        {
            sheet.AddMergedRegion(new CellRangeAddress(rowBegin, rowEnd, 4, 4));
            sheet.AddMergedRegion(new CellRangeAddress(rowBegin, rowEnd, 5, 5));
            sheet.AddMergedRegion(new CellRangeAddress(rowBegin, rowEnd, 6, 6));

            IRow currentRow = sheet.GetRow(rowBegin);
            ICell cell = currentRow.CreateCell(4);

            if (limit > 0)
            {
                cell.SetCellValue(limit);
                cell.CellStyle = styleHeading;
                cell = currentRow.CreateCell(5);
                cell.CellFormula = string.Format("IF(SUM(D{0}:D{1})>E{0},SUM(D{0}:D{1})-E{0},\"\")", rowBegin + 1, rowEnd + 1);
                cell.CellStyle = styleDoubleBold;
                cell = currentRow.CreateCell(6);
                cell.CellFormula = string.Format("IF(SUM(D{0}:D{1})<E{0},E{0}-SUM(D{0}:D{1}),\"\")", rowBegin + 1, rowEnd + 1);
                cell.CellStyle = styleDoubleBold;
            }
            else
            {
                cell.SetCellValue("");
                cell.CellStyle = styleHeading;
                cell = currentRow.CreateCell(5);
                cell.SetCellValue("");
                cell.CellStyle = styleDoubleBold;
                cell = currentRow.CreateCell(6);
                cell.SetCellValue("");
                cell.CellStyle = styleDoubleBold;
            }

            for (int i = rowBegin + 1; i <= rowEnd; i++)
            {
                currentRow = sheet.GetRow(i);
                for (int j = 4; j < 7; j++)
                {
                    cell = currentRow.CreateCell(j);
                    cell.CellStyle = styleHeading;
                }
            }
        }

        private void TotalTable(ref int row)
        {
            row++;
            IRow currentRow = sheet.CreateRow(++row);
            ICell currentCell = currentRow.CreateCell(5);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = string.Format("РТК!F{0}+ГТС!F{1}",
                Report.listSumRtkGts.ElementAt(1), Report.listSumRtkGts.ElementAt(0));

            currentRow = sheet.CreateRow(++row);
            currentCell = currentRow.CreateCell(5);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = string.Format("D{0}+D{1}+D{2}",
                summedCells.ElementAt(summedCells.Count - 2), summedCells.ElementAt(summedCells.Count - 1), totalRow);

            currentRow = sheet.CreateRow(++row);
            currentCell = currentRow.CreateCell(5);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = string.Format("F{0}-F{1}", row - 1, row);
        }

        private void ResultTable(ref int row)
        {
            row += 2;
            IRow headingRow = sheet.CreateRow(row);
            string[] caption = LimitsInitializer.Caption();
            for (int col = 1; col < caption.Length; col++)
            {
                ICell cell = headingRow.CreateCell(col + 2);
                cell.SetCellValue(caption[col]);
                cell.CellStyle = styleHeading;
            }

            IRow currentRow = sheet.CreateRow(++row);
            ICell currentCell = currentRow.CreateCell(2);
            currentCell.CellStyle = styleResultHeadline;
            currentCell.SetCellValue("ИТОГО ПО ЛИМИТИРУЕМЫМ НОМЕРАМ:");

            StringBuilder sumsBuilder = new StringBuilder();
            StringBuilder limitsBuilder = new StringBuilder();
            for (int i = 1; i < summedCells.Count; i++)
            {
                sumsBuilder.Append("D");
                sumsBuilder.Append(summedCells.ElementAt(i));
                sumsBuilder.Append("+");
                limitsBuilder.Append("E");
                limitsBuilder.Append(summedCells.ElementAt(i));
                limitsBuilder.Append("+");
            }

            sumsBuilder.Remove(sumsBuilder.Length - 1, 1);
            currentCell = currentRow.CreateCell(3);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = sumsBuilder.ToString();
            limitsBuilder.Remove(limitsBuilder.Length - 1, 1);
            currentCell = currentRow.CreateCell(4);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = limitsBuilder.ToString();
            currentCell = currentRow.CreateCell(5);
            currentCell.CellFormula = string.Format("IF(D{0}>E{0},D{0}-E{0},\"\")", row + 1);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell = currentRow.CreateCell(6);
            currentCell.CellFormula = string.Format("IF(D{0}<E{0},E{0}-D{0},\"\")", row + 1);
            currentCell.CellStyle = styleLimitsDoubleSum;

            currentRow = sheet.CreateRow(++row);
            currentCell = currentRow.CreateCell(2);
            currentCell.CellStyle = styleResultHeadline;
            currentCell.SetCellValue("ИТОГО ПО ЗАВОДУ:");
            currentCell = currentRow.CreateCell(3);
            currentCell.CellStyle = styleLimitsDoubleSum;
            currentCell.CellFormula = string.Format("D{0}+D{1}", summedCells.ElementAt(0), row);
            totalRow = row + 1;
        }

        private void FillTable(string[,] table, ref int row, bool limit)
        {
            row++;
            IRow headingRow = sheet.CreateRow(row);
            string[] caption = LimitsInitializer.Caption();
            for (int col = 0; col < caption.Length; col++)
            {
                ICell cell = headingRow.CreateCell(col + 2);
                cell.SetCellValue(caption[col]);
                cell.CellStyle = styleHeading;
            }

            row++;
            int rowStart = row;
            for (int line = 0; line < table.GetLength(0); line++)
            {
                IRow currentRow = sheet.CreateRow(row);
                ICell cell = currentRow.CreateCell(0);
                cell.SetCellValue(table[line, 0]);
                if (table[line, 2] == "")
                {
                    cell.CellStyle = styleInnerHeading;
                }
                else
                {
                    cell.CellStyle = styleTable;
                    cell = currentRow.CreateCell(1);
                    cell.CellStyle = styleTable;
                    if (table[line, 1] != "")
                        cell.SetCellValue(double.Parse(table[line, 1]));

                    cell = currentRow.CreateCell(2);
                    cell.CellStyle = styleTable;
                    cell.SetCellValue(double.Parse(table[line, 2]));

                    cell = currentRow.CreateCell(3);
                    cell.CellFormula = string.Format("SUMIF('{0}'!D:F,C{1},'{0}'!F:F)", "3счет", row + 1);
                    cell.CellStyle = styleDouble;

                    if (Initializer.GetCorporateDictionary().Keys.Contains(table[line, 2]))
                    {
                        cell = currentRow.CreateCell(7);
                        cell.SetCellValue("корп");
                        cell.CellStyle = styleCorpColor;
                    }
                }

                row++;
            }

            summedCells.Add(row + 1);
            headingRow = sheet.CreateRow(row);
            ICell cellSum = headingRow.CreateCell(2);
            cellSum.SetCellValue("ВСЕГО:");
            cellSum.CellStyle = styleLimitsDoubleSum;
            cellSum = headingRow.CreateCell(3);
            cellSum.CellFormula = string.Format("SUM(D{0}:D{1})", rowStart + 1, row);
            cellSum.CellStyle = styleLimitsDoubleSum;
            

            if (limit)
            {
                cellSum = headingRow.CreateCell(4);
                cellSum.CellFormula = string.Format("SUM(E{0}:E{1})", rowStart + 1, row);
                cellSum.CellStyle = styleLimitsDoubleSum;
                cellSum = headingRow.CreateCell(5);
                cellSum.CellFormula = string.Format("IF(D{0}>E{0},D{0}-E{0},\"\")", row + 1);
                cellSum.CellStyle = styleLimitsDoubleSum;
                cellSum = headingRow.CreateCell(6);
                cellSum.CellFormula = string.Format("IF(D{0}<E{0},E{0}-D{0},\"\")", row + 1);
                cellSum.CellStyle = styleLimitsDoubleSum;
            }
            else
            {
                cellSum = headingRow.CreateCell(4);
                cellSum.CellStyle = styleLimitsDoubleSum;
                cellSum = headingRow.CreateCell(5);
                cellSum.CellStyle = styleLimitsDoubleSum;
                cellSum = headingRow.CreateCell(6);
                cellSum.CellStyle = styleLimitsDoubleSum;
            }
        }
    }
}
