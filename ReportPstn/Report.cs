using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ReportPstn
{
    public class Report
    {
        public event EventHandler<EventArgs> Start;
        public event EventHandler<ReportEventArgs> Complete;
        internal static List<int> listSumRtkGts;
        enum SheetName
        {
            Corporate, MvzOrder, MvzPhone
        }

        enum RegionType
        {
            Rtk, Gts, All
        }

        List<List<string>> table;
        int separator;
        IWorkbook workbook;
        string fileName;
        CellStyles cellStyles;
        ICellStyle styleDouble;
        ICellStyle styleDoubleBold;
        ICellStyle styleTable;
        ICellStyle styleCaption;
        ICellStyle styleDoubleGts;
        ICellStyle styleGts;
        ICellStyle styleDoubleRtk;
        ICellStyle styleRtk;

        public Report(string fileName)
        {
            this.fileName = fileName;
            workbook = new XSSFWorkbook();
            listSumRtkGts = new List<int>();
            cellStyles = new CellStyles(workbook);
            styleDouble = cellStyles.Double();
            styleDoubleBold = cellStyles.DoubleBold();
            styleTable = cellStyles.Table();
            styleCaption = cellStyles.Caption();
            styleDoubleGts = cellStyles.DoubleGts();
            styleGts = cellStyles.Gts();
            styleDoubleRtk = cellStyles.DoubleRtk();
            styleRtk = cellStyles.Rtk();
        }

        public void Create()
        {
            Reader reader = new Reader();
            table = reader.ReadFile(fileName, out separator);

            if (Start != null)
            {
                Start(this, new EventArgs());
            }

            CreateSheetBasic(SheetName.Corporate);
            CreateSheetBasic(SheetName.MvzOrder);
            CreateSheetBasic(SheetName.MvzPhone);
            CreateSheetStatistic(RegionType.All);
            CreateSheetStatistic(RegionType.Gts);
            CreateSheetStatistic(RegionType.Rtk);
            CreateSheetMvz(RegionType.Gts);
            CreateSheetMvz(RegionType.Rtk);
            SheetLimits sheetLimits = new SheetLimits(workbook, dateString());
            sheetLimits.Create();

            ReportEventArgs reportEventArgs = new ReportEventArgs();
            if (Complete != null)
            {
                reportEventArgs.DefaultFilePath = GetOutFilePath();
                Complete(this, reportEventArgs);
            }
        }

        private void CreateSheetMvz(RegionType regionType)
        {
            ISheet sheet = null;
            int[] mvz = null;
            string rtype = string.Empty;

            switch (regionType)
            {
                case RegionType.Rtk:
                    sheet = workbook.CreateSheet("МВЗ РТК");
                    mvz = Initializer.GetMvzRtk();
                    rtype = "РТК";
                    break;
                case RegionType.Gts:
                    sheet = workbook.CreateSheet("МВЗ ГТС");
                    mvz = Initializer.GetMvzGts();
                    rtype = "ГТС";
                    break;
            }


            string[] caption = Initializer.GetMvzCaption();
            CreateCaption(caption, sheet, styleTable);

            for (int row = 0; row < mvz.Length; row++)
            {
                IRow currentRow = sheet.CreateRow(row + 1);
                ICell cell = currentRow.CreateCell(0);
                cell.SetCellValue(mvz[row]);
                cell.CellStyle = styleTable;

                cell = currentRow.CreateCell(1);
                cell.CellFormula = string.Format("VLOOKUP(A{0},'МВЗ=ЗАКАЗ'!A:B,2,(FALSE))", row + 2);
                cell.CellStyle = styleTable;

                cell = currentRow.CreateCell(2);
                cell.CellFormula = string.Format("SUMIF({0}!I:I,A{1},{0}!F:F)", rtype, row + 2);
                cell.CellStyle = styleDouble;
            }


            IRow rowSum = sheet.CreateRow(mvz.Length + 1);
            ICell cellSum = rowSum.CreateCell(0);
            cellSum.SetCellValue("Итого:");
            cellSum.CellStyle = styleCaption;

            cellSum = rowSum.CreateCell(1);
            cellSum.CellStyle = styleCaption;

            cellSum = rowSum.CreateCell(2);
            cellSum.CellFormula = string.Format("SUM(C2:C{0})", mvz.Length + 1);
            cellSum.CellStyle = styleDoubleBold;

            sheet.SetColumnWidth(0, 256 * 12);
            sheet.SetColumnWidth(1, 256 * 12);
            sheet.SetColumnWidth(2, 256 * 12);
        }

        private void CreateSheetStatistic(RegionType regionType)
        {
            ISheet sheet = null;
            int rowStart = 0;
            int rowEnd = 0;
            ICellStyle style = null;
            ICellStyle styleDouble = null;

            switch (regionType)
            {
                case RegionType.Rtk:
                    sheet = workbook.CreateSheet("РТК");
                    style = styleRtk;
                    styleDouble = styleDoubleRtk;
                    rowStart = separator;
                    rowEnd = table.Count;
                    break;
                case RegionType.Gts:
                    sheet = workbook.CreateSheet("ГТС");
                    style = styleGts;
                    styleDouble = styleDoubleGts;
                    rowStart = 0;
                    rowEnd = separator;
                    break;
                case RegionType.All:
                    sheet = workbook.CreateSheet("3счет");
                    rowStart = 0;
                    rowEnd = table.Count;
                    break;
            }

            string[] caption = Initializer.GetStatisticCaption();
            CreateCaption(caption, sheet, styleCaption);

            int row = 1;
            for (int line = rowStart; line < rowEnd; line++)
            {
                IRow currentRow = sheet.CreateRow(row++);
                List<string> list = table.ElementAt(line);

                for (int col = 0; col < list.Count; col++)
                {
                    ICell cell = currentRow.CreateCell(col);
                    string s = list.ElementAt(col);

                    if (2 < col && col < 8)
                        cell.SetCellValue(double.Parse(s));
                    else if (col == 8)
                        cell.CellFormula = string.Format("VLOOKUP(D{0},'МВЗ=Телефон'!A:B,2,(FALSE))", row);
                    else if (col == 9)
                        cell.CellFormula = string.Format("VLOOKUP(I{0},'МВЗ=ЗАКАЗ'!A:B,2,(FALSE))", row);
                    else
                        cell.SetCellValue(s);

                    if (regionType == RegionType.All)
                    {
                        if (line < separator)
                        {
                            style = styleGts;
                            styleDouble = styleDoubleGts;
                        }
                        else
                        {
                            style = styleRtk;
                            styleDouble = styleDoubleRtk;
                        }
                    }

                    cell.CellStyle = style;
                    if (col == 5)
                        cell.CellStyle = styleDouble;
                }
            }

            if (regionType == RegionType.All)
            {
                ICell cellTable = sheet.GetRow(1).CreateCell(10);
                cellTable.SetCellValue("местн");
                cellTable.CellStyle = styleGts;
                cellTable = sheet.GetRow(1).CreateCell(11);
                cellTable.CellFormula = string.Format("SUM(F2:F{0})", separator + 1);
                cellTable.CellStyle = styleDoubleGts;

                cellTable = sheet.GetRow(2).CreateCell(10);
                cellTable.SetCellValue("мг мн");
                cellTable.CellStyle = styleRtk;
                cellTable = sheet.GetRow(2).CreateCell(11);
                cellTable.CellFormula = string.Format("SUM(F{0}:F{1})", separator + 2, row);
                cellTable.CellStyle = styleDoubleRtk;

                cellTable = sheet.GetRow(3).CreateCell(10);
                cellTable.SetCellValue("итого");
                cellTable.CellStyle = styleCaption;
                cellTable = sheet.GetRow(3).CreateCell(11);
                cellTable.CellFormula = string.Format("SUM(F{0}:F{1})", 2, row);
                cellTable.CellStyle = styleDoubleBold;
            }
            else
            {
                IRow sumRow = sheet.CreateRow(row++);
                ICell sumCell = sumRow.CreateCell(5);
                sumCell.CellStyle = styleDoubleBold;
                sumCell.CellFormula = string.Format("SUM(F{0}:F{1})", 2, row - 1);
                listSumRtkGts.Add(row);
            }

            sheet.SetColumnWidth(0, 256 * 9);
            sheet.SetColumnWidth(1, 256 * 9);
            sheet.SetColumnWidth(2, 256 * 21);
            sheet.SetColumnWidth(3, 256 * 8);
            sheet.SetColumnWidth(4, 256 * 8);
            sheet.SetColumnWidth(5, 256 * 8);
            sheet.SetColumnWidth(6, 256 * 13);
            sheet.SetColumnWidth(7, 256 * 13);
            sheet.SetColumnWidth(8, 256 * 13);
            sheet.SetColumnWidth(9, 256 * 13);
        }

        private void CreateSheetBasic(SheetName sheetName)
        {
            Dictionary<string, string> dictionary = null;
            string[] caption = null;
            ISheet sheet = null; ;

            switch (sheetName)
            {
                case SheetName.Corporate:
                    sheet = workbook.CreateSheet("Корпоративка");
                    dictionary = Initializer.GetCorporateDictionary();
                    caption = Initializer.GetCorporateCaption();
                    break;
                case SheetName.MvzOrder:
                    sheet = workbook.CreateSheet("МВЗ=ЗАКАЗ");
                    dictionary = Initializer.GetMvzOrderDictionary();
                    caption = Initializer.GetMvzOrderCaption();
                    break;
                case SheetName.MvzPhone:
                    sheet = workbook.CreateSheet("МВЗ=Телефон");
                    dictionary = Initializer.GetMvzPhoneDictionary();
                    caption = Initializer.GetMvzPhoneCaption();
                    break;
            }

            CreateCaption(caption, sheet, styleTable);

            int row = 1;
            foreach (KeyValuePair<string, string> kvp in dictionary)
            {
                IRow currentRow = sheet.CreateRow(row++);

                ICell cellKey = currentRow.CreateCell(0);
                cellKey.SetCellValue(int.Parse(kvp.Key));
                cellKey.CellStyle = styleTable;

                ICell cellValue = currentRow.CreateCell(1);
                if (kvp.Value != "")
                    cellValue.SetCellValue(double.Parse(kvp.Value));
                else
                    cellValue.SetCellValue(kvp.Value);

                cellValue.CellStyle = styleTable;
            }

            sheet.SetColumnWidth(0, 256 * 21);
            sheet.SetColumnWidth(1, 256 * 21);
        }

        private static void CreateCaption(string[] caption, ISheet sheet, ICellStyle style)
        {
            IRow captionRow = sheet.CreateRow(0);
            for (int col = 0; col < caption.Length; col++)
            {
                ICell cell = captionRow.CreateCell(col);
                cell.SetCellValue(caption[col]);
                cell.CellStyle = style;
            }
        }

        public void Save()
        {
            Save(GetOutFilePath());
        }

        public void Save(string outFilePath)
        {

            using (FileStream fs = File.Create(outFilePath))
            {
                workbook.Write(fs);
            }
        }

        private string GetOutFilePath()
        {
            string outFileName = dateString() + "-saz.xlsx";
            string outPath = Path.GetDirectoryName(fileName) + Path.DirectorySeparatorChar + outFileName;
            return outPath;
        }

        private string dateString()
        {
            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-1);
            string date = dt.ToString("yyyy-MM");
            return date;
        }
    }
}
