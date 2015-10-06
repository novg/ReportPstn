using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

namespace ReportPstn
{
    sealed class CellStyles
    {
        IWorkbook workbook;
      
        public CellStyles(IWorkbook workbook)
        {
            this.workbook = workbook;
        }

        internal ICellStyle ResultHeadline()
        {
            ICellStyle style = Headline();
            style.Alignment = HorizontalAlignment.Right;
            style.GetFont(workbook).IsItalic = true;
            return style;
        }

        internal ICellStyle InnerHeadline()
        {
            ICellStyle style = Headline();
            style.GetFont(workbook).FontHeight = 8;
            return style;
        }

        internal ICellStyle CorpColor()
        {
            IFont font = Heading().GetFont(workbook);
            font.Color = IndexedColors.Green.Index;
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = IndexedColors.LightYellow.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }

        internal ICellStyle HeadingSum()
        {
            ICellStyle style = Heading();
            style.Alignment = HorizontalAlignment.Right;
            style.GetFont(workbook).IsItalic = true;
            return style;
        }

        internal ICellStyle LimitsDoubleSum()
        {
            ICellStyle style = LimitsDoubleBold();
            style.Alignment = HorizontalAlignment.Right;
            style.GetFont(workbook).IsItalic = true;
            return style;
        }

        internal ICellStyle LimitsDoubleBold()
        {
            ICellStyle style = Heading();
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            style.GetFont(workbook).Boldweight = (short)FontBoldWeight.Bold;
            return style;
        }

        internal ICellStyle LimitsDouble()
        {
            IFont font = LimitsTable().GetFont(workbook);
            ICellStyle style = Double();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(font);
            return style;
        }

        internal ICellStyle LimitsTable()
        {
            ICellStyle style = Table();
            style.Alignment = HorizontalAlignment.Left;
            style.GetFont(workbook).FontName = "Times New Roman";
            return style;
        }

        internal ICellStyle Heading()
        {
            IFont font = Headline().GetFont(workbook);
            font.FontHeight = 8;
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Medium;
            style.BottomBorderColor = IndexedColors.Black.Index;
            style.BorderLeft = BorderStyle.Medium;
            style.LeftBorderColor = IndexedColors.Black.Index;
            style.BorderRight = BorderStyle.Medium;
            style.RightBorderColor = IndexedColors.Black.Index;
            style.BorderTop = BorderStyle.Medium;
            style.TopBorderColor = IndexedColors.Black.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
            style.SetFont(font);
            return style;
        }

        internal ICellStyle Headline()
        {
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeight = 10;
            font.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            return style;
        }

        internal ICellStyle Table()
        {
            IFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.FontHeight = 8;
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = IndexedColors.Black.Index;
            style.BorderLeft = BorderStyle.Thin;
            style.LeftBorderColor = IndexedColors.Black.Index;
            style.BorderRight = BorderStyle.Thin;
            style.RightBorderColor = IndexedColors.Black.Index;
            style.BorderTop = BorderStyle.Thin;
            style.TopBorderColor = IndexedColors.Black.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
            style.SetFont(font);
            return style;
        }

        internal ICellStyle Caption()
        {
            ICellStyle style = Table();
            style.GetFont(workbook).Boldweight = (short)FontBoldWeight.Bold;
            return style;
        }

        internal ICellStyle Double()
        {
            ICellStyle style = Table();
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            return style;
        }

        internal ICellStyle Gts()
        {
            ICellStyle style = Table();
            style.FillForegroundColor = IndexedColors.LightGreen.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }

        internal ICellStyle DoubleGts()
        {
            ICellStyle style = Gts();
            style.DataFormat = Double().DataFormat;
            return style;
        }

        internal ICellStyle Rtk()
        {
            ICellStyle style = Table();
            style.FillForegroundColor = IndexedColors.LightYellow.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }

        internal ICellStyle DoubleRtk()
        {
            ICellStyle style = Rtk();
            style.DataFormat = Double().DataFormat;
            return style;
        }

        internal ICellStyle DoubleBold()
        {
            ICellStyle style = Caption();
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            return style;
        }
    }
}
