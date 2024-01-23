///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/5/27 星期六 10:47:30
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Linq;

namespace ST.NPOI.Extension.Extensions
{
    public static class ExcelStyleExtension
    {
        public static IFont CreateIFont(this IWorkbook workbook, Action<IFont> action = null)
        {
            IFont font = workbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            if (action != null)
            {
                action.Invoke(font);
            }
            return font;
        }

        public static ICellStyle CreateICellStyle(this IWorkbook workbook, Action<ICellStyle> action = null)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.SetFont(workbook.CreateIFont());
            if (action != null)
            {
                action.Invoke(cellStyle);
            }
            return cellStyle;
        }

        public static ICellStyle CommonHeaderStyle(this IWorkbook workbook, Action<ICellStyle> action = null)
        {
            var icellStyle = workbook.CreateICellStyle(x =>
            {
                x.FillPattern = FillPattern.SolidForeground;
                x.FillForegroundColor = HSSFColor.LightGreen.Index;
                x.WrapText = true;

            });
            if (action != null)
            {
                action.Invoke(icellStyle);
            }
            return icellStyle;
        }

        public static ICellStyle CommonTextStyle(this IWorkbook workbook, Action<ICellStyle> action = null)
        {
            var icellStyle = workbook.CreateICellStyle(x =>
            {
                x.WrapText = false;
            });
            if (action != null)
            {
                action.Invoke(icellStyle);
            }
            return icellStyle;
        }

        public static IRow SetRowStyle(this IRow row, ICellStyle cellStyle)
        {
            var cells = row.GetCells().Select(x => x.CellStyle = cellStyle);
            return row;
        }

        public static IRow AutoColumnWidth(this IRow row)
        {
            var sheet = row.Sheet;
            var columns = row.GetCells().Count;
            for (int i = 0; i < columns; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            return row;
        }

        public static IRow SetRowAutoFilter(this IRow row)
        {
            var index = row.GetRowIndex();
            var sheet = row.Sheet;
            var cellRangeAdress = new CellRangeAddress(0, index + 1, 0, row.GetCells().Count - 1);
            sheet.SetAutoFilter(cellRangeAdress);
            return row;
        }

        public static IRow SetFreezePane(this IRow row)
        {
            var index = row.GetRowIndex();
            var sheet = row.Sheet;
            sheet.CreateFreezePane(0, index + 1);
            return row;
        }


    }
}
