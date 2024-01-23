///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/3/30 18:18:20
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using NPOI.SS.UserModel;
using System;

namespace ST.NPOI.Extension.Extensions
{
    public static class ICellExtension
    {
        public static int GetCellIndex(this ICell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }
            var row = cell.Row;
            var cells = row.GetCells();
            var index = cells.IndexOf(cell);
            if (index < 0 || index > cells.Count - 1)
            {
                throw new IndexOutOfRangeException(nameof(index));
            }
            return index;
        }
        public static string GetCellValue(this ICell cell)
        {
            var result = string.Empty;
            if (cell == null)
            {
                return string.Empty;
            }
            var cellType = cell.CellType;
        A: switch (cellType)
            {
                case CellType.Formula:
                    cellType = cell.CachedFormulaResultType;
                    goto A;
                case CellType.Numeric:
                    result = cell.NumericCellValue.ToString();
                    break;
                case CellType.Boolean:
                    result = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    result = string.Empty;
                    break;
                case CellType.Unknown:
                    result = cell.StringCellValue;
                    break;
                case CellType.String:
                    result = cell.StringCellValue;
                    break;
                case CellType.Blank:
                    result = string.Empty;
                    break;
                default:
                    result = cell.StringCellValue;
                    break;
            }
            return result;
        }
    }
}
