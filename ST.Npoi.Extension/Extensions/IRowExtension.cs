///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/3/24 15:33:23
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ST.NPOI.Extension.Extensions
{
    public static class IRowExtension
    {
        public static int AssignValuesInOrder(this IRow row, IEnumerable<string> values, int startColumn = 1)
        {
            int init = startColumn;
            foreach (var item in values)
            {
                ICell cell = row.CreateCell(startColumn);
                cell.SetCellValue(item);
                startColumn++;
            }
            return startColumn - init;
        }

        /// <summary>
        /// Get all cells,even if they are null
        /// </summary>
        /// <param name="row"></param>
        /// <param name="predicate"></param>
        /// <returns></returns>
        public static List<ICell> GetCells(this IRow row, Predicate<ICell> predicate = null)
        {
            var cells = new List<ICell>();
            var count = row.LastCellNum;
            for (int i = 0; i < count; i++)
            {
                var cell = row.GetCell(i);
                cells.Add(cell);
            }
            return cells;
        }
        public static int GetRowIndex(this IRow row)
        {
            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }
            var sheet = row.Sheet;
            var rows = sheet.GetRows();
            var index = rows.IndexOf(row);
            if (index < 0 || index > rows.Count - 1)
            {
                throw new IndexOutOfRangeException(nameof(index));
            }
            return index;
        }
    }
}
