///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/3/22 17:58:22
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ST.NPOI.Extension.Extensions
{
    public static class ContentExtension
    {

        public static IRow SetCellValue(this IRow row, IEnumerable<string> values, int startCell = 1, ICellStyle cellStyle = null)
        {
            var count = values.Count();
            foreach (var value in values)
            {
                var cell = row?.CreateCell(startCell);
                cell.SetCellValue(value);
                if (cellStyle != null)
                {
                    cell.CellStyle = cellStyle;
                }
                startCell++;
            }
            return row;
        }
    }
}
