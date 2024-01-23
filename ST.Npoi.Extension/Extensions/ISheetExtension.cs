///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/3/24 16:40:27
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;

namespace ST.NPOI.Extension.Extensions
{
    public static class ISheetExtension
    {
        public static int? GetFirstBlankRowIndex(this ISheet sheet, int superiorLimit = 10000)
        {
            int i = 0;
        R: var row = sheet.GetRow(i);
            if (row != null)
            {
                i++;
                if (i > superiorLimit)
                {
                    return null;
                }
                else
                {
                    goto R;
                }
            }
            return i;
        }

        public static List<IRow> GetRows(this ISheet sheet, Predicate<IRow> predicate = null)
        {
            var rows = new List<IRow>();
            IEnumerator enumerator = sheet.GetRowEnumerator();
            while (enumerator.MoveNext())
            {
                IRow? row = enumerator.Current as IRow;
                if (predicate != null)
                {
                    if (predicate(row))
                    {
                        rows.Add(row);
                    }
                }
                else
                {
                    rows.Add(row);
                }
            }
            return rows;
        }



    }
}
