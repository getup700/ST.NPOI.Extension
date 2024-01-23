///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/6/21 星期三 22:09:42
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using System;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Collections;
using ST.Net.Extension.Utils;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;

namespace ST.NPOI.Extension.Attributes
{
    public static class TransNameAttributeExtension
    {
        private static Dictionary<object, object?>? getKeyValue(object obj)
        {
            if (obj == null)
            {
                return null;
            }
            var type = obj.GetType();
            var members = TypeUtils.GetTargetMembers(type, x => x.GetCustomAttribute<TransNameAttribute>() != null)
                ?.OrderBy(x => x.GetCustomAttribute<TransNameAttribute>().Index)
                ?.ToList();
            if (members.Count == 0)
            {
                return null;
            }
            var result = new Dictionary<object, object>();
            foreach (var member in members)
            {
                var transName = member.GetCustomAttribute<TransNameAttribute>();
                if (member is PropertyInfo propertyInfo)
                {
                    result.Add(transName.Name, propertyInfo.GetValue(obj));
                }
                else if (member is FieldInfo fieldInfo)
                {
                    result.Add(transName.Name, fieldInfo.GetValue(obj));
                }
            }
            return result;
        }

        public static DataTable ToDataTable<T>(T t) where T : class
        {
            if (t == null)
            {
                throw new ArgumentNullException(nameof(t));
            }
            var dic = getKeyValue(t)
                ?? throw new ArgumentNullException($"There are no properties or fields with the TransName attribute in {nameof(t)}");
            var dataTable = new DataTable();
            foreach (var pair in dic)
            {
                dataTable.Columns.Add(pair.Key.ToString());
            }
            dataTable.Rows.Add(dic.Values.ToArray());
            return dataTable;
        }

        public static DataTable ToDataTable<T>(this IEnumerable<T> elements) where T : class
        {
            if (elements.Count() == 0)
            {
                return null;
            }
            var keyValue = getKeyValue(elements.First())
                ?? throw new ArgumentNullException($"There are no properties or fields with the TransName attribute in {nameof(elements)}");
            var dataTable = new DataTable();
            foreach (var property in keyValue)
            {
                dataTable.Columns.Add(new DataColumn(property.Key.ToString(), (property.Key).GetType()));
            }
            foreach (var element in elements)
            {
                var pair = getKeyValue(element);
                dataTable.Rows.Add(pair.Values.ToArray());
            }
            return dataTable;
        }

        /*
        public static List<List<string>> InterpreteData<T>(this IEnumerable<T> elements, Predicate<PropertyInfo> predicate = null)
        {
            var type = typeof(T);
            var properties = type.GetProperties()?.ToList();
            if (predicate != null)
            {
                properties = properties.Where(x => predicate(x))?.ToList();
            }
            if (properties.Count == 0)
            {
                throw new Exception(nameof(elements));
            }
            var data = new List<List<string>>();
            var headerRowContent = properties.Select(x => x.Name).ToList();
            data.Add(headerRowContent);
            //interprete data
            foreach (var t in elements)
            {
                var propertyInfoes = t.GetType().GetProperties().ToList();
                if (predicate != null)
                {
                    propertyInfoes = propertyInfoes.Where(x => predicate(x)).ToList();
                }
                var elementList = new List<List<string>>();

                #region initial elementList Count
                var itemCount = 1;
                foreach (var property in propertyInfoes)
                {
                    var propertyType = property.PropertyType;
                    if (propertyType.IsGenericType && propertyType.GetGenericArguments().Count() == 1)
                    {
                        var nestedTypes = property.GetValue(t) as IEnumerable;
                        int count = 0;
                        foreach (var item in nestedTypes)
                        {
                            count++;
                        }
                        itemCount = itemCount * count;
                    }
                }
                for (int i = 0; i < itemCount; i++)
                {
                    elementList.Add(new List<string>());
                }
                #endregion

                foreach (var property in propertyInfoes)
                {
                    var propertyType = property.PropertyType;
                    var valueList = new List<string>();
                    //if property is IEnumerable type
                    if (propertyType.IsGenericType && propertyType.GetGenericArguments().Count() == 1)
                    {
                        var nestedTypes = property.GetValue(t) as IEnumerable;
                        if (nestedTypes == null)
                        {
                            continue;
                        }
                        var children = new List<object>();
                        foreach (var item in nestedTypes)
                        {
                            children.Add(item);
                        }

                        foreach (var item in elementList)
                        {
                            var index = elementList.IndexOf(item);
                            int childIndex = 0;
                            if (index < children.Count)
                            {
                                childIndex = index;
                            }
                            else
                            {
                                childIndex = (index + 1) % children.Count;
                            }
                            item.Add(children[index].ToString());
                        }
                    }
                    else
                    {
                        foreach (var item in elementList)
                        {
                            item.Add(property.GetValue(t)?.ToString());
                        }
                    }
                }
                data.AddRange(elementList);
            }
            return data;
        }
        */
    }
}
