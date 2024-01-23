///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/6/1 星期四 17:33:29
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using System;

namespace ST.NPOI.Extension.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class TransNameAttribute : Attribute
    {
        public TransNameAttribute(string name)
        {
            Name = name;
        }

        public TransNameAttribute(string name, int index) : this(name)
        {
            Index = index;
        }

        public string Name { get; set; }

        public int Index { get; } = 99;
    }
}
