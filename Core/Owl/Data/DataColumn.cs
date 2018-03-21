using System;
using System.Reflection;

namespace Owl.Data
{
    public class DataColumn
    {
        private DataTable _table;
        private Type type;
        internal int _hashCode;

        public DataColumn(string columnName)
        {
            this.ColumnName = columnName;
        }
        public DataColumn(string columnName, Type type)
        {
            this.ColumnName = columnName;
            this.type = type;
        }

        public string ColumnName { get; private set; }
        public DataTable Table
        {
            get => _table;
            internal set => _table = value;
        }
        public bool IsValueType => type.GetTypeInfo().IsValueType;

        public Type DataType => type;
    }
}
