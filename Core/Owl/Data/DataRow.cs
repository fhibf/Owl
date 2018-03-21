using System;
using System.Collections.Generic;

namespace Owl.Data
{
    public class DataRow
    {
        private readonly DataTable _table;
        private readonly DataColumnCollection _columns;
        private readonly Dictionary<DataColumn, object> _storage;
        internal int _tempRecord = -1;
        internal int _newRecord = -1;
        internal long _rowID = -1;

        internal DataRow(DataRowBuilder builder)
        {
            _tempRecord = builder._record;
            _table = builder._table;
            _storage = new Dictionary<DataColumn, object>();
            _columns = builder._table.Columns;
            foreach (DataColumn col in _columns)
                _storage.Add(col, null);
        }
        public object this[string columnName]
        {
            get
            {
                var column = GetDataColumn(columnName);
                return this[column];
             }
            set
            {
                DataColumn column = GetDataColumn(columnName);
                this[column] = value;
            }
        }
        public object this[DataColumn column]
        {
            get
            {
                return _storage[column];
            }
            set
            {
                if (column.Table != _table)
                {
                    // user removed column from table during OnColumnChanging event
                    throw new Exception("ColumnNotInTheTable");// (column.ColumnName, _table.TableName);
                }

                object proposed = value;
                if (null == proposed)
                {
                    if (column.IsValueType)
                    { // WebData 105963
                        throw new Exception("CannotSetToNull");
                    }
                }

                int record = _tempRecord;
                _storage[column] = proposed;
            }
        }
        public object this[int record]
        {
            get
            {
                var column = _columns[record];
                return this[column];                
            }
            set
            {
                var column = _columns[record];
                this[column] = value;
            }
        }


        internal DataColumn GetDataColumn(string columnName)
        {
            DataColumn column = _columns[columnName];
            if (null != column)
            {
                return column;
            }
            throw new Exception("ColumnNotInTheTable"); // (columnName, _table.TableName);
        }

        internal int GetDefaultRecord()
        {
            if (_tempRecord != -1)
                return _tempRecord;
            if (_newRecord != -1)
            {
                return _newRecord;
            }

            throw new Exception("Row Removed");
        }

        public long rowID
        {
            get => _rowID;
            set => _rowID = value;
        }
        public DataTable Table => _table;

        public sealed class DataRowBuilder
        {
            internal readonly DataTable _table;
            internal int _record;

            internal DataRowBuilder(DataTable table, int record)
            {
                _table = table;
                _record = record;
            }
        }

    }
}
