using System;
using System.Collections;
using System.Collections.Generic;

namespace Owl.Data
{
    public class DataColumnCollection : ICollection
    {
        private readonly List<DataColumn> _list = new List<DataColumn>();
        private readonly Dictionary<string, DataColumn> _columnFromName;
        private readonly StringComparer _hashProvider = StringComparer.OrdinalIgnoreCase;
        private readonly DataTable _table;

        public DataColumnCollection(DataTable table)
        {
            _table = table;
            _columnFromName = new Dictionary<string, DataColumn>();
        }


        #region Index
        internal int IndexOfCaseInsensitive(string name)
        {
            int hashcode = name.GetHashCode();
            int cachedI = -1;
            DataColumn column = null;
            for (int i = 0; i < Count; i++)
            {
                column = (DataColumn)_list[i];
                if ((hashcode == 0 || column._hashCode == 0 || column._hashCode == hashcode))
                {
                    if (cachedI == -1)
                        cachedI = i;
                    else
                        return -2;
                }
            }
            return cachedI;
        }

        public DataColumn this[int index] => (DataColumn)_list[index];
        public DataColumn this[string name]
        {
            get
            {
                if (null == name)
                {
                    throw new ArgumentNullException(nameof(name));
                }
                DataColumn column;
                if ((!_columnFromName.TryGetValue(name, out column)) || (column == null))
                {
                    // Case-Insensitive compares
                    int index = IndexOfCaseInsensitive(name);
                    if (0 <= index)
                    {
                        column = (DataColumn)_list[index];
                    }
                }
                return column;
            }
            set
            {
                var column = this[name];
                column = value;
            }
        }

        public DataColumn this[DataColumn column]
        {
            get
            {
                return this[column.ColumnName];
            }
            set
            {
                this[column.ColumnName] = value;
            }
        }
        #endregion

        #region Add
        public void AddRange(DataColumn[] dataColumn)
        {
            foreach (var column in dataColumn)
                Add(column);
        }

        public DataColumn Add(string columnName, Type type)
        {
            DataColumn column = new DataColumn(columnName, type);
            Add(column);
            return column;
        }
        public DataColumn Add(string columnName)
        {
            DataColumn column = new DataColumn(columnName);
            Add(column);
            return column;
        }      
        public void Add(DataColumn column)
        {
            AddAt(-1, column);
        }
        internal void AddAt(int index, DataColumn column)
        {
            if (column != null)
            {
                BaseAdd(column);
                if (index != -1)
                    ArrayAdd(index, column);
                else
                    ArrayAdd(column);
            }
        }
        private void ArrayAdd(DataColumn column)
        {
            _list.Add(column);
        }

        private void ArrayAdd(int index, DataColumn column)
        {
            _list.Insert(index, column);
        }

        private void BaseAdd(DataColumn column)
        {
            if (column == null)
                throw new ArgumentNullException(nameof(column));

            column.Table = _table;
            RegisterColumnName(column.ColumnName, column);

        }
        internal void RegisterColumnName(string name, DataColumn column)
        {
            try
            {
                _columnFromName.Add(name, column);

                if (null != column)
                {
                    column._hashCode = _hashProvider.GetHashCode(name);
                }
            }
            catch (ArgumentException)
            { // Argument exception means that there is already an existing key
                if (_columnFromName[name] != null)
                {
                    if (column != null)
                    {
                        throw new Exception("Cannot Add Duplicate");
                    }
                }
            }
        }

        #endregion

        #region Interface Methods
        public int Count => _list.Count;

        public bool IsSynchronized => throw new NotImplementedException();

        public object SyncRoot => throw new NotImplementedException();

        public void CopyTo(Array array, int index)
        {
            var arr = _list.ToArray();
            arr.CopyTo(array, index);
        }

        public IEnumerator GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        #endregion


        public int IndexOf(DataColumn column)
        {
            int columnCount = _list.Count;
            for (int i = 0; i < columnCount; ++i)
            {
                if (column == (DataColumn)_list[i])
                {
                    return i;
                }
            }
            return -1;
        }

        public DataColumnCollection Clone()
        {
            return (DataColumnCollection)this.MemberwiseClone();
        }
    }
}