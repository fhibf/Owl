using System.Collections;
using System.Collections.Generic;

namespace Owl.Data
{
    public class DataRowCollection : ICollection<DataRow>
    {
        private readonly List<DataRow> _list;
        private readonly DataTable _table;

        public DataRowCollection(DataTable table)
        {
            this._table = table;
            _list = new List<DataRow>();
        }

        public DataRow this[int index]
        {
            get
            {
                return _list[index];
            }
        }

        public void Add(DataRow row)
        {
            _table.AddRow(row, -1);
            _list.Add(row);
        }
  
        #region Interface Methods

        public int Count => _list.Count;

        public bool IsReadOnly => false;

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(DataRow item)
        {
            return _list.Contains(item);
        }

        public void CopyTo(DataRow[] array, int arrayIndex)
        {
            _list.CopyTo(array, arrayIndex);
        }

        public IEnumerator<DataRow> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public bool Remove(DataRow item)
        {
            return _list.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }
}
