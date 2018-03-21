using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Owl.Data
{
    public class DataTableCollection : ICollection<DataTable>
    {
        private readonly DataSet _dataSet = null;
        private readonly List<DataTable> _list;

        public DataTableCollection(DataSet dataSet)
        {
            _dataSet = dataSet;
            _list = new List<DataTable>();
        }

        public DataTable this[int index]
        {
            get
            {
                try
                { // Perf: use the readonly _list field directly and let ArrayList check the range
                    return (DataTable)_list[index];
                }
                catch (ArgumentOutOfRangeException)
                {
                    throw new Exception($"TableOutOfRange({index})");
                }
            }
        }

        /// <devdoc>
        ///    <para>Gets the table in the collection with the given name (not case-sensitive).</para>
        /// </devdoc>
        public DataTable this[string name]
        {
            get
            {
                int index = InternalIndexOf(name);
                if (index == -2)
                {
                    throw new Exception($"CaseInsensitiveNameConflict({name})");
                }
                if (index == -3)
                {
                    throw new Exception($"NamespaceNameConflict({name})");
                }
                return (index < 0) ? null : (DataTable)_list[index];
            }
        }

        public DataTable this[string name, string tableNamespace]
        {
            get
            {
                if (tableNamespace == null)
                    throw new ArgumentNullException(nameof(tableNamespace));
                int index = InternalIndexOf(name, tableNamespace);
                if (index == -2)
                {
                    throw new Exception($"CaseInsensitiveNameConflict({name})");
                }
                return (index < 0) ? null : (DataTable)_list[index];
            }
        }

        internal int InternalIndexOf(string tableName)
        {
            int cachedI = -1;
            if ((null != tableName) && (0 < tableName.Length))
            {
                int count = _list.Count;
                int result = 0;
                for (int i = 0; i < count; i++)
                {
                    DataTable table = (DataTable)_list[i];
                    result = NamesEqual(table.TableName, tableName, false, _dataSet.Locale);
                    if (result == 1)
                    {
                        // ok, we have found a table with the same name.
                        // let's see if there are any others with the same name
                        // if any let's return (-3) otherwise...
                        for (int j = i + 1; j < count; j++)
                        {
                            DataTable dupTable = (DataTable)_list[j];
                            if (NamesEqual(dupTable.TableName, tableName, false, _dataSet.Locale) == 1)
                                return -3;
                        }
                        //... let's just return i
                        return i;
                    }

                    if (result == -1)
                        cachedI = (cachedI == -1) ? i : -2;
                }
            }
            return cachedI;
        }
        // Return value:
        //      >= 0: find the match
        //        -1: No match
        //        -2: At least two matches with different cases
        internal int InternalIndexOf(string tableName, string tableNamespace)
        {
            int cachedI = -1;
            if ((null != tableName) && (0 < tableName.Length))
            {
                int count = _list.Count;
                int result = 0;
                for (int i = 0; i < count; i++)
                {
                    DataTable table = (DataTable)_list[i];
                    result = NamesEqual(table.TableName, tableName, false, _dataSet.Locale);
                    if ((result == 1) && (table.Namespace == tableNamespace))
                        return i;

                    if ((result == -1) && (table.Namespace == tableNamespace))
                        cachedI = (cachedI == -1) ? i : -2;
                }
            }
            return cachedI;

        }

        // Return value: 
        // > 0 (1)  : CaseSensitve equal      
        // < 0 (-1) : Case-Insensitive Equal
        // = 0      : Not Equal
        internal int NamesEqual(string s1, string s2, bool fCaseSensitive, CultureInfo locale)
        {
            if (fCaseSensitive)
            {
                if (String.Compare(s1, s2, false) == 0)
                    return 1;
                else
                    return 0;
            }

            // Case, kana and width -Insensitive compare
            if (locale.CompareInfo.Compare(s1, s2,
                CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreWidth) == 0)
            {
                if (String.Compare(s1, s2, false) == 0)
                    return 1;
                else
                    return -1;
            }
            return 0;
        }


        #region Interface Members
        public int Count => _list.Count;

        public bool IsReadOnly => throw new NotImplementedException();

        public void Add(DataTable item)
        {
            if (!_list.Contains(item))
                _list.Add(item);
        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(DataTable item)
        {
            return _list.Contains(item);
        }

        public void CopyTo(DataTable[] array, int arrayIndex)
        {
            _list.CopyTo(array, arrayIndex);
        }

        public IEnumerator<DataTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public bool Remove(DataTable item)
        {
            return _list.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
        #endregion
    }
}
