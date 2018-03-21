using System;
using System.Globalization;
using static Owl.Data.DataRow;

namespace Owl.Data
{
    public class DataTable
    {
        internal readonly DataRowCollection rowCollection;
        internal readonly DataColumnCollection columnCollection;
        private readonly DataRowBuilder rowBuilder;

        private DataSet dataSet;
        private string tableName = "";
        internal string tableNamespace = null;
        private string tablePrefix = "";
        internal long _nextRowID;

        public DataTable()
        {
            rowCollection = new DataRowCollection(this);
            columnCollection = new DataColumnCollection(this);

            rowBuilder = new DataRowBuilder(this, -1);

            _nextRowID = 1;
        }
        public DataTable(string name) : this()
        {
            tableName = name;
        }
        public DataRowCollection Rows
        {
            get
            {
                return rowCollection;
            }
        }

        public DataColumnCollection Columns
        {
            get
            {
                return columnCollection;
            }
        }

        #region Add
        public void AddRow(DataRow row)
        {
            rowCollection.Add(row);
            //AddRow(row, -1);
        }
        internal void AddRow(DataRow row, int proposedID)
        {
            InsertRow(row, proposedID, -1);
        }
        internal void InsertRow(DataRow row, long proposedID, int pos)
        {
            Exception deferredException = null;

            if (row == null)
            {
                throw new ArgumentNullException("row");
            }
            if (row.Table != this)
            {
                throw new Exception("RowAlreadyInOtherCollection");
            }
            if (row.rowID != -1)
            {
                throw new Exception("RowAlreadyInOtherCollection");
            }

            int record = row._tempRecord;
            row._tempRecord = -1;

            if (proposedID == -1)
            {
                proposedID = _nextRowID;
            }

            _nextRowID = checked(proposedID + 1);


            try
            {
                row.rowID = proposedID;
            }
            catch
            {
                row.rowID = -1;
                row._tempRecord = record;
                throw;
            }

            // since expression evaluation occurred in SetNewRecordWorker, there may have been a problem that
            // was deferred to this point.  If so, throw now since row has already been added.
            if (deferredException != null)
                throw deferredException;

        }
        #endregion

        public DataRow NewRow()
        {
            DataRow dr = NewRow(-1);
            return dr;
        }
        internal DataRow NewRow(int record)
        {
            rowBuilder._record = record;
            DataRow row = NewRowFromBuilder(rowBuilder);
            return row;
        }
        protected virtual DataRow NewRowFromBuilder(DataRowBuilder builder)
        {
            return new DataRow(builder);
        }

        //internal int NewRecordFromArray(object[] value)
        //{
        //    int colCount = columnCollection.Count; // Perf: use the readonly columnCollection field directly
        //    if (colCount < value.Length)
        //    {
        //        throw new Exception("ValueArrayLength");
        //    }
        //    int record = recordManager.NewRecordBase();
        //    try
        //    {
        //        for (int i = 0; i < value.Length; i++)
        //        {
        //            if (null != value[i])
        //            {
        //                columnCollection[i][record] = value[i];
        //            }
        //            else
        //            {
        //                columnCollection[i].Init(record);  // Increase AutoIncrementCurrent
        //            }
        //        }
        //        for (int i = value.Length; i < colCount; i++)
        //        {
        //            columnCollection[i].Init(record);
        //        }
        //        return record;
        //    }
        //    catch (Exception e)
        //    {
        //        // 
        //        if (Common.ADP.IsCatchableOrSecurityExceptionType(e))
        //        {
        //            FreeRecord(ref record); // WebData 104246
        //        }
        //        throw;
        //    }
        //}

        public string TableName
        {
            get => tableName;
            set
            {
                if (value == null)
                {
                    value = "";
                }

                if (String.Compare(tableName, value, true) != 0)
                {
                    if (dataSet != null)
                    {
                        if (value.Length == 0)
                            throw new Exception("NoTableName");
                        if ((0 == String.Compare(value, dataSet.DataSetName, true)))
                            throw new Exception($"DatasetConflictingName({dataSet.DataSetName}");
                    }
                    tableName = value;
                }
            }
        }
        public string Namespace { get; internal set; }
    }
}
