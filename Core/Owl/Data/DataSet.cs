using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Owl.Data
{
    public class DataSet
    {
        public DataSet()
        {
            Tables = new DataTableCollection(this);
        }
        public string DataSetName { get; set; }

        public DataTableCollection Tables { get; }

        public CultureInfo Locale = CultureInfo.CurrentCulture;
    }
}
