using Owl.Data;
using System.Collections.Generic;

namespace Owl.Word
{
    public class TableHeader
    {
        public List<List<HeaderColumn>> Rows { get; } = new List<List<HeaderColumn>>();

        public class HeaderColumn
        {
            public enum RowSpanStatus
            {
                None, Restart, Continue
            };
            public int ColSpan { get; set; } = 1;
            public RowSpanStatus RowSpan { get; set; } = RowSpanStatus.None;
            public DataColumn Column { get; set; }
            public string Title { get; set; } = string.Empty;
        }
    }
}
