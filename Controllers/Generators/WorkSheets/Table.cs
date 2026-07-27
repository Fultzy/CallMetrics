using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CallMetrics.Controllers.Generators.WorkSheets
{
    public enum TabletType
    {
        Standard,
        Averages
    }

    internal class Table
    {
        public string Name { get; set; }
        public int RowStart { get; set; } = 0;
        public int RowEnd { get; set; }
        public int ColumnStart { get; set; }
        public int ColumnEnd { get; set; }

        public TabletType Type { get; set; } = TabletType.Standard;

        public List<string> Headers { get; set; }
        public List<object> Subjects { get; set; }


    }
}
