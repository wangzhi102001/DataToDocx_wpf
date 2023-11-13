using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataToDocx.Models
{
    public class DataTab : ObservableObject
    {

        public string TabName { get; set; }
        public List<string>? Atts { get; set; }

        private int _count = 0;
        public int Count
        {

            get { return _count; }
            set { _count = value; OnPropertyChanged(); }

        }
        public DataTable TabContent { get; set; }
        public DataTab() { }

        public DataTab(string tabName) { TabName = tabName; }
    }
}
