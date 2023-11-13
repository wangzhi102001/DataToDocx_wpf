using DataToDocx.Models;
using System.Collections.Generic;
using System.Data;

namespace DataToDocx
{
    public class Database
    {

        public string Name { get; set; }
        public string Path { get; set; }
        public List<DataTab>? Tabs { get; set; }
        public Database(string name, string path) { Name = name; Path = path; }

        public override string ToString()
        {
            return Name;
        }

        

    }
}
