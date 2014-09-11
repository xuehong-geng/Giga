using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Test.Transformer
{
    public class TestTabularData
    {
        public String PO { get; set; }
        public int Item { get; set; }
        public String ProductCode { get; set; }
        public String ProductName { get; set; }
        public double Weight { get; set; }
        public int Qty { get; set; }
        public double UnitPrice { get; set; }
        public double Total { get; set; }
        public DateTime PODate { get; set; }
        public DateTime DueDate { get; set; }
    }
}
