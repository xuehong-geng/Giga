using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using Giga.Transformer.Configuration;

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
        public DateTime? PODate { get; set; }
        public DateTime? DueDate { get; set; }
    }

    public class RdPurchaseOrder
    {
        public String Id { get; set; }
        public String Version { get; set; }
        public String ShipTo { get; set; }
        public String ExtPo { get; set; }
        public DateTime PoDate { get; set; }
        public String Payment { get; set; }
        public String Delivery { get; set; }
        public String IncoTerms { get; set; }
        public String AdditionalNotes { get; set; }
        public String Currency { get; set; }

        public class Item
        {
            public DateTime ExwDate { get; set; }
            public int LineNumber { get; set; }
            public String Part { get; set; }
            public String Description { get; set; }
            public int Qty { get; set; }
            public double UnitPrice { get; set; }
            public double TotalPrice { get; set; }
        }

        public IEnumerable<Item> Items { get; set; }

        public override string ToString()
        {
            var b = new StringBuilder();
            b.AppendFormat("PO Number:{0}\t", Id);
            b.AppendFormat("Version:{0}\n", Version);
            b.AppendFormat("Ship to:{0}\n", ShipTo);
            b.AppendFormat("Ext PO:{0}\r", ExtPo);
            b.AppendFormat("PO Date:{0}\n", PoDate);
            b.AppendFormat("Payment:{0}\n", Payment);
            b.AppendFormat("Delivery:{0}\n", Delivery);
            b.AppendFormat("IncoTerms:{0}\n", IncoTerms);
            b.AppendFormat("---------------------Items-----------------------\n");
            b.AppendLine("ExwDate\t\tLineNumber\t\tPart\t\tDescription\t\tQty\t\tUnitPrice\t\tTotalPrice");
            foreach (var item in Items)
            {
                b.AppendFormat("{0}\t\t{1}\t\t{2}\t\t{3}\t\t{4}\t\t{5}\t\t{6}\n",
                    item.ExwDate, item.LineNumber, item.Part, item.Description, item.Qty, item.UnitPrice,
                    item.TotalPrice);
            }
            b.AppendFormat("--------------------Additional Notes---------------\n");
            b.AppendLine(AdditionalNotes);
            return b.ToString();
        }

        public RdPurchaseOrder()
        {
            Items = new List<Item>();
        }

        /// <summary>
        /// Load data from excel file
        /// </summary>
        /// <param name="filePath">Path of excel file</param>
        /// <returns></returns>
        public static RdPurchaseOrder Load(String filePath)
        {
            // Get configuration
            var cfg =
                ConfigurationManager.GetSection("Giga.Transformer") as TransformerConfigSection;
            if (cfg == null)
                throw new ConfigurationErrorsException("<Giga.Transformer> not exist in configuration!");
            // Load entities from file
            var transformer = new Giga.Transformer.Transformer(cfg);
            return transformer.LoadOne<RdPurchaseOrder>(filePath, "RdPurchaseOrder");
        }
    }
}
