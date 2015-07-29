using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;
using Giga.Transformer.Configuration;
using Giga.Transformer.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Giga.Test.Transformer
{
    [TestClass]
    public class ExcelTransformerTest
    {
        [TestMethod]
        public void TestAA2Num()
        {
            Assert.AreEqual("A", Alph26.N2A(1));
            Assert.AreEqual("J", Alph26.N2A(10));
            Assert.AreEqual("Z", Alph26.N2A(26));
            Assert.AreEqual("AA", Alph26.N2A(27));
            Assert.AreEqual(1, Alph26.A2N("A"));
            Assert.AreEqual(2, Alph26.A2N("b"));
            Assert.AreEqual(26, Alph26.A2N("Z"));
            Assert.AreEqual(27, Alph26.A2N("aA"));
            String str = "AA";
            int n = Alph26.A2N(str);
            String newStr = Alph26.N2A(n);
            Assert.AreEqual(str, newStr);

            str = "AAA";
            n = Alph26.A2N(str);
            newStr = Alph26.N2A(n);
            Assert.AreEqual(str, newStr);
        }

        protected String GetTestFilePath(String fileName)
        {
            String binDir = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
            String fileDir = Path.Combine(binDir, "Transformer");
            return Path.Combine(fileDir, fileName);
        }

        [TestMethod]
        public void TestReadTabularData_FixedRange()
        {
            // Get configuration
            var cfg =
                ConfigurationManager.GetSection("Giga.Transformer") as TransformerConfigSection;
            if (cfg == null)
                throw new ConfigurationErrorsException("<Giga.Transformer> not exist in configuration!");

            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            // Load entities from file
            var transformer = new Giga.Transformer.Transformer(cfg);
            var entities = transformer.Load<TestTabularData>(filePath, "TestNormalTabularData_FixedRange");
            foreach (TestTabularData entity in entities)
            {
                var serializer = new DataContractJsonSerializer(typeof(TestTabularData));
                var memStrm = new MemoryStream();
                var writer = new StreamWriter(memStrm, Encoding.UTF8);
                serializer.WriteObject(memStrm, entity);
                byte[] buf = memStrm.GetBuffer();
                String xmlStr = Encoding.UTF8.GetString(buf);
                Console.WriteLine(xmlStr);
            }
        }

        [TestMethod]
        public void TestWriteTabularData_FixedRange()
        {
            // Get configuration
            var cfg =
                ConfigurationManager.GetSection("Giga.Transformer") as TransformerConfigSection;
            if (cfg == null)
                throw new ConfigurationErrorsException("<Giga.Transformer> not exist in configuration!");
            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            // Create a new file 
            var newFilePath = GetTestFilePath("TransformerTest_WriteTabular.xlsx");
            if (File.Exists(newFilePath))
                File.Delete(newFilePath);
            File.Copy(filePath, newFilePath);
            // Load entities from old file
            var transformer = new Giga.Transformer.Transformer(cfg);
            var entities = transformer.Load<TestTabularData>(filePath, "TestNormalTabularData_FixedRange").ToList();
            foreach (TestTabularData entity in entities)
            {
                entity.DueDate += new TimeSpan(1, 0, 0, 0);
                entity.Item += 1;
                entity.PO += "_new";
                entity.PODate += new TimeSpan(1, 0, 0, 0);
                entity.ProductCode += "_new";
                entity.ProductName += "_new";
                entity.Qty += 1;
                entity.Total += 1;
                entity.UnitPrice += 1;
                entity.Weight += 1;
            }
            // Write entities to new file
            transformer.Save(newFilePath, "TestNormalTabularData_FixedRange", entities);
        }

        [TestMethod]
        public void TestReadTabularData_DynamicRange()
        {
            // Get configuration
            var cfg =
                ConfigurationManager.GetSection("Giga.Transformer") as TransformerConfigSection;
            if (cfg == null)
                throw new ConfigurationErrorsException("<Giga.Transformer> not exist in configuration!");
            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            // Load entities from file
            var transformer = new Giga.Transformer.Transformer(cfg);
            var entities = transformer.Load<TestTabularData>(filePath, "TestNormalTabularData_DynamicRange");
            foreach (TestTabularData entity in entities)
            {
                var serializer = new DataContractJsonSerializer(typeof(TestTabularData));
                var memStrm = new MemoryStream();
                serializer.WriteObject(memStrm, entity);
                byte[] buf = memStrm.GetBuffer();
                String xmlStr = Encoding.UTF8.GetString(buf);
                Console.WriteLine(xmlStr);
            }
        }

        [TestMethod]
        public void TestWriteTabularData_DynamicRange()
        {
            // Get configuration
            var cfg =
                ConfigurationManager.GetSection("Giga.Transformer") as TransformerConfigSection;
            if (cfg == null)
                throw new ConfigurationErrorsException("<Giga.Transformer> not exist in configuration!");
            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            // Create a new file 
            var newFilePath = GetTestFilePath("TransformerTest_WriteTabular.xlsx");
            if (File.Exists(newFilePath))
                File.Delete(newFilePath);
            File.Copy(filePath, newFilePath);
            // Load entities from old file
            var transformer = new Giga.Transformer.Transformer(cfg);
            var entities = transformer.Load<TestTabularData>(filePath, "TestNormalTabularData_DynamicRange").ToList();
            foreach (TestTabularData entity in entities)
            {
                entity.DueDate += new TimeSpan(1, 0, 0, 0);
                entity.Item += 1;
                entity.PO += "_new";
                entity.PODate += new TimeSpan(1, 0, 0, 0);
                entity.ProductCode += "_new";
                entity.ProductName += "_new";
                entity.Qty += 1;
                entity.Total += 1;
                entity.UnitPrice += 1;
                entity.Weight += 1;
            }
            // Write entities to new file
            transformer.Save(newFilePath, "TestNormalTabularData_DynamicRange", entities);
        }

        [TestMethod]
        public void TestLoadEmbededDataFromFile()
        {
            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            TestEmbededData order = TestEmbededData.Load(filePath);
            Assert.IsNotNull(order);
            Console.Write(order.ToString());
        }

        [TestMethod]
        public void TestWriteEmbededDataToFile()
        {
            // Get test file
            var filePath = GetTestFilePath("TransformerTest.xlsx");
            if (!File.Exists(filePath))
                throw new FileNotFoundException(String.Format("Test file {0} not found!", filePath));
            // Create a new file 
            var newFilePath = GetTestFilePath("TransformerTest_WriteEmbeded.xlsx");
            if (File.Exists(newFilePath))
                File.Delete(newFilePath);
            File.Copy(filePath, newFilePath);
            TestEmbededData order = TestEmbededData.Load(filePath);
            Assert.IsNotNull(order);
            Console.Write(order.ToString());
            order.Currency += "_new";
            order.Delivery += "_new";
            order.ExtPo += "_new";
            order.Id += "_new";
            order.IncoTerms += "_new";
            order.Payment += "_new";
            order.PoDate += new TimeSpan(1, 0, 0, 0);
            order.ShipTo += "_new";
            order.Version += "_new";
            order.AdditionalNotes += "_new";
            foreach (var item in order.Items)
            {
                item.Description += "_new";
                item.ExwDate += new TimeSpan(1, 0, 0, 0);
                item.LineNumber += 1;
                item.Part += "_new";
                item.Qty += 1;
                item.TotalPrice += 1;
                item.UnitPrice += 1;
            }
            order.Save(newFilePath);
        }

        [TestMethod]
        public void TestCopyWorksheet()
        {
            // Get test file
            var srcFile = GetTestFilePath("CopySource.xlsx");
            var mergeFile = GetTestFilePath("CopyTarget.xlsx");
            var tgtFile = GetTestFilePath("CopyTarget_Result.xlsx");
            if (File.Exists(tgtFile))
                File.Delete(tgtFile);
            File.Copy(mergeFile, tgtFile);
            ExcelUtiles.CopyWorksheet(srcFile, "Non open报价", tgtFile, 0);
        }
    }
}
