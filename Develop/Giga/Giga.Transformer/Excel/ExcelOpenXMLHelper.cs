using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Extension helper for accessing Excel data via Open XML SDK
    /// </summary>
    public static class ExcelOpenXMLHelper
    {
        public const String REG_REF = @"(?i)(?<Sheet>.+)!(?<Range>.+)";

        /// <summary>
        /// Get real reference of defined name
        /// </summary>
        /// <param name="doc">Spread sheet document</param>
        /// <param name="name">Defined name</param>
        /// <returns>Real reference</returns>
        public static String GetDefinedName(this SpreadsheetDocument doc, String name)
        {
            DefinedName dn = doc.WorkbookPart.Workbook.Descendants<DefinedName>().FirstOrDefault(a => a.Name == name);
            if (dn == null)
                return null;
            return dn.InnerText;
        }

        /// <summary>
        /// Get a range from excel document
        /// </summary>
        /// <param name="doc">Spread sheet document</param>
        /// <param name="referenceName">Reference name</param>
        /// <returns></returns>
        public static ExcelOpenXMLRange GetRange(this SpreadsheetDocument doc, String referenceName)
        {
            // Get sheet name and range in sheet
            var regex = new Regex(REG_REF);
            var match = regex.Match(referenceName);
            if (!match.Success)
            {   // This reference might be named
                var realRef = GetDefinedName(doc, referenceName);
                match = regex.Match(realRef);
                if (!match.Success)
                    throw new InvalidDataException(String.Format("Defined name refered to invalid cell reference {0}!",
                        realRef));
            }
            // Find the sheet
            var sheetName = match.Groups["Sheet"].Value;
            if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                sheetName = sheetName.Substring(1, sheetName.Length - 2);
            var sheet =
                doc.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.Name == sheetName);
            if (sheet == null)
                throw new InvalidDataException(String.Format("Sheet {0} not exist!", sheetName));
            var sheetPart = doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            if (sheetPart == null)
                throw new InvalidDataException(String.Format("Sheet part {0} not exist!", sheet.SheetId));
            // Find the range
            var range = match.Groups["Range"].Value;
            return new ExcelOpenXMLRange(doc, sheetPart.Worksheet, range);
        }
    }
}
