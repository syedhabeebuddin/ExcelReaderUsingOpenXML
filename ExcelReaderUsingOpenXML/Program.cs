using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderUsingOpenXML
{
    class Program
    {
        private const string query = "Insert into PPCode (Code,ProductID,CreatedUserID,CreatedDate,IsActive) " +
                                     "Select <PPCCODE>, mcp.CatalogProductID,9876, <DATE>,1 From CatalogProduct mcp " +
                                     "where mcp.code = <PARTID>";

        static void Main(string[] args)
        {
            //SpreadsheetDocument spreadsheetDocument = new SpreadsheetDocument();
            string fileName = @"E:\Learning\ExcelReaderUsingOpenXML\ExcelReaderUsingOpenXML\bin\Debug\Med Spec UPC List 2019.xlsx";
            //GetExcelData(fileName);
            fileName = "Med Spec UPC List 2019.xlsx";
            Update(fileName);
            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            //{
            //    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            //    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            //    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            //    string text;
            //    while (reader.Read())
            //    {
            //        //if (reader.ElementType == typeof(CellValue))
            //        if (reader.ElementType == typeof(Cell))
            //        {
            //            text = reader.G;
            //            Console.WriteLine(text);
            //            //Console.Read();
            //        }
            //    }

            //}

            List<String> categories;
            List<String> companies;
            ExtractCategoriesCompanies(fileName, out categories, out companies);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                string value;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.InnerText;
                        Console.Write(text + " ");

                        //var value = theCell.CellValue.InnerText;
                        if (c.DataType != null)
                        {
                            switch (c.DataType.Value)
                            {
                                case CellValues.SharedString:
                                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                    if (stringTable != null)
                                    {
                                        value = stringTable.SharedStringTable.ElementAt(int.Parse(text)).InnerText;
                                    }
                                    break;

                                case CellValues.Boolean:
                                    switch (text)
                                    {
                                        case "0":
                                            value = "FALSE";
                                            break;
                                        default:
                                            value = "TRUE";
                                            break;
                                    }
                                    break;
                            }
                        }
                    
                }
                }
                Console.WriteLine();
                //Console.ReadKey();
            }

            Console.Read();




        }

        private static void Update(string fileName)
        {
            List<string> lst = new List<string>();
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            // Look for the first row used
            var firstRowUsed = ws.FirstRowUsed();

            // Narrow down the row so that it only includes the used part
            var row = firstRowUsed.RowUsed();

            // Move to the next row (it now has the titles)
            row = row.RowBelow();

            string upcCode = string.Empty;
            string partID = string.Empty;
            string insertQuery = string.Empty;

            while (!row.Cell(1).IsEmpty())
            {
                partID = row.Cell(1).GetString();
                upcCode = row.Cell(3).GetString();

                insertQuery = query.Replace("<UPCCODE>", upcCode).Replace("<PARTID>", partID).Replace("<DATE>", DateTime.Now.ToString());

                //lst.Add(upcCode);

                row = row.RowBelow();
            }


        }

        private static void ExtractCategoriesCompanies(string northwinddataXlsx, out List<string> categories, out List<string> companies)
        {
            categories = new List<string>();
            const int coCategoryId = 1;
            const int coCategoryName = 2;

            var wb = new XLWorkbook(northwinddataXlsx);
            var ws = wb.Worksheet(1);

            // Look for the first row used
            var firstRowUsed = ws.FirstRowUsed();

            // Narrow down the row so that it only includes the used part
            var categoryRow = firstRowUsed.RowUsed();

            // Move to the next row (it now has the titles)
            categoryRow = categoryRow.RowBelow();

            //categoryRow.FirstCellUsed().CellLeft().

            // Get all categories
            while (!categoryRow.Cell(coCategoryId).IsEmpty())
            {
                String categoryName = categoryRow.Cell(coCategoryName).GetString();
                categories.Add(categoryName);

                categoryRow = categoryRow.RowBelow();
            }

            // There are many ways to get the company table.
            // Here we're using a straightforward method.
            // Another way would be to find the first row in the company table
            // by looping while row.IsEmpty()

            // First possible address of the company table:
            var firstPossibleAddress = ws.Row(categoryRow.RowNumber()).FirstCell().Address;
            // Last possible address of the company table:
            var lastPossibleAddress = ws.LastCellUsed().Address;

            // Get a range with the remainder of the worksheet data (the range used)
            var companyRange = ws.Range(firstPossibleAddress, lastPossibleAddress).RangeUsed();

            // Treat the range as a table (to be able to use the column names)
            var companyTable = companyRange.AsTable();

            // Get the list of company names
            companies = companyTable.DataRange.Rows()
              .Select(companyRow => companyRow.Field("Company Name").GetString())
              .ToList();
        }

        private static DataTable GetExcelData(string filePath)
        {
            var excelData = new DataTable();
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES\";";

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "TABLE" });
                DataTable ss = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable dtCols = conn.GetSchema("Columns");

                var sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                var adapter = new OleDbDataAdapter(String.Format("SELECT * FROM [{0}]", sheet1), conn);

                adapter.Fill(excelData);
            }

            var totalRecords = ((excelData == null || excelData.Rows == null) ? 0 : excelData.Rows.Count);
            //ConsoleLog($"Total rows in excel: {totalRecords}");

            return excelData;
        }
    }
}
