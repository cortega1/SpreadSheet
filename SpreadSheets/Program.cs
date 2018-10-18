using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadSheets
{
    class Program
    {
        //Starts in 4 because these values are the headers of the table.
        static int stringTableIndex = 4;
        static uint rowIndex = default(uint);

        //Headers of the table
        static List<string> headers = new List<string>
        {
            "First Name",
            "Last Name",
            "Unit",
            "Building",
            "Has Authorize Entry"
        };

        static List<User> users = new List<User>();

        static void Main(string[] args)
        {
            FillUsersObject();
            SpreadsheetDocument document = SpreadsheetDocument.Create("Document.xlsx", SpreadsheetDocumentType.Workbook);
            WorkbookPart wbPart = document.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet sheet = new Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Example 1"
            };

            sheets.Append(sheet);
            wbPart.Workbook.Save();

            Worksheet worksheet = wsPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            SharedStringTablePart sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable();

            AddHeaders(sharedStringPart);

            Row firstRow = new Row
            {
                RowIndex = rowIndex
            };
            
            sheetData.Append(firstRow);
            rowIndex++;

            for (int i = 0; i <= stringTableIndex; i++)
            {
                Cell cell = new Cell();
                cell.CellValue = new CellValue(i.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                firstRow.AppendChild<Cell>(cell);
                wsPart.Worksheet.Save();
            }

            foreach(var user in users)
            {
                Row data = new Row
                {
                    RowIndex = rowIndex
                };

                sheetData.Append(data);
                foreach (var propertyInfo in user.GetType()
                                .GetProperties(
                                        BindingFlags.Public
                                        | BindingFlags.Instance))
                {
                    string value = propertyInfo.GetValue(user, null).ToString();
                    AddValuesToSharedString(value, sharedStringPart);
                    Cell cell = new Cell();
                    cell.CellValue = new CellValue(stringTableIndex.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    data.AppendChild<Cell>(cell);
                }

                rowIndex++;

                wsPart.Worksheet.Save();
            }

            document.Close();
        }

        public static void AddValuesToSharedString(string value, SharedStringTablePart sharedStringPart)
        {
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
            sharedStringPart.SharedStringTable.Save();

            stringTableIndex++;
        }

        public static void AddHeaders(SharedStringTablePart sharedStringPart)
        {
            foreach(string header in headers)
            {
                sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(header)));
            }
            
            sharedStringPart.SharedStringTable.Save();
        }

        public static void FillUsersObject()
        {
            bool active = false;
            for (int i = 0; i < 10; i++)
            {
                users.Add(new User
                {
                    GivenName = $"User {i}",
                    LastName = "Ramirez",
                    Unit = $"10{i}",
                    Building = "Lawrence House",
                    EntryAuthorizationEnabled = active
                });

                active = !active;
            }
        }
    }

    public class User
    {
        public string GivenName { get; set; }
        public string LastName { get; set; }
        public string Unit { get; set; }
        public string Building { get; set; }
        public bool EntryAuthorizationEnabled { get; set; }
    }
}
