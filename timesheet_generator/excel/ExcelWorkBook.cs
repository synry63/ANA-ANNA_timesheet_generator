using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics.Tracing;
using timesheet_generator.resources;
using DocumentFormat.OpenXml.Vml;

namespace timesheet_generator.excel
{
    internal class ExcelWorkBook
    {

        public static DataTable GetDataTableForMonth(int year, int month, IDictionary<DateTime, EventResourceCalendar> events)
        {
            DataTable table = new DataTable();
            table.Columns.Add("Day", typeof(DateTime));
            table.Columns.Add("WorkAtOffice", typeof(string));
            table.Columns.Add("WorkAtHome", typeof(string));
            table.Columns.Add("NotWorked", typeof(string));
            DateTime startDate = new DateTime(year, month, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Add rows for each day of the month
            for (int day = 1; day <= daysInMonth; day++)
            {
                DataRow row = table.NewRow();
                row["Day"] = startDate.AddDays(day - 1);
                table.Rows.Add(row);
            }

            // Add data for each day of the month
            foreach (DataRow row in table.Rows)
            {
                var thisDate = Convert.ToDateTime(row["Day"]);
                var value = events.FirstOrDefault(x => x.Key == thisDate).Value;

                row["WorkAtOffice"] = (value != null && value.eventType == EventTypeCalendar.Work) ? 1 : "";
                row["WorkAtHome"] = (value != null && value.eventType == EventTypeCalendar.WorkRemote) ? 1 : "";
                row["NotWorked"] = (value == null) ? 1 : "";

            }

            return table;
        }
        public static Row getHeader(out int cpt)
        {
            cpt = 0;
            Row headerRow = new Row();
            headerRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("Date") });
            headerRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("Work") });
            headerRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("Work remote") });
            headerRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("Off") });
            cpt++;
            return headerRow;
        }
        public static Row getFooterCalc(int cptHeaderRow, int cptDataRow)
        {
            int startRow = 1;
            Row calcRow = new Row();
            calcRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("Total") });
            calcRow.Append(new Cell() { DataType = CellValues.Number, CellFormula = new CellFormula($"=SUM(B{startRow + cptHeaderRow}:B{cptDataRow + cptHeaderRow})") });
            calcRow.Append(new Cell() { DataType = CellValues.Number, CellFormula = new CellFormula($"=SUM(C{startRow + cptHeaderRow}:C{cptDataRow + cptHeaderRow})") });
            calcRow.Append(new Cell() { DataType = CellValues.Number, CellFormula = new CellFormula($"=SUM(D{startRow + cptHeaderRow}:D{cptDataRow + cptHeaderRow})") });
            
            return calcRow;
        }
        public static void ExportToExcel(DataTable table, string filePath, DateTime period)
        {
            // Create a new Excel file
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Add a new workbook
                WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Loop through each month and add a new worksheet to the workbook
                int sheetIndex = 1;
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Set the worksheet name to the month name and year (e.g. "January 2023")
                string sheetName = period.ToString("MMMM yyyy");
                UInt32Value? sheet_id = Convert.ToUInt32(sheetIndex);
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = sheet_id, Name = sheetName };
                //spreadsheet.WorkbookPart.Workbook.Append(sheet);

                sheets.Append(sheet);
                // Add data to the worksheet
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();


                // Add the headers to the worksheet
                int cptHeaderRow = 0;
                sheetData.Append(getHeader(out cptHeaderRow));

                // Add data to the worksheet
                int cptDataRow = 0;
                foreach (DataRow row in table.Rows)
                {
                    Row sheetRow = new Row();
                    foreach (DataColumn column in table.Columns)
                    {
                        var objVal = row[column].ToString();

                        var valFormated = Tools.FormatValString(objVal);

                        Cell sheetCell = new Cell() { DataType = valFormated.Keys.First(), CellValue = new CellValue(valFormated.Values.First()) };
                        sheetRow.Append(sheetCell);

                        //////if (int.TryParse(objVal, out intValue))
                        //////{
                        //////    Cell sheetCell = new Cell() { DataType = CellValues.Number, CellValue = new CellValue(intValue) };
                        //////    sheetRow.Append(sheetCell);
                        //////}
                        //////else
                        //////{
                        //////    Cell sheetCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(Tools.FormatValString(objVal)) };
                        //////    sheetRow.Append(sheetCell);
                        //////}

                    }
                    sheetData.Append(sheetRow);
                    cptDataRow++;
                }
                // Add the calc to the worksheet
                sheetData.Append(getFooterCalc(cptHeaderRow, cptDataRow));

                // Save the Worksheet  
                worksheetPart.Worksheet.Save();
                // Save the workbook            
                workbookPart.Workbook.Save();
            }

        }

    }
}
