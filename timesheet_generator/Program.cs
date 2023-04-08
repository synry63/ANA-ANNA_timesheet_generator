using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Tracing;
using System.IO;
using timesheet_generator.excel;
using timesheet_generator.google;
using System.Xml;
using Microsoft.Extensions.Configuration;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        int year, month;
        year = int.Parse(args[0]);
        month = int.Parse(args[1]);
        DateTime period = new DateTime(year,month,1);

        // Get the appsettings.json file
        var config = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
        .Build();
        // Create a new instance of GoogleCalendar using the Google Calendar API credentials specified in the appsettings.json file
        var calendar = new GoogleCalendar(config.GetSection("googleCalendar").Value);
        // Retrieve events from the specified calendar for the given year and month
        var events = calendar.GetEventsPeriodForCalendar(config.GetSection("calendarName").Value, year, month);


        Dictionary<DateTime, DataTable> tables = new Dictionary<DateTime, DataTable>();
        // Generate a DataTable containing all events for the given month
        var table = ExcelWorkBook.GetDataTableForMonth(year, month, events);

        // Get the path for the worksheet where we'll export the data to
        string workSheetPath = config.GetSection("workSheetPath").Value;
        // Export the DataTable to an Excel file with the given filename and period (month and year)
        ExcelWorkBook.ExportToExcel(table, $"{workSheetPath}/{period.ToString("MMMM yyyy")}.xlsx", period);
    }
}