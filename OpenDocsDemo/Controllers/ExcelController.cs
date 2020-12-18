using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using OpenDocsDemo.Models;
using OpenDocsDemo.Helpers;

using Microsoft.AspNetCore.Mvc;

namespace OpenDocsDemo.Controllers
{
    public class ExcelController : Controller
    {
        string docxMIMEType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateSheet()
        {
            using (var stream = new MemoryStream())
            {
                using (var excelDocument = SpreadsheetDocument.Create(stream,
                    SpreadsheetDocumentType.Workbook, true))
                {
                    var workBookPart = excelDocument.AddWorkbookPart();
                    workBookPart.Workbook = new Workbook();

                    var part = workBookPart.AddNewPart<WorksheetPart>();
                    part.Worksheet = new Worksheet(new SheetData());

                    var sheets = workBookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workBookPart.GetIdOfPart(part),
                        SheetId = 1,
                        Name = "Employees"
                    };

                    var sheetData = part.Worksheet.Elements<SheetData>().First();

                    var row = sheetData.AppendChild(new Row());

                    var header1 = ExcelHelpers.ConstructCell("Name", CellValues.String);
                    row.Append(header1);

                    var header2 = ExcelHelpers.ConstructCell("Age", CellValues.String);
                    row.Append(header2);

                    foreach (var employee in ExcelHelpers.GetEmployeeData())
                    {
                        var dataRow = sheetData.AppendChild(new Row());

                        var cell1 = ExcelHelpers.ConstructCell(employee.Name, CellValues.String);
                        dataRow.Append(cell1);

                        var cell2 = ExcelHelpers.ConstructCell(employee.Age.ToString(), CellValues.Number);
                        dataRow.Append(cell2);
                    }

                    sheets.Append(sheet);
                    workBookPart.Workbook.Save();
                    excelDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType, "Excel Sheet Basic Example.xlsx");
            }
        }

        public IActionResult ListValues()
        {
            var filePath = $@"C:\Users\tony_\Downloads\Employee Data.xlsx";

            var list = new List<Employee>();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                var rows = sheetData.Descendants<Row>();

                foreach (var row in rows)
                {
                    var employee = new Employee();
                    var index = 0;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellValue = string.Empty;

                        if (cell.DataType != null)
                        {
                            if (cell.DataType == CellValues.SharedString)
                            {
                                var id = -1;

                                if (Int32.TryParse(cell.InnerText, out id))
                                {
                                    var item = ExcelHelpers.GetSharedStringItemById(workbookPart, id);

                                    if (item.Text != null)
                                    {
                                        cellValue = item.Text.Text;
                                    }
                                    else if (item.InnerText != null)
                                    {
                                        cellValue = item.InnerText;
                                    }
                                    else if (item.InnerXml != null)
                                    {
                                        cellValue = item.InnerXml;
                                    }
                                }
                            }
                        }
                        else
                        {
                            cellValue = cell.InnerText;
                        }

                        switch (index)
                        {
                            case 0:
                                employee.Name = cellValue;
                                break;
                            case 1:
                                employee.Age = int.Parse(cellValue);
                                break;
                            default:
                                break;
                        }

                        index++;
                    }

                    list.Add(employee);
                }
            }

            return View(list);
        }

        public IActionResult InsertChart()
        {
            using (var stream = new MemoryStream())
            {
                using (var excelDocument = SpreadsheetDocument.Create(stream,
                    SpreadsheetDocumentType.Workbook, true))
                {
                    var workBookPart = excelDocument.AddWorkbookPart();
                    workBookPart.Workbook = new Workbook();

                    var part = workBookPart.AddNewPart<WorksheetPart>();
                    part.Worksheet = new Worksheet(new SheetData());

                    var sheets = workBookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workBookPart.GetIdOfPart(part),
                        SheetId = 1,
                        Name = "Employees"
                    };

                    var employeeData = ExcelHelpers.GetEmployeeData().ToDictionary(x => x.Name, x => x.Age);
                    ExcelHelpers.InsertChartInSpreadsheet(excelDocument, sheet, "Employee Data", employeeData);

                    sheets.Append(sheet);

                    workBookPart.Workbook.Save();
                    excelDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType,
                    "Excel Sheet Chart Example.xlsx");
            }
        }
    }
}
