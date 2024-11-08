using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class ExcelToXmlConverter
{
    public static XDocument ConvertLargeExcelToXml(string filePath)
    {
        var root = new XElement("Workbook");

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            var sheets = workbookPart.Workbook.Sheets;

            foreach (Sheet sheet in sheets)
            {
                var sheetData = new XElement("Sheet", new XAttribute("Name", sheet.Name));
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                using (var reader = OpenXmlReader.Create(worksheetPart))
                {
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Row))
                        {
                            var rowData = new XElement("Row");
                            reader.ReadFirstChild();

                            do
                            {
                                if (reader.ElementType == typeof(Cell))
                                {
                                    Cell cell = (Cell)reader.LoadCurrentElement();
                                    string cellValue = GetCellValue(cell, workbookPart);
                                    rowData.Add(new XElement("Cell", cellValue));
                                }
                            }
                            while (reader.ReadNextSibling());

                            sheetData.Add(rowData);
                        }
                    }
                }

                root.Add(sheetData);
            }
        }

        return new XDocument(root);
    }

    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        string value = cell.InnerText;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTable != null)
            {
                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }

        return value;
    }
}
