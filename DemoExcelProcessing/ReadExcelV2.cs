using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Collections;
using DocumentFormat.OpenXml.Drawing.Charts;

public class ReadExcelV2
{
    private List<string> sharedStrings;

    public void Parse(string path)
    {
        var deviceList = new List<DeviceDto>();
        Console.WriteLine($"Started on {System.DateTime.Now} ");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
        {
            WorkbookPart workbookPart = doc.WorkbookPart;
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            // Cache the shared strings
            sharedStrings = new List<string>();
            if (workbookPart.SharedStringTablePart != null)
            {
                foreach (var item in workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
                {
                    sharedStrings.Add(item.InnerText);
                }
            }

            int sheetCount = 0;
            foreach (Sheet sheet in sheets)
            {
                sheetCount++;
                if (sheetCount == 2) // The second sheet
                {
                    string relationshipId = sheet.Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;

                    foreach (Row row in workSheet.GetFirstChild<SheetData>())
                    {
                        var values = new string[10];
                        int index = 0;
                        foreach (Cell cell in row)
                        {
                            string cellValue = GetCellValue(cell); 
                            values[index]= cellValue;
                            index++;
                            // TODO Parse into Object
                            //Console.Write(cellValue);
                            //Console.Write(",");
                        }

                        var deviceDto = new DeviceDto();
                        deviceDto.DeviceName = values[0];
                        deviceDto.Id = TryConvertToInt(values[1]);
                        deviceDto.DeviceId = values[2];
                        deviceDto.DeviceType = values[3];
                        deviceDto.MappedSpaceName = values[4];
                        deviceDto.MappedSpaceId = TryConvertToInt(values[5]);
                        deviceDto.IpAddress = values[6];
                        deviceList.Add(deviceDto);
                    }

                    break; // No need to continue looping
                }
            }
        }
        Console.WriteLine($"Processed records count = {deviceList.Count} ");
        Console.WriteLine($"Completed on {System.DateTime.Now} ");
    }
    private int TryConvertToInt(string value)
    {
        if (value == "Unassigned")
            return -1;

        try 
        {
            return Convert.ToInt32(value);
        }
        catch (Exception)
        {
            return -1;
        }
    }
    private string GetCellValue(Cell cell)
    {
        if (cell.CellValue == null)
            return "UNKNOWN";

        string value = cell.CellValue.InnerXml;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return sharedStrings[Int32.Parse(value)];
        }
        else
        {
            return value;
        }
    }
}