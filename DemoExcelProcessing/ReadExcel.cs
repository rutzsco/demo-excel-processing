using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

public class ReadExcel
{
    public static void ParseDevices()
    {

        byte[] bytes = File.ReadAllBytes("C:\\SampleData\\SampleExcel.xlsx");
        Console.WriteLine($"Started on {System.DateTime.Now} ");
        var devices = new List<DeviceDto>();
        using (var mem = new MemoryStream())
        {
            mem.Write(bytes, 0, bytes.Length);
            mem.Seek(0, SeekOrigin.Begin);

            using (var doc = SpreadsheetDocument.Open(mem, false))
            {
                var workbookPart = doc.WorkbookPart;
                var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                var sst = sharedStringTablePart.SharedStringTable;

                var deviceId = workbookPart.Workbook.Descendants<Sheet>().Single(s => s.Name == "Devices").Id;
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(deviceId);
                var sheet = worksheetPart.Worksheet;
                var rows = sheet.Descendants<Row>().Where(x => !string.IsNullOrEmpty(x.InnerText)).ToArray();


                for (var rowCount = 1; rowCount < rows.Length; rowCount++)
                {
                    var deviceDto = new DeviceDto();
                    var rowCells = rows[rowCount].Elements<Cell>().ToArray();
                    deviceDto.DeviceName = GetCellValue(rowCells, "A", sst);
                    var Id = GetCellValue(rowCells, "B", sst);
                    deviceDto.Id = int.TryParse(Id, out _) ? Convert.ToInt32(Id) : default(int);
                    deviceDto.DeviceId = GetCellValue(rowCells, "C", sst);
                    deviceDto.DeviceType = GetCellValue(rowCells, "D", sst);
                    deviceDto.MappedSpaceName = GetCellValue(rowCells, "E", sst);
                    var mappedSpaceId = GetCellValue(rowCells, "F", sst);
                    deviceDto.MappedSpaceId = int.TryParse(mappedSpaceId, out _) ? Convert.ToInt32(mappedSpaceId) : default(int?);
                    deviceDto.IpAddress = GetCellValue(rowCells, "G", sst);

                    devices.Add(deviceDto);
                }
            }

            Console.WriteLine($"Processed records count = {devices.Count} ");
        }

        Console.WriteLine($"Completed on {System.DateTime.Now} ");
        Console.ReadLine();
    }


    private static string GetCellValue(IEnumerable<Cell> rowCells, string columnReference, SharedStringTable sharedStringTable)
    {
        var cell = rowCells.FirstOrDefault(x => Regex.Replace(x.CellReference, "[0-9]", "") == columnReference);
        if (cell?.CellValue == null)
        {
            return String.Empty;
        }

        if (cell.DataType != null && cell.DataType == "s")
        {
            var ssid = int.Parse(cell.CellValue.Text);
            return sharedStringTable.ChildElements[ssid].InnerText;
        }

        return cell.CellValue.Text;
    }
}
public class DeviceDto
{
    public int Id { get; set; }
    public string DeviceId { get; set; }
    public string DeviceName { get; set; }
    public string DeviceType { get; set; }
    public int? MappedSpaceId { get; set; }
    public string MappedSpaceName { get; set; }
    public string IpAddress { get; set; }
}