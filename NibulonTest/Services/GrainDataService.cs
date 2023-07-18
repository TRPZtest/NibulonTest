using LinqToExcel;
using NibulonTest.Models;
using OfficeOpenXml;
using System.Net;
using System.Reflection.Metadata.Ecma335;

namespace NibulonTest.Services
{
    public class GrainDataService
    {
        private readonly string _filePath;

        public GrainDataService() 
        {
            var filePath = "Task.xlsx";
            _filePath = filePath;
        }

        public List<GrainDataRecord> GetGrainDataRecords()
        {
            using var excel = new ExcelQueryFactory(_filePath);

            var test = excel.GetColumnNames("Таблиця_1");

            var inputWorksheet = excel.Worksheet<GrainDataRecord>("Таблиця_1").ToList();

            return inputWorksheet;
        }

        public List<GrainDataRecord> GetGrainDataRecordsById(int id)
        {
            using var excel = new ExcelQueryFactory(_filePath);

            var test = excel.GetColumnNames("Таблиця_1");

            var data = excel.Worksheet<GrainDataRecord>("Таблиця_1").Where(x => x.Id == id).ToList();

            return data; //to not return null we should return list
        }

        public List<GrainDataRecord> GetGrainDataRecordsByDate(DateTime begin, DateTime end)
        {
            using var excel = new ExcelQueryFactory(_filePath);

            var test = excel.GetColumnNames("Таблиця_1");

            var data = excel.Worksheet<GrainDataRecord>("Таблиця_1").Where(x => x.RecordDate >= begin && x.RecordDate <= end);

            return data.ToList(); //to not return null we should return list
        }

        public bool UpdateDataGrainRecord(GrainDataRecord record)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            bool wasChanged = false;
            FileInfo file = new FileInfo(_filePath);
            using ExcelPackage excelPackage = new(file);

            var worksheet = excelPackage.Workbook.Worksheets["Таблиця_1"];

            int rows = worksheet.Dimension.Rows - 1; // 20
            int columns = worksheet.Dimension.Columns; // 7

            var test = worksheet.Cells[1, 2].Value.ToString();

            for (int i = 2; i < rows; i++)
            {
                if (int.TryParse(worksheet.Cells[i, 1].Value.ToString(), out int id) && id == record.Id)
                {
                    int j = 1;

                    worksheet.Cells[i, ++j].Value = record.RecordDate;
                    worksheet.Cells[i, ++j].Value = record.UnitCode;
                    worksheet.Cells[i, ++j].Value = record.HarvestYear;
                    worksheet.Cells[i, ++j].Value = record.CounterpartyId;
                    worksheet.Cells[i, ++j].Value = record.Name;
                    worksheet.Cells[i, ++j].Value = record.TreatyId;
                    worksheet.Cells[i, ++j].Value = record.TMCcode;
                    worksheet.Cells[i, ++j].Value = record.Price;
                    worksheet.Cells[i, ++j].Value = record.NetQuantity;
                    worksheet.Cells[i, ++j].Value = record.Direction;
                    worksheet.Cells[i, ++j].Value = record.Moisture;
                    worksheet.Cells[i, ++j].Value = record.Trash;
                    worksheet.Cells[i, ++j].Value = record.Infection;

                    excelPackage.Save();
                    wasChanged = true;
                    return wasChanged;
                }
            }

            return wasChanged;
        }

        public List<GroupedDataRecord> GetGroupedDataRecords(List<GrainDataRecord> data)
        {
            var groupedData = from record in data
                      group record by new
                      {
                          record.RecordDate,
                          record.CounterpartyId,
                          record.Name,
                          record.TreatyId,
                          record.TMCcode,
                          record.Direction,
                          record.Moisture,
                          record.InfectionDecimal,
                          record.Trash
                      } into gcs
                      select new GroupedDataRecord()
                      {
                          RecordDate = gcs.Key.RecordDate,
                          CounterpartyId = gcs.Key.CounterpartyId,
                          Name = gcs.Key.Name,
                          TreatyId = gcs.Key.TreatyId,
                          TMCcode = gcs.Key.TMCcode,
                          Direction = gcs.Key.Direction,
                          Price = gcs.Average(x => x.Price),
                          NetQuantity = gcs.Sum(x => x.NetQuantity),
                          Moisture = gcs.Key.Moisture,
                          Trash = gcs.Key.Trash,
                          InfectionDecimal = gcs.Key.InfectionDecimal
                      };

            return groupedData.ToList();
        }

        public List<GroupedDataRecord> GetGroupedAvgDataRecords(List<GrainDataRecord> data)
        {
            var avg = from record in data
                      group record by new
                      {
                          record.RecordDate,
                          record.CounterpartyId,
                          record.Name,
                          record.TreatyId,
                          record.TMCcode,
                          record.Direction
                      } into gcs
                      select new GroupedDataRecord()
                      {
                          RecordDate = gcs.Key.RecordDate,
                          CounterpartyId = gcs.Key.CounterpartyId,
                          Name = gcs.Key.Name,
                          TreatyId = gcs.Key.TreatyId,
                          TMCcode = gcs.Key.TMCcode,
                          Direction = gcs.Key.Direction,
                          Price = gcs.Average(x => x.Price),
                          NetQuantity = gcs.Sum(x => x.NetQuantity),
                          Moisture = gcs.Average(x => x.Moisture),
                          Trash = gcs.Average(x => x.Trash),
                          InfectionDecimal = gcs.Average(x => x.InfectionDecimal)
                      };

            return avg.ToList();
        }

        public async Task<byte[]> GetReportFile(DateTime begin, DateTime end)
        {
            var data = GetGrainDataRecordsByDate(begin, end);

            var groupedData = GetGroupedDataRecords(data);

            var groupedAvgData = GetGroupedAvgDataRecords(data);

            using FileStream fileStream = File.OpenRead(_filePath);
            using MemoryStream memoryStream = new MemoryStream();

            await fileStream.CopyToAsync(memoryStream);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using ExcelPackage excelPackage = new ExcelPackage(memoryStream);


            WriteReport(excelPackage, groupedData, "Таблиця_2", 2, 1);

            WriteReport(excelPackage, groupedAvgData, "Таблиця_3", 3, 1);

            return await excelPackage.GetAsByteArrayAsync();
        }

        
     
        private void WriteReport(ExcelPackage excelPackage, List<GroupedDataRecord> records, string sheetName, int x, int y)
        {
            var worksheet = excelPackage.Workbook.Worksheets[sheetName];

            worksheet.Cells[x, y].Clear();

            worksheet.Cells[x, y].LoadFromCollection(records);
            worksheet.Column(1).Style.Numberformat.Format = "MM/dd/yyyy";

            excelPackage.Save();
        }

    }
}
