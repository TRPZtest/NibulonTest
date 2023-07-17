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

        public GrainDataService(string filePath) 
        {
            filePath = "Task.xlsx";
            _filePath = filePath;
        }
        public List<GrainDataRecord> GetGrainDataRecords()
        {
            var excel = new ExcelQueryFactory(_filePath);

            var test = excel.GetColumnNames("Таблиця_1");

            var inputWorksheet = excel.Worksheet<GrainDataRecord>("Таблиця_1").ToList();

            return inputWorksheet;
        }

        public bool UpdateDataGrainRecord(GrainDataRecord record)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            bool wasChanged = false;
            FileInfo file = new FileInfo(_filePath);
            ExcelPackage excelPackage = new(file);

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

        public void WriteGroupedDataReport(List<GroupedDataRecord> records)
        {
            WriteReport(records, "Таблиця_2", 2, 1);      
        }

        public void WriteGroupedAvgDataReport(List<GroupedDataRecord> records)
        {
            WriteReport(records, "Таблиця_3", 3, 1);           
        }

        private void WriteReport(List<GroupedDataRecord> records, string sheetName, int x, int y)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(_filePath);
            using ExcelPackage excelPackage = new ExcelPackage(file);

            var worksheet = excelPackage.Workbook.Worksheets[sheetName];

            worksheet.Cells[x, y].Clear();

            worksheet.Cells[x, y].LoadFromCollection(records);
            worksheet.Column(1).Style.Numberformat.Format = "MM/dd/yyyy";

            excelPackage.Save();
        }

    }
}
