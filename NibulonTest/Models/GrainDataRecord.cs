using LinqToExcel.Attributes;
using OfficeOpenXml.Attributes;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace NibulonTest.Models
{
    public class GrainDataRecord
    {   
        [ExcelColumn("Номер _запису")]
        public long Id { get; set; }
        [ExcelColumn("Дата обліку")]
        public DateTime RecordDate { get; set; }
        [ExcelColumn("Підрозділ _Код")]
        public int UnitCode { get; set; }
        [ExcelColumn("Рік врожаю")]
        public int HarvestYear { get; set; }
        [ExcelColumn("Контрагент")]
        public int CounterpartyId { get; set; }
        [ExcelColumn("Найменування")]
        public string Name { get; set; }
        [ExcelColumn("Унікальний номер договору")]
        public long TreatyId { get; set; }
        [ExcelColumn("ТМЦ _Код")]
        public string TMCcode { get; set; }
        [ExcelColumn("Ціна")]
        public long Price { get; set; }
        [ExcelColumn("Кількість _нетто")]
        public long NetQuantity { get; set; }
        [ExcelColumn("Напрямок")]
        public string Direction { get; set; }
        [ExcelColumn("вологість")]
        public decimal? Moisture { get; set; }
        [ExcelColumn("сміття")]
        public decimal? Trash { get; set; }
        [ExcelColumn("зараженість")]
        public string Infection { get; set; }
        [JsonIgnore]
        [EpplusIgnore]
        public decimal InfectionDecimal
        {
            get
            {
                var success = decimal.TryParse(new String(Infection?.Where(x => char.IsDigit(x)).ToArray()), out var value);
                if (success)
                    return value;
                else
                    return 0;
            }
            set
            {
                if (value == 0)
                    this.Infection = "н/обн";
                else
                    this.Infection = value.ToString() + " ст";

            }
        }
    }
}
