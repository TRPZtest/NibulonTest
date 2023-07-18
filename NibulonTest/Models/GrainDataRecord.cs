using LinqToExcel.Attributes;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace NibulonTest.Models
{
    public class GrainDataRecord
    {   
        [ExcelColumn("Номер _запису")]
        [DisplayName("Номер _запису")]
        public int Id { get; set; }
        [ExcelColumn("Дата обліку")]
        [DisplayName("Дата обліку")]
        public DateTime RecordDate { get; set; }
        [ExcelColumn("Підрозділ _Код")]
        [DisplayName("Підрозділ _Код")]
        public int UnitCode { get; set; }
        [ExcelColumn("Рік врожаю")]
        [DisplayName("Рік врожаю")]
        public int HarvestYear { get; set; }
        [ExcelColumn("Контрагент")]
        [DisplayName("Контрагент")]
        public int CounterpartyId { get; set; }
        [ExcelColumn("Найменування")]
        [DisplayName("Найменування")]
        public string Name { get; set; }
        [ExcelColumn("Унікальний номер договору")]
        [DisplayName("Унікальний номер договору")]
        public long TreatyId { get; set; }
        [ExcelColumn("ТМЦ _Код")]
        [DisplayName("ТМЦ _Код")]
        public string TMCcode { get; set; }
        [ExcelColumn("Ціна")]
        [DisplayName("Ціна")]
        public long Price { get; set; }
        [ExcelColumn("Кількість _нетто")]
        [DisplayName("Кількість _нетто")]
        public long NetQuantity { get; set; }
        [ExcelColumn("Напрямок")]
        [DisplayName("Напрямок")]
        public string Direction { get; set; }
        [ExcelColumn("вологість")]
        [DisplayName("вологість")]
        public decimal? Moisture { get; set; }
        [ExcelColumn("сміття")]
        [DisplayName("сміття")]
        public decimal? Trash { get; set; }
        [ExcelColumn("зараженість")]
        [DisplayName("зараженість")]
        public string Infection { get; set; }
        [JsonIgnore]       
        public decimal InfectionDecimal
        {
            get
            {
                var success = decimal.TryParse(Infection?.Replace("ст", ""), out var value);
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
