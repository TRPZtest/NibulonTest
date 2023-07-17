using LinqToExcel.Attributes;
using OfficeOpenXml.Attributes;
using System.ComponentModel;
using System.Text.Json.Serialization;

namespace NibulonTest.Models
{
    public class GroupedDataRecord
    {       
        [DescriptionAttribute("Дата обліку")]
        public DateTime RecordDate { get; set; }          
        [DescriptionAttribute("Контрагент")]
        public int CounterpartyId { get; set; }
        [DescriptionAttribute("Найменування")]
        public string Name { get; set; }
        [DescriptionAttribute("Унікальний номер договору")]
        public long TreatyId { get; set; }
        [DescriptionAttribute("ТМЦ _Код")]
        public string TMCcode { get; set; }
        [DescriptionAttribute("Ціна")]
        public double Price { get; set; }
        [DescriptionAttribute("Кількість _нетто")]
        public long NetQuantity { get; set; }
        [DescriptionAttribute("Напрямок")]
        public string Direction { get; set; }
        [DescriptionAttribute("вологість")]
        public decimal? Moisture { get; set; }
        [DescriptionAttribute("сміття")]
        public decimal? Trash { get; set; }
        [DescriptionAttribute("зараженість")]
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
