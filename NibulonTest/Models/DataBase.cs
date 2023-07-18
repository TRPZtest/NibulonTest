namespace NibulonTest.Models
{
    public class DataBase
    {
        public string Infection { get; set; }
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
