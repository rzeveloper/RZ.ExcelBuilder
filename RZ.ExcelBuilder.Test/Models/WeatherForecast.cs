namespace RZ.ExcelBuilder.Test.Models
{
    public class WeatherForecast
    {
        public string Rut { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public int TemperatureC { get; set; }
        public decimal TemperatureF => (decimal)(32 + (TemperatureC / 0.5556));
        public string Summary { get; set; } = string.Empty;
    }
}
