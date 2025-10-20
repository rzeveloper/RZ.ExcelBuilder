using RZ.ExcelBuilder.Core;
using RZ.ExcelBuilder.Test.Models;

public class Program
{
    static void Main()
    {
        var random = new Random();

        string[] summaries = ["Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"];

        List<WeatherForecast> data = [.. Enumerable.Range(1, 5).Select(index => new WeatherForecast
        {
            Rut = "60567492",
            Date = DateTime.Now.AddDays(index),
            TemperatureC = random.Next(-20, 55),
            Summary = summaries[random.Next(summaries.Length)]
        })];

        string[] headers = ["RUT", "Fecha", "Temperatura °C", "Temperatura °F", "Resumen"];

        List<ColumnBuilder> columns = [
            new("Rut", ColumnType.RUT),
            new("Date", ColumnType.DATE),
            new("TemperatureC", ColumnType.INTEGER),
            new("TemperatureF", ColumnType.DECIMAL),
            new("Summary", ColumnType.STRING)
        ];

        string base64 = ExcelBuilder.BuildExcel("Reporte Clima", headers, columns, data);
        File.WriteAllBytes("reporte.xlsx", Convert.FromBase64String(base64));

        Console.WriteLine("Reporte generado correctamente!");
    }
}