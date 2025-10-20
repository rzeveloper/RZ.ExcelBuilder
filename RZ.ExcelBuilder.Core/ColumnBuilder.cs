namespace RZ.ExcelBuilder.Core
{
    public class ColumnBuilder(string propertyName, ColumnType columnType, string format = "")
    {
        public string PropertyName { get; set; } = propertyName;
        public ColumnType ColumnType { get; set; } = columnType;
        public string Format { get; set; } = format;
    }

    public enum ColumnType
    {
        STRING = 1,
        INTEGER = 2,
        DECIMAL = 3,
        DATE = 4,
        RUT = 5,
    }
}
