namespace ExcelInterfaces
{
    public interface IPublicObject
    {
        string Handle { get; set; }
        string Type { get; set; }
        object Object { get; }
    }
}