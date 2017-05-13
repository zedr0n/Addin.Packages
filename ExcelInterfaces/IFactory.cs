namespace ExcelInterfaces
{
    public interface IFactory<out TPublic> where TPublic : class
    {
        TPublic Bind(string name, object instance);
    }
}