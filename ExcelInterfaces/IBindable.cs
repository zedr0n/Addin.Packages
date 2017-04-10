namespace ExcelInterfaces
{
    public interface IBindable
    {
        object Get(string propertyName);
        object Bind(string propertyName);
    }
}