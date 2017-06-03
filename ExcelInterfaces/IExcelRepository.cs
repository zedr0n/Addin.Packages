namespace ExcelInterfaces
{
    public interface IExcelRepository
    {
        object GetByHandle(string handle);
        void Add(object obj, string handleName);
        string ResolveHandle(object instance);
    }
}