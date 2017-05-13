using IoC;

namespace ExcelInterfaces
{
    public interface IFactory<out TPublic> : IInjectable where TPublic : class 
    {
        TPublic Bind(string name, object instance);
    }
}