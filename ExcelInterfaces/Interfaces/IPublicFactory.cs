using IoC;

namespace ExcelInterfaces
{
    public interface IPublicFactory<out TPublic> : IInjectable where TPublic : class 
    {
        TPublic Bind(string name, object instance);
    }
}