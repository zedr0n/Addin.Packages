using IoC;

namespace ExcelInterfaces
{
    public interface ICreator : IInjectable
    {
        IPublicObject Get(string handle);
    }
}