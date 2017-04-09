using IoC;

namespace ExcelInterfaces
{
    public interface ICreator : IInjectable
    {
        IPublicObject Create(string handle);
    }
}