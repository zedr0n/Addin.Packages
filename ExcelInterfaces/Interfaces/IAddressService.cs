using IoC;

namespace ExcelInterfaces
{
    public interface IAddressService : IInjectable
    {
        string GetAddress();
    }
}