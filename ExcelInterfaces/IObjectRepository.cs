using System;
using IoC;

namespace ExcelInterfaces
{
    public interface IObjectRepository : IInjectable
    {
        IPublicObject Get(string handle);
    }
}