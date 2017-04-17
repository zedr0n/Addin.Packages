using System;
using IoC;

namespace ExcelInterfaces
{
    public interface ICreator : IInjectable
    {
        IPublicObject Get(string handle);
        TPublic Create<TPublic, TInstance>(string handle, TInstance instance) where TPublic : Public<TInstance>
                                                                                    where TInstance : class;
    }

    public interface ICreator<out TPublic> : IInjectable
        where TPublic : IPublicObject
    {
        TPublic Default();
    }
}