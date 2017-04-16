using System;
using IoC;

namespace ExcelInterfaces
{
    public interface ICreator : IInjectable
    {
        IPublicObject Get(string handle);
        IPublicObject Create<TPublic, TInstance>(string handle, TInstance instance) where TPublic : Public<TInstance>
                                                                                    where TInstance : class;
    }

    public interface ICreator<out TPublic> : IInjectable
        where TPublic : IPublicObject
    {
        TPublic Default();
    }
}