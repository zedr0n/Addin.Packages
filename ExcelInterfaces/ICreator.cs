using System;
using IoC;

namespace ExcelInterfaces
{
    public interface ICreator : IInjectable
    {
        IPublicObject Get(string handle);
        IPublicObject Get(string handle, Type publicType);
        IPublicObject Create<TPublic, TInstance>(string handle, TInstance instance) where TPublic : Public<TInstance>
                                                                                    where TInstance : class;
    }
}