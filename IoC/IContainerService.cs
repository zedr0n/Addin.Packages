using System;

namespace IoC
{
    public interface IContainerService
    {
        object GetInstance(Type type);
        T GetInstance<T>() where T : class;
    }
}