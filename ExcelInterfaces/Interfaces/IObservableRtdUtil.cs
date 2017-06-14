using System;
using System.ComponentModel;
using System.Linq.Expressions;
using IoC;

namespace ExcelInterfaces
{
    public interface IRtdService : IInjectable
    {
        object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource);
        TProperty ObserveProperty<TInstance,TProperty>(string functionName, TInstance instance, string propertyName)
            where TInstance : INotifyPropertyChanged;

        TProperty ObserveProperty<TInstance, TProperty>(string functionName, TInstance instance,
            Expression<Func<TInstance, TProperty>> property)
            where TInstance : INotifyPropertyChanged;
    }
}