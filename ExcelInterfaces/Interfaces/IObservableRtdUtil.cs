using System;
using System.ComponentModel;
using IoC;

namespace ExcelInterfaces
{
    public interface IObservableRtdService : IInjectable
    {
        object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource);
        object ObserveProperty<TInstance,TProperty>(string functionName, TInstance instance, string propertyName)
            where TInstance : INotifyPropertyChanged;
    }
}