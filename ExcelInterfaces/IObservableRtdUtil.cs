using System;
using System.ComponentModel;

namespace ExcelInterfaces
{
    public interface IObservableRtdService
    {
        object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource);
        object ObserveProperty<TInstance,TProperty>(string functionName, TInstance instance, string propertyName)
            where TInstance : INotifyPropertyChanged;
    }
}