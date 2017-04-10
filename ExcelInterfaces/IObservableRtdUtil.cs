using System;

namespace ExcelInterfaces
{
    public interface IObservableRtdService
    {
        object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource);
    }
}