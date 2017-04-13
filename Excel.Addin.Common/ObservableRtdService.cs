using System;
using System.ComponentModel;
using ExcelDna.Registration.Utils;
using ExcelInterfaces;

namespace CommonAddin
{
    public class ObservableRtdService : IObservableRtdService
    {
        public object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource)
        {
            return ObservableRtdUtil.Observe(callerFunctionName, callerParameters, observableSource);
        }

        public object ObserveProperty<TInstance,TProperty>(string functionName, TInstance instance, string propertyName)
            where TInstance : INotifyPropertyChanged
        {
            var property = BindExtensions.GetPropertyEx<TInstance, TProperty>(instance, propertyName);
            var value = property.Compile()(instance);
            return Observe(functionName, null,
                () => instance.RxValue(property, value));
        }
    }
}