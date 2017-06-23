using System;
using System.ComponentModel;
using System.Linq.Expressions;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;

namespace CommonAddin
{
    public class RtdService : IRtdService
    {
        public object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource)
        {
            return ObservableRtdUtil.Observe(callerFunctionName, callerParameters, observableSource);
        }

        public TProperty ObserveProperty<TInstance,TProperty>(string functionName, TInstance instance, string propertyName)
            where TInstance : INotifyPropertyChanged
        {
            var property = BindExtensions.GetPropertyEx<TInstance, TProperty>(instance, propertyName);
            return ObserveProperty(functionName,instance,property);
        }

        public TProperty ObserveProperty<TInstance, TProperty>(string functionName, TInstance instance, Expression<Func<TInstance, TProperty>> property) where TInstance : INotifyPropertyChanged
        {
            return (TProperty) Observe(functionName, null, () => instance.RxValue(property));
        }

        public RtdService()
        {
            var application = (Application) ExcelDnaUtil.Application;
            application.RTD.ThrottleInterval = 0;
        }
    }
}