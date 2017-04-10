using System;
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
    }
}