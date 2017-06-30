using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace ExcelInterfaces
{
    public interface IExcelServices
    {
        TProperty AddBinding<T, TProperty>(T obj, Expression<Func<T, TProperty>> memberLambda, Func<object,TProperty> converter = null, BINDING_TYPE bindingType = BINDING_TYPE.ONE_WAY) where T : class, INotifyPropertyChanged;
        bool RegisterButton(string buttonName, string functionName, object instance);

    }
}