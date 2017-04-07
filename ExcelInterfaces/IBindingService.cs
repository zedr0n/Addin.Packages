using System;
using System.Linq.Expressions;
using IoC;

namespace ExcelInterfaces
{
    public interface IBindingService : IInjectable
    {
        void AddBinding<T, TProperty>(string cell, T obj, Expression<Func<T, TProperty>> memberLambda) where T : class;
        void OnSheetChange(object sheet, object target);
    }

    public class AddressInfo
    {
        public string Address { get; set; }
    }
}