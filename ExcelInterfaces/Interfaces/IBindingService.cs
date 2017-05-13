using System;
using System.Linq.Expressions;
using IoC;

namespace ExcelInterfaces
{
    public interface IBindingService : IInjectable
    {
        IAddressService AddressService { get; set; }
        TProperty AddBinding<T,TProperty>(T obj, string propertyName) where T : class;
        void AddBinding<T, TProperty>(T obj, Expression<Func<T, TProperty>> memberLambda) where T : class;
        /// <summary>
        /// Syncs the property values for the bound objects for changed cell
        /// </summary>
        /// <param name="sheet">Excel worksheet object</param>
        /// <param name="target">Excel target range</param>
        void OnSheetChange(object sheet, object target);
    }
}