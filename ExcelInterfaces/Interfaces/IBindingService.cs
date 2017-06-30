using System;
using System.ComponentModel;
using System.Linq.Expressions;
using IoC;

namespace ExcelInterfaces
{
    public enum BINDING_TYPE
    {
        ONE_WAY,
        TWO_WAY
    }

    public interface IBindingService : IInjectable
    {


        IAddressService AddressService { get; set; }
        TProperty AddBinding<T,TProperty>(T obj, string propertyName) where T : class, INotifyPropertyChanged;
        TProperty AddBinding<T, TProperty>(T obj, Expression<Func<T, TProperty>> memberLambda,Func<object,TProperty> converter = null ,  BINDING_TYPE bindingType = BINDING_TYPE.ONE_WAY) where T : class, INotifyPropertyChanged;
        /// <summary>
        /// Syncs the property values for the bound objects for changed cell
        /// </summary>
        /// <param name="sheet">Excel worksheet object</param>
        /// <param name="target">Excel target range</param>
        void OnSheetChange(object sheet, object target);
    }
}