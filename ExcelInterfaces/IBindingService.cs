using System;
using System.Linq.Expressions;
using IoC;

namespace ExcelInterfaces
{
    public interface IBindingService : IInjectable
    {
        void AddBinding<T,TProperty>(string cell, T obj, Expression<Func<T, TProperty>> memberLambda) where T : class;
        /// <summary>
        /// Syncs the property values for the bound objects for changed cell
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="target"></param>
        void OnSheetChange(object sheet, object target);
    }
}