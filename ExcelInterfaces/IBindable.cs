using System;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelInterfaces
{
    public interface IBindable : IPublicObject 
    {
        IBindingService BindingService { get; }
        IObservableRtdService RtdService { get; }
    }

    public static class BindExtensions
    {
        public static Expression<Func<TInstance, TProperty>> GetPropertyEx<TInstance, TProperty>(TInstance obj, string propertyName)
        {
            var pe = Expression.Parameter(obj.GetType());
            var me = Expression.Property(pe, propertyName);

            return Expression.Lambda<Func<TInstance, TProperty>>(me, pe);
        }

        public static object BindPropertyEx<TInstance, TProperty>(this IBindable This, string propertyName)
            where TInstance : class
        {
            return This.BindingService.AddBinding<TInstance, TProperty>(This.Object as TInstance, propertyName);
        }

        public static object BindProperty(this IBindable This, string propertyName)
        {
            return typeof(ExcelInterfaces.BindExtensions)
                .GetRuntimeMethod(nameof(BindExtensions.BindPropertyEx),new[] { This.GetType(), typeof(string) })
                //.GetRuntimeMethods().Single(x => x.Name == nameof(BindExtensions.BindPropertyEx))
                .MakeGenericMethod(This.Object.GetType(),This.Object.GetType().GetRuntimeProperty(propertyName).PropertyType)
                .Invoke(null, new object[] {This, propertyName});

        }

        public static object GetProperty(this IBindable This, string propertyName)
        {
            return typeof(ExcelInterfaces.BindExtensions)
                .GetRuntimeMethod(nameof(BindExtensions.GetPropertyEx), new[] { This.GetType(), typeof(string) })
                //.GetRuntimeMethods().Single(x => x.Name == nameof(BindExtensions.BindPropertyEx))
                .MakeGenericMethod(This.Object.GetType(), This.Object.GetType().GetRuntimeProperty(propertyName).PropertyType)
                .Invoke(null, new object[] { This, propertyName });

        }

        public static TProperty GetPropertyEx<TInstance, TProperty>(this IBindable This, string propertyName)
            where TInstance : class, INotifyPropertyChanged
        {
            return (TProperty)This.RtdService.ObserveProperty<TInstance, TProperty>(nameof(TInstance) + ".Get" + propertyName, This.Object as TInstance, propertyName);
        }

        /*public static object Get(this IBindable This, string propertyName)
        {
            
        }
        public static TProperty GetProperty<TInstance, TProperty>(this IBindable This, string propertyName)
            where TInstance : class
        {
            var property = GetPropertyEx<TInstance, TProperty>(This.Object as TInstance, propertyName);
            var value = property.Compile()(This.Object as TInstance);
            return (TProperty)This.RtdService.Observe(nameof(TInstance) + ".Get" + propertyName, null,
                () => (This.Object as TInstance).RxValue<TInstance, TProperty>(property, value));
        }

        public static Expression<Func<TInstance, TProperty>> GetPropertyEx<TInstance, TProperty>(TInstance obj, string propertyName)
        {
            var pe = Expression.Parameter(obj.GetType());
            var me = Expression.Property(pe, propertyName);

            return Expression.Lambda<Func<TInstance, TProperty>>(me, pe);
        }*/
    }
}