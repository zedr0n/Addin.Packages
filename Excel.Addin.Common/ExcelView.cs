using System;
using System.ComponentModel;
using System.Linq.Expressions;
using CommonAddin;
using ExcelInterfaces;
using IoC;

namespace Excel.Addin.Common
{
    public interface IExcelView { }

    public class ExcelView : IExcelView
    {
        protected readonly IBindingService _bindingService;
        protected readonly IRtdService _rtdService;

        protected TProperty Bind<T,TProperty>(T viewModel,Expression<Func<T, TProperty>> property)
            where T : class,INotifyPropertyChanged
        {
            _bindingService.AddBinding(viewModel, property);
            return Get(viewModel, property);

        }

        protected TProperty Get<T, TProperty>(T viewModel, Expression<Func<T, TProperty>> property)
            where T : class, INotifyPropertyChanged
        {
            return _rtdService.ObserveProperty(nameof(Bind) + property.GetPropertyInfo().Name + "." + viewModel.GetHashCode(), viewModel, property);
        }

        protected ExcelView(IBindingService bindingService, IRtdService rtdService)
        {
            _bindingService = bindingService;
            _rtdService = rtdService;
        }
    }

    public class ExcelView<TViewModel> : ExcelView
        where TViewModel : class, INotifyPropertyChanged
    {
        protected readonly TViewModel _viewModel;

        protected ExcelView(IBindingService bindingService, IRtdService rtdService, TViewModel viewModel) : base(bindingService, rtdService)
        {
            _viewModel = viewModel;
        }

        /// <summary>
        /// Bind the view model property change to view using RTD
        /// </summary>
        /// <param name="property"></param>
        /// <typeparam name="TProperty"></typeparam>
        /// <returns></returns>
        protected TProperty Bind<TProperty>(Expression<Func<TViewModel, TProperty>> property) => Bind(_viewModel, property);

        protected TProperty Get<TProperty>(Expression<Func<TViewModel, TProperty>> property) => Get(_viewModel, property);
    }
}