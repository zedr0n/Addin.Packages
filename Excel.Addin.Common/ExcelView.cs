using System;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using CommonAddin;
using ExcelInterfaces;
using IoC;

namespace Excel.Addin.Common
{
    public interface IExcelView { }

    public class ExcelView : IExcelView
    {
        private readonly IExcelServices _excelServices;

        protected TProperty Bind<T,TProperty>(T viewModel,Expression<Func<T, TProperty>> property, Func<object,TProperty> converter = null)
            where T : class,INotifyPropertyChanged
        {
            return _excelServices.AddBinding(viewModel, property, converter, BINDING_TYPE.TWO_WAY);
;        }

        protected TProperty Get<T, TProperty>(T viewModel, Expression<Func<T, TProperty>> property)
            where T : class, INotifyPropertyChanged
        {
            return _excelServices.AddBinding(viewModel, property);
        }

        protected ExcelView(IExcelServices excelServices)
        {
            _excelServices = excelServices;
        }

        /*protected string RegisterButton(string btnName, string onClick)
        {
            if (_registrationService.RegisterButton(btnName, onClick, this))
                return btnName + " associated with " + onClick;
            return "";
        }*/

        protected string RegisterButton(string btnName, string methodName)
        {
            var attribute = GetType().GetMethod(methodName).GetCustomAttributes(typeof(ExportAttribute), true).ToList()
                .OfType<ExportAttribute>().FirstOrDefault();

            if (attribute?.Name == null)
                throw new ArgumentException("Method is not exported");

            if (_excelServices.RegisterButton(btnName, attribute.Name, this))
                return btnName + " associated with " + attribute.Name;
            return "";
        }
    }

    public class ExcelView<TViewModel> : ExcelView
        where TViewModel : class, INotifyPropertyChanged
    {
        protected readonly TViewModel _viewModel;

        protected ExcelView(IExcelServices excelServices,TViewModel viewModel) : 
            base(excelServices)
        {
            _viewModel = viewModel;
        }

        /// <summary>
        /// Bind the view model property change to view using RTD
        /// </summary>
        /// <param name="property"></param>
        /// <param name="converter"></param>
        /// <typeparam name="TProperty"></typeparam>
        /// <returns></returns>
        protected TProperty Bind<TProperty>(Expression<Func<TViewModel, TProperty>> property,
            Func<object,TProperty> converter = null) => Bind(_viewModel, property,converter);

        protected TProperty Get<TProperty>(Expression<Func<TViewModel, TProperty>> property) => Get(_viewModel, property);
    }
}