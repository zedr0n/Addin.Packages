using System;
using System.ComponentModel;
using System.Linq.Expressions;
using ExcelInterfaces;

namespace CommonAddin
{
    public class ExcelServices : IExcelServices
    {
        private readonly IBindingService _bindingService;
        private readonly IRegistrationService _registrationService;

        public ExcelServices(IBindingService bindingService, IRegistrationService registrationService)
        {
            _bindingService = bindingService;
            _registrationService = registrationService;
        }

        public TProperty AddBinding<T, TProperty>(T obj, Expression<Func<T, TProperty>> memberLambda, Func<object,TProperty> converter = null, BINDING_TYPE bindingType = BINDING_TYPE.ONE_WAY) where T : class, INotifyPropertyChanged
        {
            return _bindingService.AddBinding(obj, memberLambda,converter, bindingType);
        }

        public bool RegisterButton(string buttonName, string functionName, object instance)
        {
            return _registrationService.RegisterButton(buttonName, functionName, instance);
        }
    }
}