using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using IoC;

using ExcelInterfaces;
using SimpleInjector;

namespace CommonAddin
{
    public static class AttributeExtension
    {
        public static ExcelFunctionAttribute ToExcelFunctionAttribute(this IExcelFunctionAttribute attribute, string name)
        {
            return new ExcelFunctionAttribute()
            {
                Category = attribute.Category,
                Name = name,
                Description = attribute.Description,
                ExplicitRegistration =  attribute.ExplicitRegistration,
                HelpTopic = attribute.HelpTopic,
                IsClusterSafe =  attribute.IsClusterSafe,
                IsExceptionSafe = attribute.IsExceptionSafe,
                IsHidden = attribute.IsHidden,
                IsMacroType = attribute.IsMacroType,
                IsThreadSafe = attribute.IsThreadSafe,
                IsVolatile = attribute.IsVolatile   
            };
        }
    }

    public class Registration
    {
        private readonly Container _container;
        private readonly IEnumerable<MethodInfo> _methods;

        public Registration(Container container, IEnumerable<MethodInfo> methods)
        {
            _container = container;
            _methods = methods;
        }

        private LambdaExpression WrapMethod(MethodInfo method)
        {
            if (method.DeclaringType == null)
                return null;

            var allParams = method.GetParameters().Select<ParameterInfo, Expression>(
                p =>
                {
                    if (p.ParameterType.GetInterfaces().Contains(typeof(IInjectable)))
                        return Expression.Constant(_container.GetInstance(p.ParameterType));
                    return Expression.Parameter(p.ParameterType, p.Name);
                }).ToList();

            var callParams = allParams.Where(p => p.NodeType == ExpressionType.Parameter)
                .Select(e => e as ParameterExpression)
                .ToArray();

            var callExpr = Expression.Call(method, allParams);

            return Expression.Lambda(callExpr, method.Name, callParams);
        }

        // register all functions with IExcel attributes
        public IEnumerable<ExcelFunctionRegistration> GetAllRegistrations()
        {
            var registrationList = new List<ExcelFunctionRegistration>();
            foreach (var methodInfo in _methods)
            {
                var lambda = WrapMethod(methodInfo);
                var attribute = (IExcelFunctionAttribute)methodInfo.GetCustomAttributes(typeof(IExcelFunctionAttribute)).Single();
                registrationList.Add(new ExcelFunctionRegistration(lambda, attribute.ToExcelFunctionAttribute(methodInfo.Name), methodInfo.GetParameters()
                    .Where(p => !p.ParameterType.GetInterfaces().Contains(typeof(IInjectable)))
                    .Select(p => new ExcelParameterRegistration(p))));
            }
            return registrationList;
        }

    }
}
