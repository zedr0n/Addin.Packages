using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private readonly Func<Container, string, IPublicObject> _factory;

        public Registration(Container container, IEnumerable<MethodInfo> methods, Func<Container, string, IPublicObject> factory )
        {
            _container = container;
            _methods = methods;
            _factory = factory;
        }

        private LambdaExpression ConvertToStatic(MethodInfo method, List<Expression> arguments, List<ParameterExpression> callArguments)
        {
            Debug.Assert(method.DeclaringType != null, "method.DeclaringType != null");

            Expression<Func<string, IPublicObject>> expr = h => _factory(_container, h); //h => Public.This(h);

            var instanceParam = Expression.Parameter(method.DeclaringType, "instance");
            var handleParam = Expression.Parameter(typeof(string), "handle");

            var block = Expression.Block(
                new[] {instanceParam},
                Expression.Assign(instanceParam,Expression.Convert(Expression.Invoke(expr, handleParam),method.DeclaringType)),
                Expression.Call(instanceParam, method, arguments)
            );

            var allArguments = new List<ParameterExpression>(callArguments);
            allArguments.Insert(0, handleParam);

            var callExpr = Expression.Lambda(block, allArguments);
            return callExpr;
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
                .Select(e => e as ParameterExpression).ToList();

            if (!method.IsStatic)
                return ConvertToStatic(method, allParams,callParams);

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
                var parameters = methodInfo.GetParameters()
                    .Where(p => !p.ParameterType.GetInterfaces().Contains(typeof(IInjectable)))
                    .Select(p => new ExcelParameterRegistration(p)).ToList();
                var name = methodInfo.Name;
                if (!methodInfo.IsStatic)
                {
                    // add handle as a string parameter to resolve the instance
                    parameters.Insert(0, new ExcelParameterRegistration(
                        new ExcelArgumentAttribute()
                        {
                            AllowReference = true,
                            Description = "Object handle",
                            Name = "Handle"
                        }));
                    if (!methodInfo.DeclaringType.IsGenericType)
                        name = methodInfo.DeclaringType.BaseType.GenericTypeArguments.First().Name + "." + name;
                    else
                        // Generic base functions use the underlying type
                        name = methodInfo.DeclaringType.GenericTypeArguments.First().Name + "." + name;
                }

                var registration = new ExcelFunctionRegistration(lambda,attribute.ToExcelFunctionAttribute(name), parameters);
                registrationList.Add(registration);
            }
            return registrationList;
        }

    }
}
