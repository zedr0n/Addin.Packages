﻿using System;
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
        private readonly IEnumerable<PropertyInfo> _properties;

        /// <summary>
        ///     Creates Public object from handle using container to resolve dependencies
        /// </summary>
        private readonly ICreator _creator;
        /// <summary>
        /// Expression to create public object from handle
        /// </summary>
        private Expression<Func<string, IPublicObject>> CreatePublic => h => _creator.Get(h);
        //private Expression<Func<IPublicObject>> CreateDefault =

        //private Expression<Func<string, IPublicObject>> CreateDefault => h => _creator.Get();

        public Registration(Container container, IEnumerable<MethodInfo> methods, IEnumerable<PropertyInfo> properties)
        {
            _container = container;
            _methods = methods;
            _creator = _container.GetInstance<ICreator>();
            _properties = properties;
        }

        /// <summary>
        ///     Create static method from member method by invoking factory using Container
        /// </summary>
        /// <param name="method">Test</param>
        /// <param name="arguments"></param>
        /// <param name="callArguments"></param>
        /// <returns></returns>
        // #StaticConversion Create static method from member method by invoking factory using Container
        private LambdaExpression ConvertToStatic(MethodInfo method, List<ParameterExpression> arguments)
        {
            Debug.Assert(method.DeclaringType != null, "method.DeclaringType != null");

            //Expression<Func<string, IPublicObject>> createExpression =;

            var instanceParam = Expression.Parameter(method.DeclaringType, "instance");
            var handleParam = Expression.Parameter(typeof(string), "handle");

            var isPublic = method.ReturnType.GetInterfaces().Contains(typeof(IPublicObject)) && method.ReturnType != typeof(IPublicObject);

            Expression methodCallExpression = Expression.Call(instanceParam, method, arguments);
            if (isPublic)
                methodCallExpression = Expression.Convert(methodCallExpression, typeof(IPublicObject));

            var block = Expression.Block(
                new[] {instanceParam},
                Expression.Assign(instanceParam,Expression.Convert(Expression.Invoke(CreatePublic, handleParam),method.DeclaringType)),
                methodCallExpression
            );

            var allArguments = new List<ParameterExpression>(arguments);
            allArguments.Insert(0, handleParam);

            var callExpr = Expression.Lambda(block, allArguments);
            return callExpr;
        }

        [DebuggerStepThrough]
        private LambdaExpression ConvertFactoryToStatic(MethodInfo method, List<ParameterExpression> arguments)
        {
            return (LambdaExpression) typeof(Registration).GetMethod(nameof(ConvertFactoryToStaticEx))
                .MakeGenericMethod(method.DeclaringType)
                .Invoke(this, new object[] {method, arguments});
        }

        /// <summary>
        ///     Create static method from factory create method by invoking create method from default public object using Container
        /// </summary>
        /// <param name="method">Test</param>
        /// <param name="arguments"></param>
        /// <param name="callArguments"></param>
        /// <returns></returns>
        // #DefaultConversion Create default instance of public object before invoking the method
        public LambdaExpression ConvertFactoryToStaticEx<T>(MethodInfo method, List<ParameterExpression> arguments) where T : IPublicObject
        {
            Debug.Assert(method.DeclaringType != null, "method.DeclaringType != null");

            var creator = _container.GetInstance<ICreator<T>>();

            Expression<Func<IPublicObject>> createExpression = () => creator.Default();

            var instanceParam = Expression.Parameter(method.DeclaringType, "instance");

            var isPublic = method.ReturnType.GetInterfaces().Contains(typeof(IPublicObject)) && method.ReturnType != typeof(IPublicObject);

            Expression methodCallExpression = Expression.Call(instanceParam, method, arguments);
            if (isPublic)
                methodCallExpression = Expression.Convert(methodCallExpression, typeof(IPublicObject));

            var block = Expression.Block(
                new[] { instanceParam },
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(createExpression), method.DeclaringType)),
                methodCallExpression
            );

            var callExpr = Expression.Lambda(block, arguments);
            return callExpr;
        }

        /// <summary>
        /// Wraps the method into a lambda to associate with excel UDF
        /// </summary>
        /// <param name="method"></param>
        /// <returns></returns>
        private LambdaExpression WrapMethod(MethodInfo method, bool isFactory)
        {
            if (method.DeclaringType == null)
                return null;

            var arguments = method.GetParameters()
                .Select<ParameterInfo,Expression>(p => Expression.Parameter(p.ParameterType,p.Name))
                .Select(e => e as ParameterExpression).ToList();

            if (!method.IsStatic)
                return !isFactory ? ConvertToStatic(method, arguments) : ConvertFactoryToStatic(method, arguments);

            var callExpr = Expression.Call(method, arguments);
            return Expression.Lambda(callExpr, method.Name, arguments);
        }

        /// <summary>
        /// Wraps the property into a RTD getter lambda to associate with excel UDF
        /// </summary>
        /// <returns></returns>
        private LambdaExpression GetProperty(PropertyInfo property)
        {
            var instanceParam = Expression.Parameter(property.DeclaringType, "instance");
            var handleParam = Expression.Parameter(typeof(string), "handle");
            var method = typeof(BindExtensions).GetMethod(nameof(BindExtensions.GetProperty));
            var block = Expression.Block(
                new[] { instanceParam },
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(CreatePublic, handleParam), property.DeclaringType)),
                Expression.Call(method,instanceParam,Expression.Constant(property.Name))
            );

            var allArguments = new List<ParameterExpression> { handleParam };

            var callExpr = Expression.Lambda(block, allArguments);
            return callExpr;
        }

        /// <summary>
        /// Wraps the property into a RTD bind lambda to associate with excel UDF
        /// </summary>
        /// <returns></returns>
        private LambdaExpression BindProperty(PropertyInfo property)
        {
            var instanceParam = Expression.Parameter(property.DeclaringType, "instance");
            var handleParam = Expression.Parameter(typeof(string), "handle");
            var method = typeof(BindExtensions).GetMethod(nameof(BindExtensions.BindProperty));

            var block = Expression.Block(
                new[] { instanceParam },
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(CreatePublic, handleParam), property.DeclaringType)),
                Expression.Call(method,instanceParam, Expression.Constant(property.Name))
            );

            var allArguments = new List<ParameterExpression> { handleParam };

            var callExpr = Expression.Lambda(block, allArguments);
            return callExpr;
        }

        /// <summary>
        ///     Export all methods marked with [IExcelFunction] as excel functions
        /// </summary>
        /// <returns></returns>
        public IEnumerable<ExcelFunctionRegistration> GetAllRegistrations()
        {
            var registrationList = new List<ExcelFunctionRegistration>();
            foreach (var methodInfo in _methods)
            {
                var attribute = (IExcelFunctionAttribute)methodInfo.GetCustomAttributes(typeof(IExcelFunctionAttribute)).Single();
                var lambda = WrapMethod(methodInfo, attribute.IsFactory);
                var parameters = methodInfo.GetParameters()
                    .Where(p => !p.ParameterType.GetInterfaces().Contains(typeof(IInjectable)))
                    .Select(p => new ExcelParameterRegistration(p)).ToList();
                var name = methodInfo.Name;
                if (!methodInfo.IsStatic)
                {
                    if(!attribute.IsFactory)
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
            // #PropertyConversion Register Get* and Bind* methods for properties marked with IExcelFunction
            foreach (var property in _properties.Where(p => p.DeclaringType.GetInterfaces().Contains(typeof(IBindable))))
            {
                var getLambda = GetProperty(property);
                var attribute = (IExcelFunctionAttribute)property.GetCustomAttributes(typeof(IExcelFunctionAttribute)).Single();
                var handleParameter = new ExcelParameterRegistration(new ExcelArgumentAttribute() { Name = "Handle" });

                // we do not register twice the base properties
                if (property.ReflectedType != property.DeclaringType)
                    continue;

                var className = property.ReflectedType.Name.Replace("Public","");

                var getRegistration = new ExcelFunctionRegistration(getLambda,
                    //attribute.ToExcelFunctionAttribute(property.DeclaringType.BaseType.GenericTypeArguments.First().Name + "." + "Get" +property.Name),
                    attribute.ToExcelFunctionAttribute(className + "." + "Get" + property.Name),
                    new List<ExcelParameterRegistration>() { handleParameter });
                registrationList.Add(getRegistration);

                var bindLambda = BindProperty(property);
                var bindRegistration = new ExcelFunctionRegistration(bindLambda,
                    attribute.ToExcelFunctionAttribute(className + "." + "Bind" + property.Name),
                    new List<ExcelParameterRegistration>() { handleParameter });
                registrationList.Add(bindRegistration);

            }
            return registrationList;
        }

    }
}
