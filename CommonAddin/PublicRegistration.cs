using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;

using ExcelInterfaces;

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

    public static class PublicRegistration
    {
        private static IEnumerable<MethodInfo> FindAllMethods()
        {
            var allMethods = new List<MethodInfo>();
            // load the Public assemblies
            foreach (var theAssembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (theAssembly.GetCustomAttributes(typeof (PublicAttribute), false).Length == 0)
                    continue; 

                allMethods.AddRange(theAssembly.GetTypes()
                    .SelectMany(t => t.GetMethods())
                    .Where(m => m.GetCustomAttributes(typeof(IExcelFunctionAttribute), false).Length > 0));
            }

            return allMethods;
        }

        private static LambdaExpression WrapMethod(MethodInfo method)
        {
            if (method.DeclaringType == null)
                return null;

            var callParams = method.GetParameters().Select(p => Expression.Parameter(p.ParameterType, p.Name)).ToList();
            var callExpr = Expression.Call(method, callParams);

            return Expression.Lambda(callExpr, method.Name, callParams);
        }

        // register all functions with IExcel attributes
        public static IEnumerable<ExcelFunctionRegistration> GetAllRegistrations()
        {
            var registrationList = new List<ExcelFunctionRegistration>();
            foreach (var methodInfo in FindAllMethods())
            {
                var lambda = WrapMethod(methodInfo);
                var attribute = (IExcelFunctionAttribute) methodInfo.GetCustomAttributes(typeof (IExcelFunctionAttribute)).Single();
                registrationList.Add(new ExcelFunctionRegistration(lambda, attribute.ToExcelFunctionAttribute(methodInfo.Name),methodInfo.GetParameters().Select(p => new ExcelParameterRegistration(p))));
            }
            return registrationList;
        }
    }
}
