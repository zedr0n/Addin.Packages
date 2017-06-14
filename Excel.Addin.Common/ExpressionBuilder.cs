using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelInterfaces;
using IoC;

namespace Excel.Addin.Common
{
    public class ExpressionBuilder
    {
        private readonly IContainerService _containerService;
        private readonly IExcelRepository _repository;

        public ExpressionBuilder(IContainerService containerService, IExcelRepository repository)
        {
            _containerService = containerService;
            _repository = repository;
        }

        public LambdaExpression BuildMethodExpression(MethodInfo methodInfo)
        {
            var instanceType = methodInfo.DeclaringType;

            if (typeof(IFactory).IsAssignableFrom(instanceType) || methodInfo.GetCustomAttribute<ExportAttribute>().IsFactory)
                return BuildFactoryExpression(methodInfo);
            else if (methodInfo.IsStatic)
                return BuildStaticExpression(methodInfo);
            else
                return BuildInstanceExpression(methodInfo);
        }



        private void ProcessParameters(ParameterInfo[] parameterInfos, out List<ParameterExpression> parameters, out List<Expression> resolvedParameters)
        {
            Expression<Func<string, object>> getExpression = handle => _repository.GetByHandle(handle);
            parameters = new List<ParameterExpression>();
            resolvedParameters = new List<Expression>();

            foreach (var p in parameterInfos)
            {
                var param = Expression.Parameter(p.ParameterType, p.Name);
                if (p.ParameterType.IsPrimitive || p.ParameterType == typeof(string))
                {
                    parameters.Add(param);
                    resolvedParameters.Add(param);
                }
                else // otherwise this is not handled by ExcelDna
                {
                    var handleParam = Expression.Parameter(typeof(string), p.Name);
                    parameters.Add(handleParam);

                    var blockExpression = Expression.Block(
                        Expression.Convert(Expression.Invoke(getExpression, handleParam), p.ParameterType)  // return _repository.GetByHandle(handle);
                    );

                    resolvedParameters.Add(blockExpression);
                }
            }
        }


        /// <summary>
        /// updates the parameter in the expression
        /// </summary>
        class ParameterUpdateVisitor : ExpressionVisitor
        {
            private readonly ParameterExpression _oldParameter;
            private readonly Expression _newParameter;

            public ParameterUpdateVisitor(ParameterExpression oldParameter, Expression newParameter)
            {
                _oldParameter = oldParameter;
                _newParameter = newParameter;
            }

            protected override Expression VisitParameter(ParameterExpression node)
            {
                if (object.ReferenceEquals(node, _oldParameter))
                    return _newParameter;

                return base.VisitParameter(node);
            }
        }

        private LambdaExpression BuildFactoryExpression(MethodInfo methodInfo)
        {
            var instanceType = methodInfo.DeclaringType;

            if (instanceType == null)
                throw new ArgumentException("Method is invalid");

            var parameters = methodInfo.GetParameters();
            //var parameterExpressions = new List<ParameterExpression>();
            //var resolvedParameterExpressions = new List<Expression>();
            //var parameterExpressions = parameters.Select(p => Expression.Parameter(p.ParameterType,p.Name)).ToList();
            ProcessParameters(parameters,out var parameterExpressions, out var resolvedParameterExpressions);
            var instanceParam = Expression.Parameter(instanceType);
            var objectParam = Expression.Parameter(methodInfo.ReturnType);
            var handleNameParam = Expression.Parameter(typeof(string),"Handle");

            Expression<Func<object>> instanceExpression = () => _containerService.GetInstance(instanceType); 
            Expression<Action<object, string>> addExpression = (instance, handleName) => _repository.Add(instance,handleName);
            Expression<Func<object, string>> resolveExpression = obj => _repository.ResolveHandle(obj);
            //Expression<Func<string, object>> getExpression = handle => _repository.GetByHandle(handle);

            var blockExpression = Expression.Block(
                new[] { instanceParam, objectParam },   
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(instanceExpression),instanceType)),                        // var instance = () => _containerService.GetInstance(instanceType)
                Expression.Assign(objectParam,Expression.Convert(Expression.Call(instanceParam,methodInfo,resolvedParameterExpressions),methodInfo.ReturnType)),  // var obj = instance.Create(...)
                Expression.Invoke(addExpression,objectParam,handleNameParam),                                   // _repository.Add(obj, handleName)
                Expression.Invoke(resolveExpression,objectParam)                                                // return _repository.ResolveHandle(obj)
            );

            var allParameterExpressions = new List<ParameterExpression>(parameterExpressions);
            allParameterExpressions.Insert(0, handleNameParam);

            var lambdaExpression = Expression.Lambda(blockExpression, allParameterExpressions);

            return lambdaExpression;
        }

        /// <summary>
        /// Wraps the property into a RTD getter lambda to associate with excel UDF
        /// </summary>
        /// <returns></returns>
        public LambdaExpression BuildPropertyExpression(PropertyInfo propertyInfo)
        {
            var instanceParam = Expression.Parameter(propertyInfo.DeclaringType, "instance");
            var handleParam = Expression.Parameter(typeof(string), "handle");
            var method = typeof(BindExtensions).GetMethod(nameof(BindExtensions.GetProperty));

            Expression<Func<string, object>> instanceExpression = h => _repository.GetByHandle(h);

            var block = Expression.Block(
                new[] { instanceParam },
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(instanceExpression, handleParam), propertyInfo.DeclaringType)),
                Expression.Call(method, instanceParam, Expression.Constant(propertyInfo.Name))
            );

            var allArguments = new List<ParameterExpression> { handleParam };

            var callExpr = Expression.Lambda(block, allArguments);
            return callExpr;
        }

        private LambdaExpression BuildInstanceExpression(MethodInfo methodInfo)
        {
            var instanceType = methodInfo.DeclaringType;
            if (instanceType == null)
                throw new ArgumentException("Method is invalid");

            var parameters = methodInfo.GetParameters();
            var parameterExpressions = parameters.Select(p => Expression.Parameter(p.ParameterType,p.Name)).ToList();

            var instanceParam = Expression.Parameter(instanceType);
            var handleParam = Expression.Parameter(typeof(string),"Handle");
            var returnParam = Expression.Parameter(methodInfo.ReturnType);

            Expression<Func<string,object>> instanceExpression = (h) => _repository.GetByHandle(h);
            Expression<Func<object, string>> returnExpression = o => _repository.ResolveHandle(o);

            var isPrimitive = methodInfo.ReturnType.IsPrimitive || methodInfo.ReturnType == typeof(string);

            var invokeBlock = Expression.Block(
                // var instance = _repository.GetByHandle(handle)
                // var o = instance.Invoke(...)
                // return o
                new[] { instanceParam },
                Expression.Assign(instanceParam, Expression.Convert(Expression.Invoke(instanceExpression, handleParam), methodInfo.DeclaringType)),
                Expression.Call(instanceParam, methodInfo, parameterExpressions)     
            );
            var blockExpression = invokeBlock;
            if(!isPrimitive)
                blockExpression = Expression.Block(
                    // return isPrimitive ? o : _repository.ResolveHandle(o)
                    new[] { returnParam },
                    Expression.Assign(returnParam,invokeBlock),
                    Expression.Invoke(returnExpression,returnParam)
                );


            var allParameterExpressions = new List<ParameterExpression>(parameterExpressions);
            allParameterExpressions.Insert(0, handleParam);

            var lambdaExpression = Expression.Lambda(blockExpression, allParameterExpressions);

            return lambdaExpression;
        }
        private LambdaExpression BuildStaticExpression(MethodInfo methodInfo)
        {
            if (!methodInfo.IsStatic)
                throw new ArgumentException("Method is invalid");
            var parameters = methodInfo.GetParameters();
            var parameterExpressions = parameters.Select(p => Expression.Parameter(p.ParameterType,p.Name)).ToList();

            var methodCallExpression = Expression.Call(null, methodInfo, parameterExpressions);

            var lambdaExpression = Expression.Lambda(methodCallExpression, parameterExpressions);       // return Invoke()
            return lambdaExpression;
        }
    }
}