using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelInterfaces;
using Syncfusion.Calculate;

namespace SFAddin
{
    public class LibraryFunctionEx
    {
        public CalcEngine.LibraryFunction LibraryFunction;
        public string Name;
    }

    public class PublicRegistration
    { 
        private static void LoadReferences()
        {
            var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies().ToList();
            var loadedPaths = loadedAssemblies.Where(a => !a.IsDynamic).Select(a => a.Location).Distinct().ToArray();

            var referencedPaths = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll");
            var toLoad = referencedPaths.Where(r => !loadedPaths.Contains(r, StringComparer.InvariantCultureIgnoreCase)).ToList();
            toLoad.ForEach(path => loadedAssemblies.Add(AppDomain.CurrentDomain.Load(AssemblyName.GetAssemblyName(path))));
        }

        private static IEnumerable<MethodInfo> FindAllMethods()
        {
            LoadReferences();

            var allMethods = new List<MethodInfo>();
            // load the Public assemblies
            foreach (var theAssembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (theAssembly.GetCustomAttributes(typeof(PublicAttribute), false).Length == 0)
                    continue;

                allMethods.AddRange(theAssembly.GetTypes()
                    .SelectMany(t => t.GetMethods())
                    .Where(m => m.GetCustomAttributes(typeof(IExcelFunctionAttribute), false).Length > 0));
            }

            return allMethods;
        }

        private static dynamic CreateArray(dynamic input, Type t)
        {
            if (input is Array)
            {
                dynamic arrayArg = Array.CreateInstance(t, input.Count);
                for (var i = 0; i < input.Count; ++i)
                {
                    arrayArg[i] = (dynamic) Convert.ChangeType(input[i], t);
                }
                return (dynamic) arrayArg;
            }
            else
            {
                dynamic arrayArg = Array.CreateInstance(t, 1);
                arrayArg[0] = (dynamic)Convert.ChangeType(input, t);
                return (dynamic) arrayArg;
            }
        }

        private static CalcEngine.LibraryFunction WrapMethod(CalcEngine calcEngine,MethodInfo method)
        {
            if (method.DeclaringType == null)
                return null;

            var parameters = method.GetParameters().ToList();
            var maxParameterCount = parameters.Count;
            var nonOptionalParameterCount = parameters.Count(p => !p.IsOptional);

            Func<string, string> libFunc = args =>
            {
                var allArgs = calcEngine.SplitArgsPreservingQuotedCommas(args).ToList();
                var ret = "Error";
                if (allArgs.Count < nonOptionalParameterCount || allArgs.Count > maxParameterCount)
                    return ret;

                var parsedArgs =
                    allArgs.Select(arg => calcEngine.ComputeIsRef(arg) == "TRUE" ? calcEngine.GetValueFromArg(arg) : arg).ToList();

                var finalArgs = new List<object>();
                parameters.ForEach(x => finalArgs.Add(Type.Missing) );

                for (var i = 0; i < parsedArgs.Count; ++i)
                {
                    try
                    {
                        var paramType = parameters[i].ParameterType;
                        if (paramType.IsArray)
                        {
                            finalArgs[i] = (dynamic) CreateArray(parsedArgs[i], paramType.GetElementType());
                        }
                        else
                        {
                            finalArgs[i] = Convert.ChangeType(parsedArgs[i], paramType);
                        }
                    }
                    catch (Exception ex)
                    {
                        return ex.Message;
                    }
                }

                try
                {
                    ret = (string)method.Invoke(null, finalArgs.ToArray());
                }
                catch (Exception ex)
                {
                    return ex.Message; 
                }

                return ret;
            };

            return libFunc.Invoke;
        }

        public static IEnumerable<LibraryFunctionEx> WrapAllMethods(CalcEngine calcEngine)
        {
            return FindAllMethods().Select(m => new LibraryFunctionEx() { Name = m.Name, LibraryFunction = WrapMethod(calcEngine,m) } );
        } 
    }
}
