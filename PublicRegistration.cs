using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
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
                    arrayArg[i] = Convert.ChangeType(input[i], t);
                    RemoveCharacter(arrayArg[i], "\"");
                }
                return arrayArg;
            }
            else
            {
                RemoveCharacter(input, "\"");
                dynamic arrayArg = Array.CreateInstance(t, 1);
                arrayArg[0] = Convert.ChangeType(input, t);

                return arrayArg;
            }
        }

        private static object RemoveCharacter(object input,string character)
        {
            var str = input as string;
            return str?.Replace(character, string.Empty) ?? input;
        }

        private static List<string> GetArguments(CalcEngine calcEngine, string args)
        {
            var allArgs = calcEngine.SplitArgsPreservingQuotedCommas(args).ToList();

            for (var i = 0; i < allArgs.Count; ++i)
            {
                var arg = allArgs[i];
                calcEngine.AdjustRangeArg(ref arg);
                allArgs[i] = arg;
            }

            return allArgs;
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
                var allArgs = GetArguments(calcEngine, args);
                var ret = "Error";
                if (allArgs.Count < nonOptionalParameterCount || allArgs.Count > maxParameterCount)
                    return ret;

                var finalArgs = new List<object>();
                parameters.ForEach(x => finalArgs.Add(Type.Missing) );

                for (var i = 0; i < allArgs.Count; ++i)
                {
                    try
                    {
                        var paramType = parameters[i].ParameterType;
                        if (paramType.IsArray)
                        {
                            finalArgs[i] = CreateArray(allArgs[i], paramType.GetElementType());
                        }
                        else
                        {
                            finalArgs[i] = RemoveCharacter(Convert.ChangeType(allArgs[i], paramType),"\"");
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

        public static void RegisterAllMethods(CalcEngine calcEngine)
        {
            WrapAllMethods(calcEngine).ToList().ForEach(x => calcEngine.AddFunction(x.Name,x.LibraryFunction));
        }
    }
}
