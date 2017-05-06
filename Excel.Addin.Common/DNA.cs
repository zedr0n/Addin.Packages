using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;
using SimpleInjector;
using SimpleInjector.Advanced;
using Application = Microsoft.Office.Interop.Excel.Application;
using Error = ExcelInterfaces.Error;

namespace CommonAddin
{
    public static class Extensions
    {
        public static bool IsDefault<T>(this T obj)
        {
            return EqualityComparer<T>.Default.Equals(obj, default(T));
        }

        //public static ExcelFunctionAttribute()
    }

    public class ExcelAddin : IExcelAddIn
    {
        public Container Container { get; set; }
        public List<MethodInfo> Methods { get; set; } = new List<MethodInfo>();
        public List<PropertyInfo> Properties { get; set; } = new List<PropertyInfo>();


        /// <summary>
        /// Convert Error exceptions to string messages to be displayed in UDF cell
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private object OnError(object ex)
        {
            var errorMessage = ex as Error;
            if (errorMessage == null)
                return ExcelError.ExcelErrorValue;

            return errorMessage.Message;
        }

        public virtual void AutoOpen()
        { 
            ExcelIntegration.RegisterUnhandledExceptionHandler(OnError);

            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig();

            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit registrations are the only ones - there is no default processing
            ExcelRegistration.GetExcelFunctions()
                             .ProcessParameterConversions(conversionConfig)
                             .ProcessParamsRegistrations()
                             .RegisterFunctions();

            var registration = new Registration(Container,Methods,Properties);
            var bindingService = Container.GetInstance<IBindingService>();
            var application = (Application)ExcelDnaUtil.Application;
            application.SheetChange += bindingService.OnSheetChange;

            registration.GetAllRegistrations()
                .ProcessParameterConversions(conversionConfig)
                .ProcessParamsRegistrations()
                .ProcessAsyncRegistrations()
                .RegisterFunctions();
        }
        public void AutoClose()
        {
        }

        /// <summary>
        /// Convert excel range parameter to arrays and null objects to NA
        /// </summary>
        /// <returns></returns>
        private ParameterConversionConfiguration GetParameterConversionConfig()
        {
            var creator = Container.GetInstance<ICreator>();

            var paramConversionConfig = new ParameterConversionConfiguration()

                // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToInt32).ToArray())
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())
                // #ParameterConversion Convert handle to public object
                .AddParameterConversion((object obj) => creator.Get((string) obj))
                //.AddParameterConversion((string handle) => handle.Contains("::") ? Container.GetInstance<ICreator>().Create(handle) : handle )
                //.AddParameterConversion((Type type, ExcelParameterRegistration paramReg) =>
                //    (Expression<Func<object, IPublicObject>>)(obj => creator.Create((string)obj)), typeof(IPublicObject))
                //paramReg.ArgumentAttribute.Name == "oTransaction" ? (Expression<Func<object, IPublicObject>>)(obj => creator.Create((string)obj)) : null, typeof(IPublicObject))
                //(Expression<Func<object,IPublicObject>>) (obj => paramReg.ArgumentAttribute.Name == "hTransaction" ? creator.Create((string) obj) : null),typeof(IPublicObject))

                .AddReturnConversion((object obj) => obj.IsDefault() ? ExcelError.ExcelErrorNA : obj)
                // #ReturnConversion Convert public objects to its handle for excel display
                .AddReturnConversion((IPublicObject obj) => obj.Handle );

            return paramConversionConfig;
        }

    }

}
