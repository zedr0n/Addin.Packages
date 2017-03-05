using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelInterfaces;
using SimpleInjector;

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
        public IEnumerable<MethodInfo> Methods { get; set; }

        public virtual void AutoOpen()
        { 
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex =>
            {
                var errorMessage = ex as Error;
                if (errorMessage == null)
                    return ExcelError.ExcelErrorValue;

                return errorMessage.Message;
            });

            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig();

            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit registrations are the only ones - there is no default processing
            ExcelRegistration.GetExcelFunctions()
                             .ProcessParameterConversions(conversionConfig)
                             .ProcessParamsRegistrations()
                             .RegisterFunctions();

            var registration = new Registration(Container,Methods);

            registration.GetAllRegistrations()
                .ProcessParameterConversions(conversionConfig)
                .ProcessParamsRegistrations()
                .RegisterFunctions();
        }
        public void AutoClose()
        {
        }

        private static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            var paramConversionConfig = new ParameterConversionConfiguration()

                // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToInt32).ToArray())
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())

                .AddReturnConversion((object obj) => obj.IsDefault() ? ExcelError.ExcelErrorNA : obj);

            return paramConversionConfig;
        }

    }

}
