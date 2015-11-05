using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace TestAddin
{
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => ExcelError.ExcelErrorValue);

            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig();

            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit registrations are the only ones - there is no default processing
            ExcelRegistration.GetExcelFunctions()
                             .ProcessParameterConversions(conversionConfig)
                             .ProcessParamsRegistrations()
                             .RegisterFunctions();
        }
        public void AutoClose()
        {
        }

        private static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            // NOTE: The parameter conversion list is processed once per parameter.
            //       Parameter conversions will apply from most inside, to most outside.
            //       So to apply a conversion chain like
            //           string -> Type1 -> Type2
            //       we need to register in the (reverse) order
            //           Type1 -> Type2
            //           string -> Type1
            //
            //       (If the registration were in the order
            //           string -> Type1
            //           Type1 -> Type2
            //       the parameter (starting as Type2) would not match the first conversion,
            //       then the second conversion (Type1 -> Type2) would be applied, and no more,
            //       leaving the parameter having Type1 (and probably not eligible for Excel registration.)
            //      
            //
            //       Return conversions are also applied from most inside to most outside.
            //       So to apply a return conversion chain like
            //           Type1 -> Type2 -> string
            //       we need to register the ReturnConversions as
            //           Type1 -> Type2 
            //           Type2 -> string
            //       

            var paramConversionConfig = new ParameterConversionConfiguration()

                // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))

                //  .AddParameterConversion((string value) => convert2(convert1(value)));

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())

                .AddReturnConversion((object obj) => obj.IsDefault() ? ExcelError.ExcelErrorNA : obj);

            return paramConversionConfig;
        }

    }

}
