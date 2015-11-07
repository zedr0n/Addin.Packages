using CommonAddin;
using ExcelDna.Integration;

namespace $rootnamespace$
{
    public class Test
    {
        [ExcelFunction]
        public static string TestExcel()
        {
            return "Test completed";
        }
    }

    public class $rootnamespace$ : ExcelAddin
    {
    }
}
