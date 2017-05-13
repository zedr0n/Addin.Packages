using ExcelDna.Integration;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;

namespace CommonAddin
{
    public class StatusService : IStatusService
    {
        public void Set(string status)
        {
            var application = (Application)ExcelDnaUtil.Application;
            application.StatusBar = status;
        }

        public void Clear()
        {
            Set("Ready");
        }
    }
}