using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;

namespace CommonAddin
{
    public class RegistrationService : IRegistrationService
    {
        private readonly Dictionary<string,string> _buttonHandles = new Dictionary<string, string>();

        public IStatusService StatusService { get; set; }

        public bool RegisterButton(string buttonName, string functionName, string handle)
        {
            var application = (Application)ExcelDnaUtil.Application;
            var worksheet = application.ActiveSheet as Worksheet;
            var button = worksheet.Buttons(buttonName) as Button;
            button.OnAction = functionName;

            if (_buttonHandles.ContainsKey(buttonName) && _buttonHandles[buttonName] == handle)
                return false;
            _buttonHandles[buttonName] = handle;
            return true;
        }

        /// <summary>
        ///     Get the handle of the object used for registration for the caller object
        /// </summary>
        /// <returns>Handle of the object associated with the form object</returns>
        // #RegistrationService GetAssociatedHandle
        public string GetAssociatedHandle()
        {
            var reference = XlCall.Excel(XlCall.xlfCaller) as string;
            return reference != null ? _buttonHandles[reference] : null;
        }

        public RegistrationService(IStatusService statusService)
        {
            StatusService = statusService;
        }
    }
}