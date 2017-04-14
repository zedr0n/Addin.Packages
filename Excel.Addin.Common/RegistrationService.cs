using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;

namespace CommonAddin
{
    public class RegistrationService : IRegistrationService
    {
        private readonly Dictionary<string,string> _buttonHandles = new Dictionary<string, string>();

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

        public string GetAssociatedHandle()
        {
            var reference = XlCall.Excel(XlCall.xlfCaller) as string;
            return reference != null ? _buttonHandles[reference] : null;
        }
    }
}