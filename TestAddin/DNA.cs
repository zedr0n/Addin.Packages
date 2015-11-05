using System.IO.Pipes;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace TestAddin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("DNA")]

    public class Dna
    {
        public Ticker LoadTicker(string name)
        {
            return Ticker.LoadFromDb(name);
        }
    }

    [ComVisible(false)]
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

}
