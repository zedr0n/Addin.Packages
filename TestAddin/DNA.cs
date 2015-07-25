using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;

namespace TestAddin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("DNA")]

    public class DNA
    {
        public Ticker Ticker_readFromXML(string xml)
        {
            return Ticker.readFromXML(xml);
        }
    }
}
