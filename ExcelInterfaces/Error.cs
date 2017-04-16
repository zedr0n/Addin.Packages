using System;

namespace ExcelInterfaces
{
    public class Error : Exception
    {
        public Error(string message) :
            base("#Err: " + message)
        {
        }
    }

    public class PropertyMissing : Error
    {
        public PropertyMissing(string paramName) :
            base("Parameter missing : " + paramName)
        {
        }
    }
    public class ObjectMissing : Error
    {
        public ObjectMissing(string handle) :
            base(handle.Contains("#Err") ? "Object missing" : "Object missing : " + handle)
        {
        }
    }
}