using System;

namespace ExcelInterfaces
{
    [AttributeUsage(AttributeTargets.Assembly,Inherited = false, AllowMultiple = false)]
    public class PublicAttribute : Attribute
    {
        
    }

    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class IExcelFunctionAttribute : Attribute
    {
        public string Name = null;
        public string Description = null;
        public string Category = null;
        public string HelpTopic = null;
        public bool IsVolatile = false;
        public bool IsHidden = false;
        public bool IsExceptionSafe = false;
        public bool IsMacroType = false;
        public bool IsThreadSafe = false;
        public bool IsClusterSafe = false;
        public bool ExplicitRegistration = false;

        public IExcelFunctionAttribute()
        {
        }
    }

    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public class IExcelArgumentAttribute : Attribute
    {
        public string Name = null;
        public string Description = null;

        public IExcelArgumentAttribute()
        {
        }

        public IExcelArgumentAttribute(string description)
        {
            Description = description;
        }
    }
}
