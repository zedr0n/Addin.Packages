using System;

namespace ExcelInterfaces
{
    [AttributeUsage(AttributeTargets.Assembly,Inherited = false, AllowMultiple = false)]
    public class PublicAttribute : Attribute
    {
        
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class IExcelFunctionAttribute : Attribute
    {
        public string Name = null;
        public string Description = null;
        public readonly string Category = null;
        public readonly string HelpTopic = null;
        public bool IsVolatile = false;
        public bool IsHidden = false;
        public bool IsExceptionSafe = false;
        public bool IsMacroType = false;
        public bool IsThreadSafe = false;
        public bool IsClusterSafe = false;
        public bool ExplicitRegistration = false;
        public bool IsFactory = false;

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
