using System;
using IoC;

namespace ExcelInterfaces
{
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class IExcelFunctionAttribute : ExportAttribute
    {
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
