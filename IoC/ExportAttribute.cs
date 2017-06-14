using System;

namespace IoC
{
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Constructor)]
    public class ExportAttribute : Attribute
    {
        public string Name;
        public string Description;
        public bool IsFactory;
        public bool IsMacroType;

        public ExportAttribute(string name, string description, bool isFactory =  false)
        {
            Name = name;
            Description = description;
            IsFactory = isFactory;
        }

        public ExportAttribute() { }
        public ExportAttribute(string name, bool isFactory = false)
        {
            Name = name;
            IsFactory = isFactory;
        }
    }
}