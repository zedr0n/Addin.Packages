using System;

namespace IoC
{
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    public class ExportAttribute : Attribute
    {
        public string Name;
        public string Description;

        public ExportAttribute(string name, string description)
        {
            Name = name;
            Description = description;
        }

        public ExportAttribute() { }
    }
}