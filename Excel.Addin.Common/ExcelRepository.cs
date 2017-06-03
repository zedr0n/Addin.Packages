using System;
using System.Collections.Generic;
using System.Linq;
using ExcelInterfaces;

namespace Excel.Addin.Common
{
    public class ExcelRepository : IExcelRepository
    {
        private readonly Dictionary<string, object> _dictionary = new Dictionary<string, object>();

        private readonly IRegistrationService _registrationService;

        public ExcelRepository(IRegistrationService registrationService)
        {
            _registrationService = registrationService;
        }

        public object GetByHandle(string handle)
        {
            if (handle == "")
                // #RegistrationService get the handle associated with the button
                handle = _registrationService.GetAssociatedHandle();

            if (!_dictionary.ContainsKey(handle))
                throw new ObjectMissing(handle);

            return _dictionary[handle];
        }

        public void Add(object obj, string handleName)
        {
            var handle = TimestampHandle(handleName + "::" + obj.GetType().Name);
            if(_dictionary.ContainsKey(handle))
                throw new Error("Object with this handle already exists");

            _dictionary[handle] = obj;
        }

        public string ResolveHandle(object instance)
        {
            foreach (var pair in _dictionary)
                if (pair.Value == instance)
                    return pair.Key;

            throw new Error("Object has not been added to the repository");
        }

        private static string TimestampHandle(string handle)
        {
            return handle + "::" + DateTime.Now.ToString("hh:mm:ss.ff");
        }

        public static string StripHandle(string handle)
        {
            return handle.Split(new char[] { ':' }).FirstOrDefault();
        }
    }
}